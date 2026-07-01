#Requires -Version 5.1
# Unit tests for src/Scan.ps1 pure helpers + the real directory walk (no admin needed).

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly
}

Describe 'Get-DataFileType' {
    It 'classifies <path> as <expected>' -ForEach @(
        @{ path = 'C:\x\a.pst'; expected = 'PST' }
        @{ path = 'C:\x\a.PST'; expected = 'PST' }
        @{ path = 'C:\x\a.ost'; expected = 'OST' }
        @{ path = 'a.OsT';      expected = 'OST' }
    ) {
        Get-DataFileType $path | Should -Be $expected
    }
    It 'returns $null for non-data extensions' {
        Get-DataFileType 'C:\x\a.txt' | Should -BeNullOrEmpty
        Get-DataFileType 'C:\x\a'     | Should -BeNullOrEmpty
    }
}

Describe 'Format-FileSize' {
    It 'formats <bytes> as <expected>' -ForEach @(
        @{ bytes = 0;        expected = '0 B' }
        @{ bytes = 512;      expected = '512 B' }
        @{ bytes = 1024;     expected = '1.0 KB' }
        @{ bytes = 1536;     expected = '1.5 KB' }
        @{ bytes = 1048576;  expected = '1.0 MB' }
        @{ bytes = 1073741824; expected = '1.0 GB' }
    ) {
        Format-FileSize ([long]$bytes) | Should -Be $expected
    }
    It 'clamps negatives to 0 B' {
        Format-FileSize -1 | Should -Be '0 B'
    }
}

Describe 'Find-OutlookDataFile' {
    BeforeAll {
        $base = Join-Path $TestDrive 'tree'
        New-Item -ItemType Directory -Force -Path (Join-Path $base 'sub\deep') | Out-Null
        Set-Content -Path (Join-Path $base 'root.pst')      -Value 'x'
        Set-Content -Path (Join-Path $base 'sub\mid.ost')   -Value 'x'
        Set-Content -Path (Join-Path $base 'sub\deep\d.pst') -Value 'x'
        Set-Content -Path (Join-Path $base 'sub\note.txt')  -Value 'x'
    }

    It 'finds all pst/ost recursively and ignores other files' {
        $found = Find-OutlookDataFile -Roots @($base)
        $found.Count | Should -Be 3
        ($found.Type | Sort-Object) | Should -Be @('OST', 'PST', 'PST')
    }

    It 'returns records with Path, Type and Size' {
        $found = Find-OutlookDataFile -Roots @($base)
        $rec = $found | Where-Object { $_.Path -like '*root.pst' }
        $rec | Should -Not -BeNullOrEmpty
        $rec.Type | Should -Be 'PST'
        $rec.Size | Should -Match '\d'
    }

    It 'honors the cancel check' {
        $found = Find-OutlookDataFile -Roots @($base) -CancelCheck { $true }
        $found.Count | Should -Be 0
    }
}

Describe 'Test-FileLocked' {
    It 'returns $false for a file no one holds open' {
        $p = Join-Path $TestDrive 'free.pst'
        Set-Content -Path $p -Value 'x'
        Test-FileLocked $p | Should -BeFalse
    }
    It 'returns $true while another handle holds the file exclusively' {
        $p = Join-Path $TestDrive 'held.pst'
        Set-Content -Path $p -Value 'x'
        $fs = [System.IO.File]::Open($p, 'Open', 'Read', 'None')
        try { Test-FileLocked $p | Should -BeTrue }
        finally { $fs.Close(); $fs.Dispose() }
    }
}

Describe 'New-DataFileRecord' {
    It 'builds a full record for a real data file' {
        $p = Join-Path $TestDrive 'rec.pst'
        Set-Content -Path $p -Value 'hello'
        $rec = New-DataFileRecord $p
        $rec.Path          | Should -Be $p
        $rec.Type          | Should -Be 'PST'
        $rec.SizeBytes     | Should -BeGreaterThan 0
        $rec.Size          | Should -Match '\d'
        $rec.LastWriteTime | Should -Match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$'
        $rec.Locked        | Should -BeFalse
    }
    It 'falls back to SizeBytes -1 / null mtime when FileInfo cannot be read' {
        # An empty path makes the FileInfo constructor throw, exercising the catch so the
        # -1L / $null defaults survive. (A merely missing path does not throw on .NET Core.)
        $rec = New-DataFileRecord ''
        $rec.SizeBytes     | Should -Be -1
        $rec.LastWriteTime | Should -BeNullOrEmpty
    }
}

Describe 'Write-JsonFile' {
    It 'writes valid JSON and leaves no .tmp behind' {
        $p = Join-Path $TestDrive 'out.json'
        Write-JsonFile -Path $p -Object ([pscustomobject]@{ A = 1; B = 'x' })
        Test-Path "$p.tmp" | Should -BeFalse
        $back = Get-Content -Path $p -Raw | ConvertFrom-Json
        $back.A | Should -Be 1
        $back.B | Should -Be 'x'
    }
}

Describe 'Read-ScanProgress' {
    It 'returns the parsed object for a valid progress file' {
        $p = Join-Path $TestDrive 'prog.json'
        Write-JsonFile -Path $p -Object ([pscustomobject]@{ DirCount = 7; Done = $false })
        $r = Read-ScanProgress -Path $p
        $r.DirCount | Should -Be 7
        $r.Done     | Should -BeFalse
    }
    It 'returns $null when the file is missing' {
        Read-ScanProgress -Path (Join-Path $TestDrive 'nope.json') | Should -BeNullOrEmpty
    }
    It 'returns $null on malformed JSON' {
        $p = Join-Path $TestDrive 'bad.json'
        Set-Content -Path $p -Value 'not json {'
        Read-ScanProgress -Path $p | Should -BeNullOrEmpty
    }
}

Describe 'Read-ScanResult' {
    It 'returns $null when the file is missing' {
        Read-ScanResult -Path (Join-Path $TestDrive 'missing.json') | Should -BeNullOrEmpty
    }
    It 'returns $null on malformed JSON' {
        $p = Join-Path $TestDrive 'badres.json'
        Set-Content -Path $p -Value '{ broken'
        Read-ScanResult -Path $p | Should -BeNullOrEmpty
    }
    It 'normalizes a single-file result back to a length-1 array' {
        # The crux: ConvertTo-Json collapses a 1-element array, so the reader must
        # re-wrap Files so callers always get a collection.
        $p = Join-Path $TestDrive 'one.json'
        Write-JsonFile -Path $p -Object ([pscustomobject]@{
            Cancelled = $false; Count = 1
            Files = @([pscustomobject]@{ Path = 'C:\a.pst'; Type = 'PST' })
        })
        $r = Read-ScanResult -Path $p
        $r.Cancelled    | Should -BeFalse
        $r.Count        | Should -Be 1
        @($r.Files).Count | Should -Be 1
        $r.Files[0].Type | Should -Be 'PST'
    }
    It 'keeps a multi-file result as an array' {
        $p = Join-Path $TestDrive 'many.json'
        Write-JsonFile -Path $p -Object ([pscustomobject]@{
            Cancelled = $true; Count = 2
            Files = @(
                [pscustomobject]@{ Path = 'C:\a.pst'; Type = 'PST' }
                [pscustomobject]@{ Path = 'C:\b.ost'; Type = 'OST' }
            )
        })
        $r = Read-ScanResult -Path $p
        $r.Cancelled      | Should -BeTrue
        @($r.Files).Count | Should -Be 2
    }
    It 'yields an empty Files array when the result has none' {
        $p = Join-Path $TestDrive 'empty.json'
        Set-Content -Path $p -Value '{ "Cancelled": false, "Count": 0 }'
        $r = Read-ScanResult -Path $p
        @($r.Files).Count | Should -Be 0
    }
}

Describe 'Stop-ElevatedScan' {
    It 'creates the cancel sentinel file' {
        $p = Join-Path $TestDrive 'run.cancel'
        Stop-ElevatedScan -CancelPath $p
        Test-Path $p | Should -BeTrue
    }
}

Describe 'Invoke-DiskScanHelper (integration)' {
    It 'walks the roots and writes a result plus a done-progress file' {
        $tree = Join-Path $TestDrive 'scan'
        New-Item -ItemType Directory -Force -Path (Join-Path $tree 'a') | Out-Null
        Set-Content -Path (Join-Path $tree 'one.pst')   -Value 'x'
        Set-Content -Path (Join-Path $tree 'a\two.ost') -Value 'x'
        Set-Content -Path (Join-Path $tree 'a\skip.txt') -Value 'x'
        $resPath  = Join-Path $TestDrive 'r.json'
        $progPath = Join-Path $TestDrive 'p.json'

        Invoke-DiskScanHelper -ResultPath $resPath -ProgressPath $progPath -Roots @($tree)

        $res = Read-ScanResult -Path $resPath
        $res.Count                       | Should -Be 2
        $res.Cancelled                   | Should -BeFalse
        ($res.Files.Type | Sort-Object)  | Should -Be @('OST', 'PST')
        (Read-ScanProgress -Path $progPath).Done | Should -BeTrue
    }
    It 'marks the result Cancelled when the sentinel exists up front' {
        $tree = Join-Path $TestDrive 'scan2'
        New-Item -ItemType Directory -Force -Path $tree | Out-Null
        Set-Content -Path (Join-Path $tree 'm.pst') -Value 'x'
        $resPath = Join-Path $TestDrive 'r2.json'
        New-Item -ItemType File -Path "$resPath.cancel" -Force | Out-Null

        Invoke-DiskScanHelper -ResultPath $resPath -Roots @($tree)

        (Read-ScanResult -Path $resPath).Cancelled | Should -BeTrue
    }
}

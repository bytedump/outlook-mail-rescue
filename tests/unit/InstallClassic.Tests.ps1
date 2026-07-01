#Requires -Version 5.1
# Unit tests for the pure ODT config builder in src/InstallClassic.ps1.

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly
}

Describe 'Format-OdtConfigXml' {
    It 'installs the requested product and bitness' {
        $xml = Format-OdtConfigXml -Bitness '64' -ProductId 'O365ProPlusRetail'
        $xml | Should -Match 'OfficeClientEdition="64"'
        $xml | Should -Match 'Product ID="O365ProPlusRetail"'
    }
    It 'excludes the other Office apps but NOT Outlook' {
        $xml = Format-OdtConfigXml
        $xml | Should -Match 'ExcludeApp ID="Word"'
        $xml | Should -Match 'ExcludeApp ID="Excel"'
        $xml | Should -Not -Match 'ExcludeApp ID="Outlook"'
    }
    It 'runs silently and accepts the EULA' {
        $xml = Format-OdtConfigXml
        $xml | Should -Match 'Display Level="None"'
        $xml | Should -Match 'AcceptEULA="TRUE"'
    }
    It 'honors 32-bit' {
        (Format-OdtConfigXml -Bitness '32') | Should -Match 'OfficeClientEdition="32"'
    }
    It 'adds a Logging element only when a log path is given' {
        (Format-OdtConfigXml -LogPath 'C:\Temp\odtlog') | Should -Match 'Logging Level="Standard" Path="C:\\Temp\\odtlog"'
        (Format-OdtConfigXml) | Should -Not -Match 'Logging'
    }
}

Describe 'Get-OdtInstallPhase' {
    It 'reports Preparing when there is no log yet' {
        Get-OdtInstallPhase -LogLines @() | Should -Be 'Preparing'
    }
    It 'detects the download stage' {
        Get-OdtInstallPhase -LogLines @('... Downloading the Office package ...') | Should -Be 'Downloading'
    }
    It 'detects the install stage' {
        Get-OdtInstallPhase -LogLines @('Apply stage: applying files') | Should -Be 'Installing'
    }
    It 'detects completion' {
        Get-OdtInstallPhase -LogLines @('Installation succeeded') | Should -Be 'Finishing'
    }
}

Describe 'New-OdtConfigXml' {
    It 'writes the config XML to disk and returns the path' {
        $p = Join-Path $TestDrive 'config.xml'
        $out = New-OdtConfigXml -Path $p -Bitness '64' -ProductId 'O365ProPlusRetail'
        $out | Should -Be $p
        Test-Path $p | Should -BeTrue
        (Get-Content -Path $p -Raw) | Should -Match 'OfficeClientEdition="64"'
    }
}

Describe 'Test-OdtAvailable' {
    It 'is true when the setup.exe exists' {
        $p = Join-Path $TestDrive 'setup.exe'
        Set-Content -Path $p -Value 'x'
        Test-OdtAvailable $p | Should -BeTrue
    }
    It 'is false for a missing path' {
        Test-OdtAvailable (Join-Path $TestDrive 'nope.exe') | Should -BeFalse
    }
    It 'is false for an empty path' {
        Test-OdtAvailable '' | Should -BeFalse
    }
}

Describe 'Read-InstallResult' {
    It 'parses a valid result file' {
        $p = Join-Path $TestDrive 'res.json'
        Write-JsonFile -Path $p -Object ([pscustomobject]@{ Success = $true; ExitCode = 0; Message = 'ok' })
        $r = Read-InstallResult -Path $p
        $r.Success | Should -BeTrue
        $r.Message | Should -Be 'ok'
    }
    It 'returns $null when the file is missing' {
        Read-InstallResult -Path (Join-Path $TestDrive 'gone.json') | Should -BeNullOrEmpty
    }
    It 'returns $null on malformed JSON' {
        $p = Join-Path $TestDrive 'bad.json'
        Set-Content -Path $p -Value '{ nope'
        Read-InstallResult -Path $p | Should -BeNullOrEmpty
    }
}

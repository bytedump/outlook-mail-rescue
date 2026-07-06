#Requires -Version 5.1
# Unit tests for the CI-able piece of src/Gui.ps1. The WinForms builders need a live
# desktop and stay manual e2e; Get-MailRescueConfig is pure filesystem and covered here.

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly
}

Describe 'Get-MailRescueConfig' {
    It 'loads config.ps1 from the given directory when present' {
        $dir = Join-Path $TestDrive 'cfg'
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        Set-Content -Path (Join-Path $dir 'config.ps1') -Value '@{ OutputFolder = "C:\Custom"; FileNameTemplate = "x.pst" }'
        $cfg = Get-MailRescueConfig -ScriptDir $dir
        $cfg.OutputFolder     | Should -Be 'C:\Custom'
        $cfg.FileNameTemplate | Should -Be 'x.pst'
    }

    It 'prefers config.ps1 over config.example.ps1' {
        $dir = Join-Path $TestDrive 'cfg2'
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        Set-Content -Path (Join-Path $dir 'config.example.ps1') -Value '@{ OutputFolder = "EXAMPLE" }'
        Set-Content -Path (Join-Path $dir 'config.ps1')         -Value '@{ OutputFolder = "REAL" }'
        (Get-MailRescueConfig -ScriptDir $dir).OutputFolder | Should -Be 'REAL'
    }

    It 'returns built-in defaults when no config file exists' {
        $dir = Join-Path $TestDrive 'empty'
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        $cfg = Get-MailRescueConfig -ScriptDir $dir
        $cfg.FileNameTemplate   | Should -Be '{username}_{stamp}.pst'
        $cfg.ScanRoots[0]       | Should -Be 'C:\'
        $cfg.SyncTimeoutMinutes | Should -Be 30
    }

    It 'falls back to defaults when the config file throws' {
        $dir = Join-Path $TestDrive 'broken'
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
        Set-Content -Path (Join-Path $dir 'config.ps1') -Value 'throw "boom"'
        Mock Write-Log {}
        (Get-MailRescueConfig -ScriptDir $dir).FileNameTemplate | Should -Be '{username}_{stamp}.pst'
    }
}

Describe 'Test-CanAutoFill' {
    It 'fills an empty/whitespace/null field' -ForEach @(
        @{ cur = '' }
        @{ cur = '   ' }
        @{ cur = $null }
    ) {
        Test-CanAutoFill -CurrentText $cur -LastAutoSet 'whatever' | Should -BeTrue
    }

    It 'does not throw on a null current text' {
        { Test-CanAutoFill -CurrentText $null -LastAutoSet 'x' } | Should -Not -Throw
    }

    It 'fills when the field still shows only the cue text' {
        Test-CanAutoFill -CurrentText 'Detecting...' -LastAutoSet '' -CueText 'Detecting...' | Should -BeTrue
    }

    It 'upgrades when the field still holds our last auto-set value' {
        # none -> displayname -> smtp: field untouched, so a better detection overwrites.
        Test-CanAutoFill -CurrentText 'John Smith' -LastAutoSet 'John Smith' | Should -BeTrue
    }

    It 'never overwrites a value the user typed/pasted' {
        Test-CanAutoFill -CurrentText 'someone.else@corp.com' -LastAutoSet 'John Smith' | Should -BeFalse
    }

    It 'treats a case change as a user edit (case-sensitive)' {
        Test-CanAutoFill -CurrentText 'A@B.com' -LastAutoSet 'a@b.com' | Should -BeFalse
    }

    It 'does not fill over user text when we never auto-set anything' {
        Test-CanAutoFill -CurrentText 'typed by hand' -LastAutoSet '' | Should -BeFalse
    }

    It 'does not treat real user text as the cue' {
        Test-CanAutoFill -CurrentText 'typed' -LastAutoSet '' -CueText 'Detecting...' | Should -BeFalse
    }
}

Describe 'Test-PathUnderOneDrive' {
    It 'flags a path under the OneDrive root' {
        Test-PathUnderOneDrive -Path 'C:\Users\user\OneDrive\Desktop\MailRescue\x.pst' -OneDriveRoots @('C:\Users\user\OneDrive') |
            Should -BeTrue
    }

    It 'passes a path that is NOT under any OneDrive root' {
        Test-PathUnderOneDrive -Path 'C:\Users\user\Desktop\MailRescue\x.pst' -OneDriveRoots @('C:\Users\user\OneDrive') |
            Should -BeFalse
    }

    It 'is boundary-safe: a lookalike sibling is not "under" the root' {
        Test-PathUnderOneDrive -Path 'C:\Users\user\OneDriveBackup\x.pst' -OneDriveRoots @('C:\Users\user\OneDrive') |
            Should -BeFalse
    }

    It 'matches case-insensitively (Windows paths)' {
        Test-PathUnderOneDrive -Path 'c:\users\USER\onedrive\desktop\x.pst' -OneDriveRoots @('C:\Users\user\OneDrive') |
            Should -BeTrue
    }

    It 'tolerates a trailing slash on the root' {
        Test-PathUnderOneDrive -Path 'C:\Users\user\OneDrive\x.pst' -OneDriveRoots @('C:\Users\user\OneDrive\') |
            Should -BeTrue
    }

    It 'normalizes forward slashes' {
        Test-PathUnderOneDrive -Path 'C:/Users/user/OneDrive/Desktop/x.pst' -OneDriveRoots @('C:\Users\user\OneDrive') |
            Should -BeTrue
    }

    It 'treats the root itself as "under" (writing at the root)' {
        Test-PathUnderOneDrive -Path 'C:\Users\user\OneDrive' -OneDriveRoots @('C:\Users\user\OneDrive') | Should -BeTrue
    }

    It 'matches any root when several are given' {
        $roots = @('C:\Users\user\OneDrive', 'C:\Users\user\OneDrive - Contoso')
        Test-PathUnderOneDrive -Path 'C:\Users\user\OneDrive - Contoso\Docs\x.pst' -OneDriveRoots $roots |
            Should -BeTrue
    }

    It 'returns false for empty roots or empty path' -ForEach @(
        @{ p = 'C:\Users\user\OneDrive\x.pst'; roots = @() }
        @{ p = 'C:\Users\user\OneDrive\x.pst'; roots = $null }
        @{ p = ''; roots = @('C:\Users\user\OneDrive') }
        @{ p = $null; roots = @('C:\Users\user\OneDrive') }
    ) {
        Test-PathUnderOneDrive -Path $p -OneDriveRoots $roots | Should -BeFalse
    }

    It 'skips null/empty entries in the roots list' {
        Test-PathUnderOneDrive -Path 'C:\Users\user\Desktop\x.pst' -OneDriveRoots @($null, '', 'C:\Users\user\OneDrive') |
            Should -BeFalse
    }
}

Describe 'Get-UniqueFileName' {
    It 'returns the name unchanged when nothing clashes' {
        Get-UniqueFileName -FileName 'backup.pst' -Existing @('other.pst') | Should -Be 'backup.pst'
    }

    It "appends ' (2)' on a single clash" {
        Get-UniqueFileName -FileName 'backup.pst' -Existing @('backup.pst') | Should -Be 'backup (2).pst'
    }

    It "bumps to ' (3)' when the name and ' (2)' are both taken" {
        Get-UniqueFileName -FileName 'name.pst' -Existing @('name.pst', 'name (2).pst') | Should -Be 'name (3).pst'
    }

    It 'clashes case-insensitively (Windows file names)' {
        Get-UniqueFileName -FileName 'backup.pst' -Existing @('BACKUP.PST') | Should -Be 'backup (2).pst'
    }

    It 'preserves a multi-dot base and only treats the last segment as the extension' {
        Get-UniqueFileName -FileName 'owner@company.com.br.05-07-2026.pst' -Existing @('owner@company.com.br.05-07-2026.pst') |
            Should -Be 'owner@company.com.br.05-07-2026 (2).pst'
    }

    It 'handles a name with no extension' {
        Get-UniqueFileName -FileName 'report' -Existing @('report') | Should -Be 'report (2)'
    }

    It 'returns the name unchanged for empty or null Existing' -ForEach @(
        @{ existing = @() }
        @{ existing = $null }
    ) {
        Get-UniqueFileName -FileName 'backup.pst' -Existing $existing | Should -Be 'backup.pst'
    }
}

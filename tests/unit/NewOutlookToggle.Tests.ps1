#Requires -Version 5.1
# Tests the registry backup/restore semantics against a throwaway HKCU key, so the
# revert behavior of the new->classic switch is verified for real (HKCU, no admin).

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly
    $script:TestKey = "HKCU:\Software\OutlookMailRescueTest_$([guid]::NewGuid().ToString('N'))"
}

AfterAll {
    if ($script:TestKey -and (Test-Path $script:TestKey)) {
        Remove-Item -Path $script:TestKey -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Describe 'Set-UseNewOutlook / Restore-RegistryValue' {
    It 'restores a pre-existing value' {
        New-Item -Path $script:TestKey -Force | Out-Null
        Set-ItemProperty -Path $script:TestKey -Name 'UseNewOutlook' -Value 1 -Type DWord

        $backup = Set-UseNewOutlook -Value 0 -Path $script:TestKey -Name 'UseNewOutlook'
        $backup.Existed | Should -BeTrue
        [int]$backup.OldValue | Should -Be 1
        (Get-RegistryValue $script:TestKey 'UseNewOutlook') | Should -Be 0

        Restore-RegistryValue -Backup $backup
        (Get-RegistryValue $script:TestKey 'UseNewOutlook') | Should -Be 1
    }

    It 'removes the value when it was absent before' {
        Remove-ItemProperty -Path $script:TestKey -Name 'UseNewOutlook' -ErrorAction SilentlyContinue

        $backup = Set-UseNewOutlook -Value 0 -Path $script:TestKey -Name 'UseNewOutlook'
        $backup.Existed | Should -BeFalse
        (Get-RegistryValue $script:TestKey 'UseNewOutlook') | Should -Be 0

        Restore-RegistryValue -Backup $backup
        (Get-RegistryValue $script:TestKey 'UseNewOutlook') | Should -BeNullOrEmpty
    }
}

Describe 'Get-UseNewOutlook' {
    It 'returns the value read from the UseNewOutlook preference' {
        Mock Get-RegistryValue { 0 } -ParameterFilter { $Name -eq 'UseNewOutlook' }
        Get-UseNewOutlook | Should -Be 0
        Should -Invoke Get-RegistryValue -Times 1 -Exactly
    }
    It 'returns $null when the preference is absent' {
        Mock Get-RegistryValue { $null }
        Get-UseNewOutlook | Should -BeNullOrEmpty
    }
}

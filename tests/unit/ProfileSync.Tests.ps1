#Requires -Version 5.1
# Unit tests for the pure stabilization helper in src/ProfileSync.ps1.

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly
}

Describe 'Test-SyncStabilized' {
    It 'is true when the last N readings are equal' {
        Test-SyncStabilized -Samples @(1, 5, 10, 10, 10) -StableReadings 3 | Should -BeTrue
    }
    It 'is false while the count is still growing' {
        Test-SyncStabilized -Samples @(1, 5, 10) -StableReadings 3 | Should -BeFalse
    }
    It 'is false before enough readings exist' {
        Test-SyncStabilized -Samples @(10, 10) -StableReadings 3 | Should -BeFalse
    }
    It 'treats a flat zero as stabilized (empty mailbox edge)' {
        Test-SyncStabilized -Samples @(0, 0, 0) -StableReadings 3 | Should -BeTrue
    }
    It 'handles null/empty input' {
        Test-SyncStabilized -Samples @() -StableReadings 3 | Should -BeFalse
    }
}

Describe 'Test-OutlookProfileExists' {
    It 'is true when Get-OutlookInfo reports a MAPI profile' {
        Mock Get-OutlookInfo { [pscustomobject]@{ HasMapiProfile = $true } }
        Test-OutlookProfileExists | Should -BeTrue
    }
    It 'is false when no MAPI profile exists' {
        Mock Get-OutlookInfo { [pscustomobject]@{ HasMapiProfile = $false } }
        Test-OutlookProfileExists | Should -BeFalse
    }
}

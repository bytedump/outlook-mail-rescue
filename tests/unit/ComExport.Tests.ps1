#Requires -Version 5.1
# Unit tests for the pure helpers in src/ComExport.ps1 (no Outlook/COM required).

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly

    function New-Store { param($Name, $Path, $Type)
        [pscustomobject]@{ DisplayName = $Name; FilePath = $Path; ExchangeStoreType = $Type }
    }
}

Describe 'Remove-InvalidFileNameChars' {
    It 'replaces invalid characters with underscore' {
        Remove-InvalidFileNameChars 'a:b*c?d' | Should -Be 'a_b_c_d'
    }
    It 'leaves valid names untouched' {
        Remove-InvalidFileNameChars 'john.doe_INC1.pst' | Should -Be 'john.doe_INC1.pst'
    }
}

Describe 'Format-PathToken' {
    It 'collapses whitespace to underscore' { Format-PathToken 'John  Doe' 'x' | Should -Be 'John_Doe' }
    It 'sanitizes invalid chars'            { Format-PathToken 'a/b' 'x'       | Should -Be 'a_b' }
    It 'uses fallback when empty'           { Format-PathToken '' 'fallback'   | Should -Be 'fallback' }
    It 'uses fallback when whitespace'      { Format-PathToken '   ' 'fb'      | Should -Be 'fb' }
}

Describe 'Get-ExportFileName' {
    It 'builds the default template' {
        Get-ExportFileName -Username 'john.doe' -Ticket 'INC123' -Stamp '20260628-120000' |
            Should -Be 'john.doe_INC123_20260628-120000.pst'
    }
    It 'falls back for empty username/ticket' {
        Get-ExportFileName -Username '' -Ticket '' -Stamp 'S' | Should -Be 'user_noticket_S.pst'
    }
    It 'sanitizes tokens with invalid chars' {
        Get-ExportFileName -Username 'John Doe' -Ticket 'INC/12:3' -Stamp 'S' |
            Should -Be 'John_Doe_INC_12_3_S.pst'
    }
}

Describe 'Get-IdentityPstFileName' {
    It 'names the file with the full primary SMTP, no stamp' {
        Get-IdentityPstFileName 'owner@company.com' |
            Should -BeExactly 'owner@company.com.pst'
    }

    It 'names a simple SMTP literally' {
        Get-IdentityPstFileName 'a.b@c.com' | Should -BeExactly 'a.b@c.com.pst'
    }

    It 'returns $null for no usable identity (caller must not auto-name)' -ForEach @(
        @{ in = '' }
        @{ in = '   ' }
        @{ in = '..' }
    ) {
        Get-IdentityPstFileName $in | Should -BeNullOrEmpty
    }

    It 'does not throw and returns $null on $null input' {
        { Get-IdentityPstFileName $null } | Should -Not -Throw
        Get-IdentityPstFileName $null | Should -BeNullOrEmpty
    }

    It 'neutralizes a reserved device name and still ends in .pst' {
        Get-IdentityPstFileName 'NUL' | Should -BeExactly '_NUL.pst'
    }

    It 'always ends in .pst when it returns a name' {
        Get-IdentityPstFileName 'someone@example.org' | Should -Match '(?i)\.pst$'
    }

    It 'passes MaxLength through to the token, then appends .pst' {
        $r = Get-IdentityPstFileName ('a' * 300) -MaxLength 10
        $r | Should -BeExactly ('a' * 10 + '.pst')
    }
}

Describe 'Get-DatedPstFileName' {
    It 'appends the day-month-year stamp before .pst' {
        Get-DatedPstFileName 'owner@company.com' -DateStamp '04-07-2026' |
            Should -BeExactly 'owner@company.com.04-07-2026.pst'
    }

    It 'returns $null for no usable identity (parity with Get-IdentityPstFileName)' -ForEach @(
        @{ in = '' }
        @{ in = '   ' }
        @{ in = '..' }
    ) {
        Get-DatedPstFileName $in -DateStamp '04-07-2026' | Should -BeNullOrEmpty
    }

    It 'does not throw and returns $null on $null input' {
        { Get-DatedPstFileName $null -DateStamp '04-07-2026' } | Should -Not -Throw
        Get-DatedPstFileName $null -DateStamp '04-07-2026' | Should -BeNullOrEmpty
    }

    It 'defaults to today (dd-MM-yyyy) when no stamp is passed' {
        Get-DatedPstFileName 'someone@example.org' |
            Should -Match '(?i)^someone@example\.org\.\d{2}-\d{2}-\d{4}\.pst$'
    }
}

Describe 'Get-ExportOutcome' {
    It 'is complete when every source item was copied' {
        $o = Get-ExportOutcome -SourceItems 27341 -CopiedItems 27341
        $o.Degraded     | Should -BeFalse
        $o.Title        | Should -Be 'Export complete'
        $o.Reasons.Count | Should -Be 0
    }

    It 'is complete for an empty mailbox (0 of 0)' {
        $o = Get-ExportOutcome -SourceItems 0 -CopiedItems 0
        $o.Degraded | Should -BeFalse
        $o.Title    | Should -Be 'Export complete'
    }

    It 'is incomplete when fewer items were copied than the source has' {
        $o = Get-ExportOutcome -SourceItems 27341 -CopiedItems 27000
        $o.Degraded      | Should -BeTrue
        $o.Title         | Should -Be 'Export incomplete'
        $o.Reasons.Count | Should -Be 1
        $o.Reasons[0]    | Should -Match 'missing'
    }

    It 'flags a count mismatch in the other direction (possible duplication)' {
        $o = Get-ExportOutcome -SourceItems 50 -CopiedItems 100
        $o.Degraded   | Should -BeTrue
        $o.Reasons[0] | Should -Match '(?i)duplicat'
    }

    It 'is incomplete when the mailbox sync stalled' {
        $o = Get-ExportOutcome -SourceItems 10 -CopiedItems 10 -SyncStalled $true
        $o.Degraded   | Should -BeTrue
        $o.Title      | Should -Be 'Export incomplete'
        $o.Reasons[0] | Should -Match '(?i)stall'
    }

    It 'is incomplete when cross-account copy was blocked' {
        $o = Get-ExportOutcome -SourceItems 10 -CopiedItems 10 -CrossAccountBlocked $true
        $o.Degraded   | Should -BeTrue
        $o.Reasons[0] | Should -Match '(?i)cross-account|policy'
    }

    It 'collects every reason when several conditions trip' {
        $o = Get-ExportOutcome -SourceItems 100 -CopiedItems 90 -SyncStalled $true -CrossAccountBlocked $true
        $o.Degraded      | Should -BeTrue
        $o.Reasons.Count | Should -Be 3
    }
}

Describe 'Resolve-OwnedProcessId' {
    It 'claims the single new PID that appeared after launch' {
        Resolve-OwnedProcessId -Before @(100) -After @(100, 200) | Should -Be 200
    }

    It 'claims the new PID when none were running before' {
        Resolve-OwnedProcessId -Before @() -After @(300) | Should -Be 300
    }

    It 'claims nothing when we attached to an already-running instance' {
        Resolve-OwnedProcessId -Before @(100) -After @(100) | Should -BeNullOrEmpty
    }

    It 'claims nothing when several new instances appeared (ambiguous - do not kill the wrong one)' {
        Resolve-OwnedProcessId -Before @(100) -After @(100, 200, 300) | Should -BeNullOrEmpty
    }

    It 'claims nothing when a process exited and none were started' {
        Resolve-OwnedProcessId -Before @(100, 200) -After @(200) | Should -BeNullOrEmpty
    }

    It 'tolerates empty/omitted snapshots' {
        Resolve-OwnedProcessId | Should -BeNullOrEmpty
        Resolve-OwnedProcessId -Before @() -After @() | Should -BeNullOrEmpty
    }
}

Describe 'Test-ProcessExited' {
    It 'reports exited when the PID is gone from the current snapshot' {
        Test-ProcessExited -ProcessId 200 -Current @(100) | Should -BeTrue
    }

    It 'reports not exited when the PID is still present' {
        Test-ProcessExited -ProcessId 200 -Current @(100, 200) | Should -BeFalse
    }

    It 'reports exited against an empty snapshot' {
        Test-ProcessExited -ProcessId 200 -Current @() | Should -BeTrue
    }
}

Describe 'Test-ValidExportName' {
    It 'accepts a real SMTP or display name' -ForEach @(
        @{ name = 'owner@company.com' }
        @{ name = 'John Smith' }
        @{ name = 'user@corp.com' }   # NOT the bare literal 'user'
    ) {
        Test-ValidExportName -Name $name | Should -BeTrue
    }

    It 'rejects blank/whitespace/null' -ForEach @(
        @{ name = '' }
        @{ name = '   ' }
        @{ name = $null }
    ) {
        Test-ValidExportName -Name $name | Should -BeFalse
    }

    It 'rejects the banned literal "user" regardless of case/padding' -ForEach @(
        @{ name = 'user' }
        @{ name = 'USER' }
        @{ name = '  User  ' }
    ) {
        Test-ValidExportName -Name $name | Should -BeFalse
    }

    It 'rejects the in-progress detection sentinel' -ForEach @(
        @{ name = 'Detecting...' }
        @{ name = 'Detecting owner...' }
    ) {
        Test-ValidExportName -Name $name | Should -BeFalse
    }

    It 'rejects the field when it still shows the cue text' {
        Test-ValidExportName -Name 'Pick or detect the user' -CueText 'Pick or detect the user' | Should -BeFalse
    }

    It 'rejects a degenerate name that yields no usable file name' -ForEach @(
        @{ name = '..' }
        @{ name = '.' }
    ) {
        Test-ValidExportName -Name $name | Should -BeFalse
    }
}

Describe 'Get-StoreExportPlan' {
    It 'includes the primary mailbox' {
        $p = Get-StoreExportPlan -Stores @((New-Store 'Mailbox - X' 'C:\x.ost' 0))
        $p[0].Include | Should -BeTrue
        $p[0].Category | Should -Be 'primary'
    }
    It 'includes mounted PST data files (type 3 / null)' {
        (Get-StoreExportPlan -Stores @((New-Store 'Old PST' 'C:\old.pst' 3)))[0].Category | Should -Be 'datafile'
        (Get-StoreExportPlan -Stores @((New-Store 'Loose' 'C:\loose.pst' $null)))[0].Include | Should -BeTrue
    }
    It 'skips public folders unless opted in' {
        (Get-StoreExportPlan -Stores @((New-Store 'Public Folders' '' 2)))[0].Include | Should -BeFalse
        (Get-StoreExportPlan -Stores @((New-Store 'Public Folders' '' 2)) -IncludePublicFolders $true)[0].Include | Should -BeTrue
    }
    It 'includes online archive by default, skippable' {
        (Get-StoreExportPlan -Stores @((New-Store 'Online Archive - X' 'C:\a.ost' 1)))[0].Category | Should -Be 'archive'
        (Get-StoreExportPlan -Stores @((New-Store 'Online Archive - X' 'C:\a.ost' 1)))[0].Include | Should -BeTrue
        (Get-StoreExportPlan -Stores @((New-Store 'Online Archive - X' 'C:\a.ost' 1)) -IncludeArchive $false)[0].Include | Should -BeFalse
    }
    It 'skips shared/delegate mailboxes unless opted in' {
        $shared = New-Store 'Team Inbox' 'C:\s.ost' 1
        (Get-StoreExportPlan -Stores @($shared))[0].Category | Should -Be 'shared'
        (Get-StoreExportPlan -Stores @($shared))[0].Include | Should -BeFalse
        (Get-StoreExportPlan -Stores @($shared) -IncludeShared $true)[0].Include | Should -BeTrue
    }
    It 'skips the target PST itself' {
        $p = Get-StoreExportPlan -Stores @((New-Store 'Export' 'C:\out\export.pst' 3)) -TargetPstPath 'C:\out\export.pst'
        $p[0].Include | Should -BeFalse
        $p[0].Category | Should -Be 'target'
    }
    It 'includes unknown store types to avoid data loss' {
        $p = Get-StoreExportPlan -Stores @((New-Store 'Mystery' '' 99))
        $p[0].Include | Should -BeTrue
        $p[0].Category | Should -Be 'unknown'
    }
}

Describe 'Test-CrossAccountCopyEnabled' {
    It 'returns $true when the DisableCrossAccountCopy policy is not set' {
        Mock Get-RegistryValue { $null }
        Test-CrossAccountCopyEnabled | Should -BeTrue
    }
    It 'treats a blank policy value as not set' {
        Mock Get-RegistryValue { '   ' }
        Test-CrossAccountCopyEnabled | Should -BeTrue
    }
    It 'returns $false when DisableCrossAccountCopy is set (export may silently skip)' {
        Mock Get-RegistryValue { 1 }
        Mock Write-Log {}
        Test-CrossAccountCopyEnabled | Should -BeFalse
    }
    It 'queries the DisableCrossAccountCopy value under the Outlook key' {
        Mock Get-RegistryValue { $null }
        Test-CrossAccountCopyEnabled | Out-Null
        Should -Invoke Get-RegistryValue -ParameterFilter { $Name -eq 'DisableCrossAccountCopy' }
    }
}

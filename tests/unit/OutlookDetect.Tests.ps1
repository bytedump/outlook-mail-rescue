#Requires -Version 5.1
# Unit tests for the pure helpers in src/OutlookDetect.ps1.
# Loaded via the -LoadOnly seam so no real Outlook/registry is touched.

BeforeAll {
    $entry = Join-Path $PSScriptRoot '..\..\Invoke-MailRescue.ps1'
    . $entry -LoadOnly

    # Build a fake Outlook AddressEntry so the SMTP-read fallback order can be unit-tested
    # without a live account. Omit a member to simulate it being missing/throwing.
    function New-FakeEntry {
        param(
            [string]$Type,
            [string]$Address,
            [scriptblock]$GetProperty,   # param($schema) -> string|throw
            $ExchangeUser,               # object with .PrimarySmtpAddress, or $null
            [switch]$NoPropertyAccessor,
            [switch]$NoExchangeUserMethod
        )
        $e = [pscustomobject]@{ Type = $Type; Address = $Address }
        if (-not $NoPropertyAccessor) {
            $pa = [pscustomobject]@{}
            $body = if ($GetProperty) { $GetProperty } else { { throw 'no proptag' } }
            $pa | Add-Member -MemberType ScriptMethod -Name GetProperty -Value $body
            $e | Add-Member -MemberType NoteProperty -Name PropertyAccessor -Value $pa
        }
        if (-not $NoExchangeUserMethod) {
            $e | Add-Member -MemberType NoteProperty -Name _ExchangeUser -Value $ExchangeUser
            $e | Add-Member -MemberType ScriptMethod -Name GetExchangeUser -Value { $this._ExchangeUser }
        }
        return $e
    }
}

Describe 'ConvertTo-OutlookBitness' {
    It 'maps <raw> to <expected>' -ForEach @(
        @{ raw = 'x64';   expected = 'x64' }
        @{ raw = '64';    expected = 'x64' }
        @{ raw = 'amd64'; expected = 'x64' }
        @{ raw = 'X86';   expected = 'x86' }
        @{ raw = '32';    expected = 'x86' }
    ) {
        ConvertTo-OutlookBitness $raw | Should -Be $expected
    }

    It 'returns $null for empty or unknown input' {
        ConvertTo-OutlookBitness '' | Should -BeNullOrEmpty
        ConvertTo-OutlookBitness 'sparc' | Should -BeNullOrEmpty
    }
}

Describe 'Resolve-OutlookFlavor' {
    It 'classic running always wins' {
        Resolve-OutlookFlavor -NewInstalled $true -ClassicInstalled $true `
            -NewRunning $true -ClassicRunning $true -UseNewOutlook 1 | Should -Be 'classic'
    }
    It 'new running when classic not running' {
        Resolve-OutlookFlavor -NewInstalled $true -ClassicInstalled $true `
            -NewRunning $true -ClassicRunning $false -UseNewOutlook 0 | Should -Be 'new'
    }
    It 'prefers new when UseNewOutlook=1 and new installed (none running)' {
        Resolve-OutlookFlavor -NewInstalled $true -ClassicInstalled $true `
            -NewRunning $false -ClassicRunning $false -UseNewOutlook 1 | Should -Be 'new'
    }
    It 'prefers classic when UseNewOutlook=0 and classic installed (none running)' {
        Resolve-OutlookFlavor -NewInstalled $true -ClassicInstalled $true `
            -NewRunning $false -ClassicRunning $false -UseNewOutlook 0 | Should -Be 'classic'
    }
    It 'falls back to classic when installed and no preference' {
        Resolve-OutlookFlavor -NewInstalled $false -ClassicInstalled $true `
            -NewRunning $false -ClassicRunning $false -UseNewOutlook $null | Should -Be 'classic'
    }
    It 'falls back to new when only new installed' {
        Resolve-OutlookFlavor -NewInstalled $true -ClassicInstalled $false `
            -NewRunning $false -ClassicRunning $false -UseNewOutlook $null | Should -Be 'new'
    }
    It 'returns none when nothing installed' {
        Resolve-OutlookFlavor -NewInstalled $false -ClassicInstalled $false `
            -NewRunning $false -ClassicRunning $false -UseNewOutlook $null | Should -Be 'none'
    }
}

Describe 'Resolve-PowerShellHostPath' {
    It 'wants 64-bit from a 32-bit process -> Sysnative' {
        Resolve-PowerShellHostPath -Want64 $true -CurrentIs64 $false | Should -Match '(?i)\\Sysnative\\'
    }
    It 'wants 64-bit from a 64-bit process -> System32' {
        Resolve-PowerShellHostPath -Want64 $true -CurrentIs64 $true | Should -Match '(?i)\\System32\\'
    }
    It 'wants 32-bit -> SysWOW64 regardless of current bitness' {
        Resolve-PowerShellHostPath -Want64 $false -CurrentIs64 $true | Should -Match '(?i)\\SysWOW64\\'
        Resolve-PowerShellHostPath -Want64 $false -CurrentIs64 $false | Should -Match '(?i)\\SysWOW64\\'
    }
    It 'always points at powershell.exe' {
        Resolve-PowerShellHostPath -Want64 $true -CurrentIs64 $true | Should -Match '(?i)powershell\.exe$'
    }
}

Describe 'Get-MismatchedPowerShellHost' {
    It 'returns $null when classic Outlook is not installed' {
        Get-MismatchedPowerShellHost -OutlookInfo ([pscustomobject]@{ ClassicInstalled = $false; ClassicBitness = 'x64' }) |
            Should -BeNullOrEmpty
    }
    It 'returns $null when the classic bitness is unknown' {
        Get-MismatchedPowerShellHost -OutlookInfo ([pscustomobject]@{ ClassicInstalled = $true; ClassicBitness = $null }) |
            Should -BeNullOrEmpty
    }
    It 'returns $null when Outlook bitness already matches the current process' {
        $same = if ([Environment]::Is64BitProcess) { 'x64' } else { 'x86' }
        Get-MismatchedPowerShellHost -OutlookInfo ([pscustomobject]@{ ClassicInstalled = $true; ClassicBitness = $same }) |
            Should -BeNullOrEmpty
    }
    It 'returns the opposite-bitness powershell.exe path on a mismatch' {
        $other = if ([Environment]::Is64BitProcess) { 'x86' } else { 'x64' }
        $r = Get-MismatchedPowerShellHost -OutlookInfo ([pscustomobject]@{ ClassicInstalled = $true; ClassicBitness = $other })
        $r | Should -Match '(?i)powershell\.exe$'
    }
    It 'falls back to Get-OutlookInfo when no -OutlookInfo is passed' {
        Mock Get-OutlookInfo { [pscustomobject]@{ ClassicInstalled = $false; ClassicBitness = $null } }
        Get-MismatchedPowerShellHost | Should -BeNullOrEmpty
        Should -Invoke Get-OutlookInfo -Times 1 -Exactly
    }
}

Describe 'Resolve-IdentityToken' {
    It 'returns $null for empty or whitespace input' -ForEach @(
        @{ in = '' }
        @{ in = '   ' }
        @{ in = "`t" }
    ) {
        Resolve-IdentityToken $in | Should -BeNullOrEmpty
    }

    It 'does not throw and returns $null on $null input' {
        { Resolve-IdentityToken $null } | Should -Not -Throw
        Resolve-IdentityToken $null | Should -BeNullOrEmpty
    }

    It 'keeps a normal SMTP address literally' {
        Resolve-IdentityToken 'a.b@c.com' | Should -BeExactly 'a.b@c.com'
    }

    It 'keeps the full work SMTP literally (identity of record, no stamp)' {
        Resolve-IdentityToken 'owner@company.com' |
            Should -BeExactly 'owner@company.com'
    }

    It 'never yields a dot-only token (no path traversal)' -ForEach @(
        @{ in = '.' }
        @{ in = '..' }
        @{ in = '...' }
    ) {
        Resolve-IdentityToken $in | Should -BeNullOrEmpty
    }

    It 'neutralizes reserved DOS device names' -ForEach @(
        @{ in = 'NUL' }
        @{ in = 'nul' }
        @{ in = 'CON' }
        @{ in = 'COM1' }
        @{ in = 'LPT9' }
    ) {
        $t = Resolve-IdentityToken $in
        $t | Should -Not -BeNullOrEmpty
        $t | Should -Not -Match '^(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$'
    }

    It 'strips the RTL-override (bidi) control char' {
        $rtl = "a$([char]0x202E)b@c.com"
        Resolve-IdentityToken $rtl | Should -BeExactly 'ab@c.com'
    }

    It 'replaces path separators and illegal chars (legacyExchangeDN)' {
        $dn = '/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=abc'
        $t = Resolve-IdentityToken $dn -MaxLength 200
        $t | Should -Not -BeNullOrEmpty
        $t | Should -Not -Match '[\\/:*?"<>|]'
    }

    It 'truncates an overlong identity to the default MaxLength (64)' {
        $long = 'a' * 300
        $t = Resolve-IdentityToken $long
        $t.Length | Should -BeLessOrEqual 64
        ($t -replace 'a', '') | Should -BeExactly ''
    }

    It 'honors a custom MaxLength' {
        (Resolve-IdentityToken ('b' * 50) -MaxLength 10).Length | Should -Be 10
    }

    It 'tolerates partial SMTP fragments without throwing' -ForEach @(
        @{ in = 'x@'; want = 'x@' }
        @{ in = '@y'; want = '@y' }
    ) {
        Resolve-IdentityToken $in | Should -BeExactly $want
    }
}

Describe 'Resolve-DetectedIdentity' {
    It 'tier smtp: valid primary SMTP wins, names the file with it' {
        $r = Resolve-DetectedIdentity -PrimarySmtp 'owner@company.com' -DisplayName 'Jane Doe'
        $r.Tier           | Should -Be 'smtp'
        $r.PrimarySmtp    | Should -BeExactly 'owner@company.com'
        $r.NameSourceText | Should -BeExactly 'owner@company.com'
        $r.DisplayName    | Should -BeExactly 'Jane Doe'
        $r.HasIdentity    | Should -BeTrue
    }

    It 'tier smtp wins even when exchange + display name are also present' {
        $r = Resolve-DetectedIdentity -PrimarySmtp 'a.b@c.com' `
            -ExchangeAddress '/o=Org/cn=Recipients/cn=abc' -DisplayName 'A B'
        $r.Tier           | Should -Be 'smtp'
        $r.NameSourceText | Should -BeExactly 'a.b@c.com'
    }

    It 'trims the primary SMTP' {
        $r = Resolve-DetectedIdentity -PrimarySmtp '  a@b.co  '
        $r.Tier        | Should -Be 'smtp'
        $r.PrimarySmtp | Should -BeExactly 'a@b.co'
    }

    It 'rejects a malformed SMTP and has no exchange/display fallback -> none' -ForEach @(
        @{ in = 'user' }
        @{ in = 'x@' }
        @{ in = '@y' }
    ) {
        $r = Resolve-DetectedIdentity -PrimarySmtp $in
        $r.Tier        | Should -Be 'none'
        $r.PrimarySmtp | Should -BeNullOrEmpty
        $r.HasIdentity | Should -BeFalse
    }

    It 'malformed SMTP but a display name exists -> displayname tier, never the bogus SMTP' {
        $r = Resolve-DetectedIdentity -PrimarySmtp 'user' -DisplayName 'Real Name'
        $r.Tier           | Should -Be 'displayname'
        $r.PrimarySmtp    | Should -BeNullOrEmpty
        $r.NameSourceText | Should -BeExactly 'Real Name'
    }

    It 'tier exchange: prefers the display name for the file name (DisplayName-preferred decision)' {
        $r = Resolve-DetectedIdentity -ExchangeAddress '/o=ExchangeLabs/cn=Recipients/cn=abc' -DisplayName 'John Smith'
        $r.Tier           | Should -Be 'exchange'
        $r.PrimarySmtp    | Should -BeNullOrEmpty
        $r.NameSourceText | Should -BeExactly 'John Smith'
        $r.HasIdentity    | Should -BeTrue
    }

    It 'tier exchange with no display name -> falls back to the legacyDN as the name source' {
        $dn = '/o=ExchangeLabs/cn=Recipients/cn=abc'
        $r = Resolve-DetectedIdentity -ExchangeAddress $dn
        $r.Tier           | Should -Be 'exchange'
        $r.NameSourceText | Should -BeExactly $dn
    }

    It 'tier displayname when only a display name is known' {
        $r = Resolve-DetectedIdentity -DisplayName 'Only Name'
        $r.Tier           | Should -Be 'displayname'
        $r.NameSourceText | Should -BeExactly 'Only Name'
        $r.HasIdentity    | Should -BeTrue
    }

    It 'tier none when nothing usable is present' -ForEach @(
        @{ smtp = ''; ex = ''; dn = '' }
        @{ smtp = '   '; ex = "`t"; dn = '  ' }
        @{ smtp = $null; ex = $null; dn = $null }
    ) {
        $r = Resolve-DetectedIdentity -PrimarySmtp $smtp -ExchangeAddress $ex -DisplayName $dn
        $r.Tier           | Should -Be 'none'
        $r.HasIdentity    | Should -BeFalse
        $r.NameSourceText | Should -BeNullOrEmpty
        $r.PrimarySmtp    | Should -BeNullOrEmpty
    }
}

Describe 'Get-PrimarySmtpFromEntry' {
    It 'returns $null for a null entry' {
        Get-PrimarySmtpFromEntry $null | Should -BeNullOrEmpty
    }

    It 'reads a direct SMTP-type entry from .Address' {
        $e = New-FakeEntry -Type 'SMTP' -Address 'direct@corp.com' -GetProperty { throw 'should not be called' }
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'direct@corp.com'
    }

    It 'reads PR_SMTP_ADDRESS via the Unicode proptag for an Exchange entry' {
        $e = New-FakeEntry -Type 'EX' -GetProperty { param($s) if ($s -match '001F') { 'uni@corp.com' } else { throw 'ansi not used' } }
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'uni@corp.com'
    }

    It 'falls back from Unicode to the ANSI proptag' {
        $e = New-FakeEntry -Type 'EX' -GetProperty { param($s) if ($s -match '001F') { throw 'unicode failed' } else { 'ansi@corp.com' } }
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'ansi@corp.com'
    }

    It 'falls back to GetExchangeUser().PrimarySmtpAddress when both proptags fail' {
        $eu = [pscustomobject]@{ PrimarySmtpAddress = 'exch@corp.com' }
        $e = New-FakeEntry -Type 'EX' -GetProperty { throw 'proptag failed' } -ExchangeUser $eu
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'exch@corp.com'
    }

    It 'null-checks GetExchangeUser (returns $null) without throwing' {
        $e = New-FakeEntry -Type 'EX' -GetProperty { throw 'proptag failed' } -ExchangeUser $null
        Get-PrimarySmtpFromEntry $e | Should -BeNullOrEmpty
    }

    It 'returns $null when every source fails or is missing' {
        $e = New-FakeEntry -Type 'EX' -NoPropertyAccessor -NoExchangeUserMethod
        Get-PrimarySmtpFromEntry $e | Should -BeNullOrEmpty
    }

    It 'skips an empty proptag value and keeps trying' {
        $e = New-FakeEntry -Type 'EX' -GetProperty { param($s) if ($s -match '001F') { '   ' } else { 'ansi@corp.com' } }
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'ansi@corp.com'
    }

    It 'skips an SMTP-type entry with a blank address and falls through to the proptag' {
        $e = New-FakeEntry -Type 'SMTP' -Address '' -GetProperty { 'proptag@corp.com' }
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'proptag@corp.com'
    }

    It 'trims the returned address' {
        $e = New-FakeEntry -Type 'SMTP' -Address '  spaced@corp.com  '
        Get-PrimarySmtpFromEntry $e | Should -BeExactly 'spaced@corp.com'
    }
}

Describe 'Invoke-WithRpcRetry' {
    # COMException reports RPC_E_CALL_REJECTED (0x80010001) as the signed int32 -2147418111.
    It 'returns immediately on success without sleeping' {
        $c = @{ calls = 0; sleeps = 0 }
        $action = { $c.calls++; 'ok' }.GetNewClosure()
        $sleep  = { $c.sleeps++ }.GetNewClosure()
        Invoke-WithRpcRetry -Action $action -Sleep $sleep | Should -Be 'ok'
        $c.calls  | Should -Be 1
        $c.sleeps | Should -Be 0
    }

    It 'retries on RPC_E_CALL_REJECTED, then returns the value' {
        $c = @{ calls = 0; sleeps = 0 }
        $action = {
            $c.calls++
            if ($c.calls -lt 3) { throw [System.Runtime.InteropServices.COMException]::new('rejected', -2147418111) }
            'ok'
        }.GetNewClosure()
        $sleep = { $c.sleeps++ }.GetNewClosure()
        Invoke-WithRpcRetry -Action $action -MaxAttempts 5 -Sleep $sleep | Should -Be 'ok'
        $c.calls  | Should -Be 3
        $c.sleeps | Should -Be 2
    }

    It 'retries while the result is null, then returns the first non-null' {
        $c = @{ calls = 0 }
        $action = { $c.calls++; if ($c.calls -lt 3) { $null } else { 'current-user' } }.GetNewClosure()
        Invoke-WithRpcRetry -Action $action -MaxAttempts 5 -Sleep { } | Should -Be 'current-user'
        $c.calls | Should -Be 3
    }

    It 'gives up and returns $null after exhausting attempts on persistent rejection' {
        $c = @{ calls = 0; sleeps = 0 }
        $action = { $c.calls++; throw [System.Runtime.InteropServices.COMException]::new('rejected', -2147418111) }.GetNewClosure()
        $sleep  = { $c.sleeps++ }.GetNewClosure()
        Invoke-WithRpcRetry -Action $action -MaxAttempts 4 -Sleep $sleep | Should -BeNullOrEmpty
        $c.calls  | Should -Be 4
        $c.sleeps | Should -Be 3
    }

    It 'returns $null when every attempt yields null (never a degraded value)' {
        $c = @{ calls = 0 }
        $action = { $c.calls++; $null }.GetNewClosure()
        Invoke-WithRpcRetry -Action $action -MaxAttempts 3 -Sleep { } | Should -BeNullOrEmpty
        $c.calls | Should -Be 3
    }

    It 'rethrows a non-RPC error immediately (does not mask genuine failures)' {
        $c = @{ calls = 0 }
        $action = { $c.calls++; throw [System.Runtime.InteropServices.COMException]::new('other', -1) }.GetNewClosure()
        { Invoke-WithRpcRetry -Action $action -MaxAttempts 5 -Sleep { } } | Should -Throw
        $c.calls | Should -Be 1
    }
}

Describe 'Compare-ReadbackIdentity' {
    It 'matches identical identities' {
        Compare-ReadbackIdentity -Confirmed 'a@b.com' -Reread 'a@b.com' | Should -Be 'match'
    }

    It 'matches case- and whitespace-insensitively' {
        Compare-ReadbackIdentity -Confirmed '  A@B.com ' -Reread 'a@b.com' | Should -Be 'match'
    }

    It 'flags a mismatch when the live re-read is a different person (ABORT)' {
        Compare-ReadbackIdentity -Confirmed 'leaver@corp.com' -Reread 'someone.else@corp.com' | Should -Be 'mismatch'
    }

    It 'reports unverified when the re-read came back empty (WARN, do not abort)' -ForEach @(
        @{ reread = '' }
        @{ reread = '   ' }
        @{ reread = $null }
    ) {
        Compare-ReadbackIdentity -Confirmed 'a@b.com' -Reread $reread | Should -Be 'unverified'
    }

    It 'reports unverified when there is nothing confirmed to compare against' {
        Compare-ReadbackIdentity -Confirmed '' -Reread 'a@b.com' | Should -Be 'unverified'
    }
}

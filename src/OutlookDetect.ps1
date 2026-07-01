#Requires -Version 5.1
<#
.SYNOPSIS
    Detect the Outlook landscape on a machine: new vs classic, install paths,
    bitness, the UseNewOutlook preference, and whether a MAPI profile exists.
.DESCRIPTION
    Get-OutlookInfo gathers everything the decision tree needs. The flavor-resolution
    and host-bitness logic are split into PURE helpers (Resolve-OutlookFlavor,
    Resolve-PowerShellHostPath) so they can be unit-tested without a real Outlook.
    Dot-sourced by Invoke-MailRescue.ps1.
#>

# Safe registry read: returns the value or $null, never throws (StrictMode-friendly).
function Get-RegistryValue {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop
        return $item.$Name
    } catch { return $null }
}

# Normalize an Office bitness string to 'x86' | 'x64' | $null. Pure.
function ConvertTo-OutlookBitness {
    param([string]$Raw)
    if ([string]::IsNullOrWhiteSpace($Raw)) { return $null }
    switch -Regex ($Raw.Trim().ToLower()) {
        '^(x64|64|amd64)$' { return 'x64' }
        '^(x86|32)$'       { return 'x86' }
        default            { return $null }
    }
}

# Decide which Outlook is "active" from already-gathered facts. Pure + testable.
# Inputs: booleans for installed/running, and the UseNewOutlook int (or $null).
function Resolve-OutlookFlavor {
    param(
        [bool]$NewInstalled,
        [bool]$ClassicInstalled,
        [bool]$NewRunning,
        [bool]$ClassicRunning,
        [Nullable[int]]$UseNewOutlook
    )
    if ($ClassicRunning) { return 'classic' }
    if ($NewRunning)     { return 'new' }
    # Nothing running: infer from preference, then from what is installed.
    if ($UseNewOutlook -eq 1 -and $NewInstalled) { return 'new' }
    if ($UseNewOutlook -eq 0 -and $ClassicInstalled) { return 'classic' }
    if ($ClassicInstalled) { return 'classic' }
    if ($NewInstalled)     { return 'new' }
    return 'none'
}

# Resolve the powershell.exe path for a desired bitness from the current process.
# Pure: depends only on $env:WINDIR and the two booleans. Returns a full path.
#   Want 64-bit from a 32-bit process -> Sysnative (escapes WOW64 redirection).
#   Want 64-bit from a 64-bit process -> System32.
#   Want 32-bit (from either)         -> SysWOW64.
function Resolve-PowerShellHostPath {
    param([Parameter(Mandatory)][bool]$Want64, [Parameter(Mandatory)][bool]$CurrentIs64)
    $win = $env:WINDIR
    if ($Want64) {
        $sub = if ($CurrentIs64) { 'System32' } else { 'Sysnative' }
    } else {
        $sub = 'SysWOW64'
    }
    return (Join-Path $win "$sub\WindowsPowerShell\v1.0\powershell.exe")
}

# Reserved DOS device names that are illegal as a bare file name on Windows.
$script:MrReservedDeviceNames = @('CON', 'PRN', 'AUX', 'NUL') +
    (1..9 | ForEach-Object { "COM$_" }) + (1..9 | ForEach-Object { "LPT$_" })

# Turn a raw identity (primary SMTP, display name, legacyExchangeDN, or junk) into a
# file-system-safe name token, or $null when nothing usable remains. PURE - no COM/IO.
# Keeps '@' and '.', so a full primary SMTP survives literally
# (owner@company.com). But it can NEVER produce '.', '..', a path
# separator, a control/bidi char, a reserved device name, or an overlong string.
# Callers MUST treat $null as "no usable name" (leave the field empty / require manual
# entry) - never fall back to the literal 'user'.
function Resolve-IdentityToken {
    param([string]$Identity, [int]$MaxLength = 64)

    if ([string]::IsNullOrWhiteSpace($Identity)) { return $null }

    # 1. Drop Unicode control & format chars (Cc/Cf): RTL-override, zero-width, etc.
    $t = $Identity -replace '[\p{Cc}\p{Cf}]', ''
    # 2. Trim, then collapse internal whitespace runs to '_'.
    $t = ($t.Trim() -replace '\s+', '_')
    # 3. Replace Windows-invalid file-name chars ( / \ : * ? " < > | ... ); '@'/'.' survive.
    $t = Remove-InvalidFileNameChars $t
    # 4. Strip leading/trailing dots, spaces, underscores -> '.'/'..' collapse to empty.
    $t = $t -replace '^[\s._]+', '' -replace '[\s._]+$', ''
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    # 5. Neutralize a bare reserved device name (CON, NUL, COM1, ...).
    if ($script:MrReservedDeviceNames -contains $t.ToUpperInvariant()) { $t = "_$t" }
    # 6. Cap length, then re-strip any trailing separator the cut may have exposed.
    if ($t.Length -gt $MaxLength) {
        $t = $t.Substring(0, $MaxLength) -replace '[\s._]+$', ''
    }
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    return $t
}

# Classify a detected identity into a confidence tier and pick the file-name basis.
# Inputs are the raw bits the COM probe gathers (any may be empty): the primary SMTP,
# an Exchange address (legacyExchangeDN / X500), and the display name. PURE.
#   smtp        -> a real-looking primary SMTP (local@domain.tld); names the file with it
#   exchange    -> only a legacyDN; names the file with the DisplayName when known, else
#                  the legacyDN itself (DisplayName-preferred decision)
#   displayname -> only a display name
#   none        -> nothing usable; the caller leaves the field empty (never 'user')
# PrimarySmtp (the identity of record for the confirm dialog / manifest / PST root-folder
# name) is set ONLY on the smtp tier. NameSourceText is what feeds Get-IdentityPstFileName.
function Resolve-DetectedIdentity {
    param([string]$PrimarySmtp, [string]$ExchangeAddress, [string]$DisplayName)

    $smtp = if ([string]::IsNullOrWhiteSpace($PrimarySmtp))     { $null } else { $PrimarySmtp.Trim() }
    $exch = if ([string]::IsNullOrWhiteSpace($ExchangeAddress)) { $null } else { $ExchangeAddress.Trim() }
    $name = if ([string]::IsNullOrWhiteSpace($DisplayName))     { $null } else { $DisplayName.Trim() }

    $smtpValid = [bool]($smtp -and ($smtp -match '^[^@\s]+@[^@\s]+\.[^@\s]+$'))

    if     ($smtpValid) { $tier = 'smtp';        $src = $smtp }
    elseif ($exch)      { $tier = 'exchange';    $src = if ($name) { $name } else { $exch } }
    elseif ($name)      { $tier = 'displayname'; $src = $name }
    else                { $tier = 'none';        $src = $null }

    return [pscustomobject]@{
        Tier           = $tier
        PrimarySmtp    = if ($smtpValid) { $smtp } else { $null }
        DisplayName    = $name
        NameSourceText = $src
        HasIdentity    = ($tier -ne 'none')
    }
}

# PR_SMTP_ADDRESS, addressed by its MAPI proptag URL. PropertyAccessor.GetProperty takes
# the schema as a URL STRING - passing the raw int 0x39FE001F throws on EVERY account
# (the naive bug). 0x..001F is the Unicode (PT_UNICODE) variant, 0x..001E the ANSI one.
$script:MrSmtpProptagUnicode = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001F'
$script:MrSmtpProptagAnsi    = 'http://schemas.microsoft.com/mapi/proptag/0x39FE001E'

# Read the primary SMTP from an Outlook AddressEntry-like object (guard #8), trying every
# source in order and isolating each in try/catch so one failing path falls through to the
# next: a direct SMTP-type entry's .Address, then PR_SMTP_ADDRESS via the Unicode then ANSI
# proptag URL, then GetExchangeUser().PrimarySmtpAddress (null-checked). Returns the SMTP
# or $null. Testable with a fake entry; the live call site passes Namespace.CurrentUser.
function Get-PrimarySmtpFromEntry {
    param($AddressEntry)
    if ($null -eq $AddressEntry) { return $null }

    $clean = { param($v) if ([string]::IsNullOrWhiteSpace($v)) { $null } else { ([string]$v).Trim() } }

    # 1. A direct SMTP-type entry already carries the SMTP in .Address.
    try {
        if ($AddressEntry.Type -eq 'SMTP') {
            $a = & $clean $AddressEntry.Address
            if ($a) { return $a }
        }
    } catch { }

    # 2./3. PR_SMTP_ADDRESS via the proptag URL: Unicode first, then ANSI.
    foreach ($schema in @($script:MrSmtpProptagUnicode, $script:MrSmtpProptagAnsi)) {
        try {
            $pa = $AddressEntry.PropertyAccessor
            if ($pa) {
                $v = & $clean $pa.GetProperty($schema)
                if ($v) { return $v }
            }
        } catch { }
    }

    # 4. Exchange user object (null-check before .PrimarySmtpAddress).
    try {
        $eu = $AddressEntry.GetExchangeUser()
        if ($eu) {
            $v = & $clean $eu.PrimarySmtpAddress
            if ($v) { return $v }
        }
    } catch { }

    return $null
}

# RPC_E_CALL_REJECTED, as the signed int32 a COMException reports (0x80010001).
$script:MrRpcCallRejected = -2147418111

# Retry a COM call that can fail while Outlook is still completing logon (guard #9). Polls
# $Action with a fixed backoff, retrying when it throws RPC_E_CALL_REJECTED or returns
# $null (logon may not be ready yet), until it yields a non-null result or attempts run
# out. Never collapses to a degraded value on the first throw, and rethrows a genuine
# (non-RPC) error instead of masking it. $Sleep is injectable so tests do not really wait.
# Returns the first non-null result, or $null if every attempt failed.
function Invoke-WithRpcRetry {
    param(
        [Parameter(Mandatory)][scriptblock]$Action,
        [int]$MaxAttempts = 10,
        [int]$DelayMs = 500,
        [scriptblock]$Sleep = { param($ms) Start-Sleep -Milliseconds $ms }
    )
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        $retryable = $false
        try {
            $r = & $Action
            if ($null -ne $r) { return $r }
            $retryable = $true   # null result: logon may not be done yet
        } catch {
            $retryable = ($_.Exception.HResult -eq $script:MrRpcCallRejected)
            if (-not $retryable) { throw }   # a genuine error: surface it, do not mask
        }
        if (-not $retryable) { break }
        if ($i -lt $MaxAttempts) { & $Sleep $DelayMs }
    }
    return $null
}

# Compare the identity confirmed at prefill time against the one re-read live inside the
# export runspace (guard #4). 'match' = same person (proceed); 'mismatch' = a DIFFERENT
# identity surfaced, so the profile/CurrentUser changed under us (abort); 'unverified' =
# the re-read produced nothing or there was nothing to compare against (warn, but do not
# block on a failed re-read). Case- and whitespace-insensitive. Pure.
function Compare-ReadbackIdentity {
    param([string]$Confirmed, [string]$Reread)
    $norm = { param($v) if ([string]::IsNullOrWhiteSpace($v)) { $null } else { $v.Trim().ToLowerInvariant() } }
    $c = & $norm $Confirmed
    $r = & $norm $Reread
    if (-not $c -or -not $r) { return 'unverified' }
    if ($c -eq $r) { return 'match' }
    return 'mismatch'
}

# Probe classic Outlook over COM for the logged-on owner's identity, wiring the tested
# helpers together: snapshot OUTLOOK.EXE PIDs, attach/create Outlook.Application, MAPI
# Logon with showDialog:$false (never raise a profile picker), poll CurrentUser through the
# RPC retry, read the primary SMTP, and classify with Resolve-DetectedIdentity. Always
# returns a Resolve-DetectedIdentity result (tier 'none' on any failure - never throws,
# never the literal 'user'). Quits ONLY an Outlook instance this call started (ownership
# by PID diff). LIVE-ONLY: the COM calls cannot run on a machine without classic Outlook,
# so this orchestration is validated by the manual e2e, not unit tests - every helper it
# composes is unit-tested. Not wired into the GUI yet; the attach-only gate (guard #5) is
# the caller's responsibility.
function Get-DetectedOwnerIdentity {
    [CmdletBinding()]
    param([int]$RpcMaxAttempts = 20, [int]$RpcDelayMs = 500)

    $before = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $outlook = $null; $ns = $null; $ownedPid = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        $after = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $ownedPid = Resolve-OwnedProcessId -Before $before -After $after

        $ns = $outlook.GetNamespace('MAPI')
        try { $ns.Logon($null, $null, $false, $false) } catch { }  # ignore "already logged on"

        $getCurrentUser = { $ns.CurrentUser }.GetNewClosure()
        $entry = Invoke-WithRpcRetry -Action $getCurrentUser -MaxAttempts $RpcMaxAttempts -DelayMs $RpcDelayMs
        if (-not $entry) { return (Resolve-DetectedIdentity) }

        $ae = $null;   try { $ae = $entry.AddressEntry } catch { }
        $disp = $null; try { $disp = $entry.Name } catch { }
        $exch = $null; try { if ($ae) { $exch = $ae.Address } } catch { }
        $smtp = Get-PrimarySmtpFromEntry $ae

        return Resolve-DetectedIdentity -PrimarySmtp $smtp -ExchangeAddress $exch -DisplayName $disp
    } catch {
        Write-Log WARN "Identity probe failed: $($_.Exception.Message)"
        return (Resolve-DetectedIdentity)
    } finally {
        Remove-ComRef $ns
        # Quit ONLY if we started Outlook from nothing. If any OUTLOOK.EXE existed before this
        # probe ($before non-empty) we attached to the user's live session - and .Quit() closes
        # the CONNECTED application, not a PID, so it would shut the user's Outlook even though
        # $ownedPid pointed at a transient COM-activation PID. Never quit in that case (guard #1
        # hardening; this is what closed an already-open classic Outlook during the #7 test).
        if ($ownedPid -and $outlook -and $before.Count -eq 0) {
            Write-Log INFO "Identity probe: quitting the Outlook it launched (PID $ownedPid)."
            try { $outlook.Quit() } catch { }
        }
        Remove-ComRef $outlook
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

# Auto-launch classic Outlook and KEEP it open (guard #7). Unlike Get-DetectedOwnerIdentity (which
# quits the instance it starts), this leaves Outlook running for the rest of the export workflow;
# the GUI owns it and quits it on FormClosing. Launches the real OUTLOOK.EXE via Start-Process (a
# COM-only launch can shut down again once its refs are released - a UI process stays up), then
# claims the new PID with the unit-tested Resolve-OwnedProcessId diff: exactly one new PID = ours;
# zero/ambiguous = a race, own nothing rather than risk quitting someone else's Outlook. Polls
# briefly because OUTLOOK.EXE takes a moment to appear. Returns the owned PID, or $null. LIVE-ONLY.
function Start-OwnedOutlook {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ClassicPath, [int]$MaxWaitMs = 10000, [int]$PollMs = 500)

    $before = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if ($before.Count -gt 0) {
        # Classic Outlook is already running (the construction-time check was stale). Attach
        # happens later via the probe; we did not start it, so we own nothing and must not quit it.
        Write-Log INFO 'Classic Outlook already running at startup; not auto-launching (own nothing).'
        return $null
    }
    try {
        Start-Process -FilePath $ClassicPath | Out-Null
    } catch {
        Write-Log WARN "Could not auto-launch classic Outlook: $($_.Exception.Message)"
        return $null
    }

    $ownedPid = $null
    $tries = [Math]::Max(1, [int]($MaxWaitMs / $PollMs))
    for ($i = 0; $i -lt $tries; $i++) {
        $after = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $ownedPid = Resolve-OwnedProcessId -Before $before -After $after
        if ($ownedPid) { break }
        Start-Sleep -Milliseconds $PollMs
    }
    if ($ownedPid) { Write-Log INFO "Auto-launched classic Outlook (PID $ownedPid); owned until app close." }
    else { Write-Log WARN 'Auto-launched classic Outlook but could not claim its PID (own nothing).' }
    return $ownedPid
}

# Gracefully quit a classic Outlook instance the GUI owns (guard #7), when the app closes. Quits
# only if our PID is still alive; re-attaches over COM and calls Quit (never force-kills - that
# risks PST corruption), then polls up to ~30s for the PID to leave (Test-ProcessExited). Returns
# $true if it exited. LIVE-ONLY.
function Stop-OwnedOutlook {
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$OwnedPid, [int]$MaxWaitSeconds = 30)

    $alive = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if (Test-ProcessExited -ProcessId $OwnedPid -Current $alive) { return $true }

    $outlook = $null
    try {
        $outlook = New-Object -ComObject Outlook.Application
        try { $outlook.Quit() } catch { }
    } catch {
        Write-Log WARN "Could not quit owned Outlook (PID $OwnedPid): $($_.Exception.Message)"
    } finally {
        Remove-ComRef $outlook
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    $gone = $false
    for ($i = 0; $i -lt $MaxWaitSeconds; $i++) {
        $left = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        if (Test-ProcessExited -ProcessId $OwnedPid -Current $left) { $gone = $true; break }
        Start-Sleep -Milliseconds 1000
    }
    if (-not $gone) { Write-Log WARN "Owned Outlook (PID $OwnedPid) still running ~${MaxWaitSeconds}s after Quit; left as-is (no force-kill)." }
    return $gone
}

# Office versions classic Outlook may live under (newest first).
$script:MrOfficeVersions = @('16.0', '15.0')

function Get-OutlookInfo {
    [CmdletBinding()]
    param()

    # --- Classic install + path (App Paths is the canonical "installed + where") ---
    $classicPath = Get-RegistryValue 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' '(default)'
    if (-not $classicPath) {
        $classicPath = Get-RegistryValue 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE' '(default)'
    }
    $classicInstalled = [bool]($classicPath -and (Test-Path $classicPath))

    # --- Classic bitness: prefer Click-to-Run Platform, then per-version Bitness ---
    $bitnessRaw = Get-RegistryValue 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' 'Platform'
    if (-not $bitnessRaw) {
        foreach ($v in $script:MrOfficeVersions) {
            $bitnessRaw = Get-RegistryValue "HKLM:\SOFTWARE\Microsoft\Office\$v\Outlook" 'Bitness'
            if ($bitnessRaw) { break }
        }
    }
    $classicBitness = ConvertTo-OutlookBitness $bitnessRaw
    # Last-resort heuristic from the install path.
    if (-not $classicBitness -and $classicPath) {
        $classicBitness = if ($classicPath -match '(?i)Program Files \(x86\)') { 'x86' } else { 'x64' }
    }

    # --- New Outlook (MSIX) ---
    $newInstalled = $false
    $newVersion   = $null
    try {
        $pkg = Get-AppxPackage -Name 'Microsoft.OutlookForWindows' -ErrorAction Stop | Select-Object -First 1
        if ($pkg) { $newInstalled = $true; $newVersion = $pkg.Version }
    } catch { }

    # --- Running processes ---
    $classicRunning = [bool](Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue)
    $newRunning     = [bool](Get-Process -Name 'olk' -ErrorAction SilentlyContinue)

    # --- UseNewOutlook preference (HKCU) ---
    $useNew = $null
    foreach ($v in $script:MrOfficeVersions) {
        $val = Get-RegistryValue "HKCU:\Software\Microsoft\office\$v\outlook\preferences" 'UseNewOutlook'
        if ($null -ne $val) { $useNew = [int]$val; break }
    }

    # --- MAPI profile present? (either modern Office or the legacy Messaging path) ---
    $hasProfile = $false
    foreach ($v in $script:MrOfficeVersions) {
        if (Test-Path "HKCU:\Software\Microsoft\Office\$v\Outlook\Profiles\*") { $hasProfile = $true; break }
    }
    if (-not $hasProfile) {
        $hasProfile = [bool](Test-Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\*')
    }

    $flavor = Resolve-OutlookFlavor -NewInstalled $newInstalled -ClassicInstalled $classicInstalled `
        -NewRunning $newRunning -ClassicRunning $classicRunning -UseNewOutlook $useNew

    return [pscustomobject]@{
        ClassicInstalled = $classicInstalled
        ClassicPath      = $classicPath
        ClassicBitness   = $classicBitness
        NewInstalled     = $newInstalled
        NewVersion       = $newVersion
        ClassicRunning   = $classicRunning
        NewRunning       = $newRunning
        UseNewOutlook    = $useNew
        HasMapiProfile   = $hasProfile
        ActiveFlavor     = $flavor
    }
}

# Return a powershell.exe path whose bitness matches Outlook but differs from the
# current process (so the caller should relaunch there); $null when already matched
# or when Outlook/its bitness is unknown.
function Get-MismatchedPowerShellHost {
    [CmdletBinding()]
    param([pscustomobject]$OutlookInfo)
    if (-not $OutlookInfo) { $OutlookInfo = Get-OutlookInfo }
    if (-not $OutlookInfo.ClassicInstalled -or -not $OutlookInfo.ClassicBitness) { return $null }

    $currentIs64 = [Environment]::Is64BitProcess
    $outlookIs64 = ($OutlookInfo.ClassicBitness -eq 'x64')
    if ($currentIs64 -eq $outlookIs64) { return $null }

    $path = Resolve-PowerShellHostPath -Want64 $outlookIs64 -CurrentIs64 $currentIs64
    if (Test-Path $path) { return $path }
    return $null
}

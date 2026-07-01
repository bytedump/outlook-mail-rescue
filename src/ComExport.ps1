#Requires -Version 5.1
<#
.SYNOPSIS
    Full mailbox export to a single Unicode PST via the classic Outlook COM model.
.DESCRIPTION
    There is NO COM equivalent of the File > Import/Export wizard, so the export is a
    folder copy: create a new Unicode PST (AddStoreEx), then CopyTo each top-level
    folder of every included source store (CopyTo brings the whole subtree + items in
    one call - do NOT recurse afterward or it double-copies). Item counts are validated
    per top-level folder because CopyTo skips failed items silently.

    Pure helpers (Get-ExportFileName, Get-StoreExportPlan, sanitizers) are unit-tested.
    The COM functions require classic Outlook and are validated manually (see README /
    plan verification) - they cannot run on a machine that only has the new Outlook.
#>

$script:OlStoreUnicode = 2   # OlStoreType.olStoreUnicode

# OlExchangeStoreType values, for readability.
$script:OlPrimaryExchangeMailbox    = 0
$script:OlExchangeMailbox           = 1   # delegate / shared
$script:OlExchangePublicFolder      = 2
$script:OlNotExchange               = 3   # PST / data file
$script:OlAdditionalExchangeMailbox = 4

# ---------------- Pure helpers (unit-tested) ----------------

# Strip characters Windows forbids in file names. Pure.
function Remove-InvalidFileNameChars {
    param([string]$Name)
    if ([string]::IsNullOrEmpty($Name)) { return $Name }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $sb = [System.Text.StringBuilder]::new()
    foreach ($ch in $Name.ToCharArray()) {
        if ($invalid -contains $ch) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

# Normalize a single name token (username/ticket): trim, spaces->_, drop invalid,
# fall back to a default when empty. Pure.
function Format-PathToken {
    param([string]$Value, [string]$Fallback)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Fallback }
    $v = ($Value.Trim() -replace '\s+', '_')
    $v = Remove-InvalidFileNameChars $v
    if ([string]::IsNullOrWhiteSpace($v)) { return $Fallback }
    return $v
}

# Build the output PST file name from the template. Pass -Stamp for deterministic
# tests; defaults to now otherwise. Pure given a Stamp.
function Get-ExportFileName {
    param(
        [string]$Template = '{username}_{ticket}_{stamp}.pst',
        [string]$Username,
        [string]$Ticket,
        [string]$Stamp
    )
    if ([string]::IsNullOrWhiteSpace($Stamp)) { $Stamp = Get-Date -Format 'yyyyMMdd-HHmmss' }
    $u = Format-PathToken $Username 'user'
    $t = Format-PathToken $Ticket 'noticket'
    $name = $Template.Replace('{username}', $u).Replace('{ticket}', $t).Replace('{stamp}', $Stamp)
    return (Remove-InvalidFileNameChars $name)
}

# Build the PST file name from a detected/typed identity: full primary SMTP, no stamp
# (per project decision -> owner@company.com.pst). Returns $null when
# the identity yields no usable token; the caller MUST treat $null as "no auto name"
# (leave the field empty / require manual entry) and never fall back to 'user'. Pure.
# Collision handling (when the file already exists) is the caller's job - see guard #10.
function Get-IdentityPstFileName {
    param([string]$Identity, [int]$MaxLength = 64)
    $token = Resolve-IdentityToken -Identity $Identity -MaxLength $MaxLength
    if (-not $token) { return $null }
    return "$token.pst"
}

# Gate the export on a real, usable identity in the name field (guard #6). Rejects blank,
# the cue/placeholder text, the banned literal 'user' (the old fallback), the in-progress
# 'Detecting...' sentinel, and any value that yields no usable file name (e.g. '.'/'..').
# The GUI keeps Export disabled until this returns $true. Pure.
function Test-ValidExportName {
    param([string]$Name, [string]$CueText)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $n = $Name.Trim()
    if ($CueText -and [string]::Equals($n, $CueText.Trim(), [System.StringComparison]::Ordinal)) { return $false }
    if ($n -match '^(?i)user$') { return $false }
    if ($n -match '^(?i)detecting') { return $false }
    if (-not (Get-IdentityPstFileName -Identity $n)) { return $false }
    return $true
}

# Derive the export outcome from the result counts/flags (guard #13). "Complete" means
# every source item landed in the PST and nothing was blocked or stalled; anything less
# is "incomplete" so the technician never offboards a mailbox that was not fully captured.
# A count mismatch in EITHER direction is degraded: fewer = missing items, more = a
# possible CopyTo double-copy. Pure.
function Get-ExportOutcome {
    param(
        [Parameter(Mandatory)][int]$SourceItems,
        [Parameter(Mandatory)][int]$CopiedItems,
        [bool]$SyncStalled,
        [bool]$CrossAccountBlocked
    )
    $reasons = @()
    if ($CopiedItems -lt $SourceItems) {
        $reasons += "Copied $CopiedItems of $SourceItems items ($($SourceItems - $CopiedItems) missing)"
    } elseif ($CopiedItems -gt $SourceItems) {
        $reasons += "Copied $CopiedItems of $SourceItems items ($($CopiedItems - $SourceItems) more than source - possible duplication)"
    }
    if ($SyncStalled) {
        $reasons += 'Mailbox sync did not finish (stalled) - recent items may be missing'
    }
    if ($CrossAccountBlocked) {
        $reasons += 'Cross-account copy blocked by policy - delegate/shared stores were skipped'
    }
    $degraded = $reasons.Count -gt 0
    return [pscustomobject]@{
        Degraded = $degraded
        Reasons  = $reasons
        Title    = if ($degraded) { 'Export incomplete' } else { 'Export complete' }
    }
}

# Decide which stores to export. Operates on plain objects with DisplayName, FilePath,
# ExchangeStoreType (int) - so it is unit-testable without COM. Returns the same
# records annotated with Include (bool), Reason, Category. Pure.
function Get-StoreExportPlan {
    param(
        [Parameter(Mandatory)][object[]]$Stores,
        [bool]$IncludeArchive = $true,
        [bool]$IncludeShared = $false,
        [bool]$IncludePublicFolders = $false,
        [string]$TargetPstPath
    )
    $plan = foreach ($s in $Stores) {
        $type = $null
        if ($s.PSObject.Properties.Name -contains 'ExchangeStoreType' -and $null -ne $s.ExchangeStoreType) {
            $type = [int]$s.ExchangeStoreType
        }
        $name = [string]$s.DisplayName
        $path = [string]$s.FilePath
        $include = $true; $reason = ''; $category = ''

        if ($TargetPstPath -and $path -and ($path -ieq $TargetPstPath)) {
            $include = $false; $reason = 'target PST (skip self)'; $category = 'target'
        }
        elseif ($type -eq $script:OlExchangePublicFolder) {
            $category = 'public'; $include = $IncludePublicFolders
            $reason = if ($include) { 'public folders (opted in)' } else { 'public folders (skipped)' }
        }
        elseif ($type -eq $script:OlPrimaryExchangeMailbox) {
            $include = $true; $reason = 'primary mailbox'; $category = 'primary'
        }
        elseif ($type -eq $script:OlNotExchange -or $null -eq $type) {
            $include = $true; $reason = 'mounted data file (PST)'; $category = 'datafile'
        }
        elseif ($type -eq $script:OlExchangeMailbox -or $type -eq $script:OlAdditionalExchangeMailbox) {
            if ($name -match '(?i)archive') {
                $category = 'archive'; $include = $IncludeArchive
                $reason = if ($include) { 'online archive' } else { 'online archive (skipped)' }
            } else {
                $category = 'shared'; $include = $IncludeShared
                $reason = if ($include) { 'shared/delegate mailbox (opted in)' } else { 'shared/delegate mailbox (skipped)' }
            }
        }
        else {
            $include = $true; $reason = "unknown store type ($type) - included to avoid data loss"; $category = 'unknown'
        }

        [pscustomobject]@{
            DisplayName       = $name
            FilePath          = $path
            ExchangeStoreType = $type
            Include           = $include
            Reason            = $reason
            Category          = $category
            ComStore          = if ($s.PSObject.Properties.Name -contains 'ComStore') { $s.ComStore } else { $null }
        }
    }
    return @($plan)
}

# ---------------- COM functions (classic Outlook; validated manually) ----------------

# CopyTo silently no-ops when DisableCrossAccountCopy is set. Return $true if safe.
function Test-CrossAccountCopyEnabled {
    foreach ($v in '16.0', '15.0') {
        $val = Get-RegistryValue "HKCU:\Software\Microsoft\Office\$v\Outlook" 'DisableCrossAccountCopy'
        if ($null -ne $val -and "$val".Trim() -ne '') {
            Write-Log WARN "DisableCrossAccountCopy is set ($v): '$val'. CopyTo may silently skip data."
            return $false
        }
    }
    return $true
}

# Snapshot Session.Stores into plain records (keeping a live .ComStore reference).
function Get-OutlookStores {
    param([Parameter(Mandatory)]$Namespace)
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($s in $Namespace.Stores) {
        $type = $null; $name = $null; $path = $null
        try { $name = $s.DisplayName } catch { }
        try { $path = $s.FilePath } catch { }
        try { $type = [int]$s.ExchangeStoreType } catch { }
        $list.Add([pscustomobject]@{
            DisplayName = $name; FilePath = $path; ExchangeStoreType = $type; ComStore = $s
        })
    }
    return @($list)
}

# Create the target Unicode PST and return its Store COM object.
function New-PstStore {
    param([Parameter(Mandatory)]$Namespace, [Parameter(Mandatory)][string]$PstPath)
    $dir = [System.IO.Path]::GetDirectoryName($PstPath)
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    if (Test-Path $PstPath) { throw "Target PST already exists: $PstPath" }

    $Namespace.AddStoreEx($PstPath, $script:OlStoreUnicode)
    foreach ($s in $Namespace.Stores) { if ($s.FilePath -ieq $PstPath) { return $s } }
    throw "Failed to create or locate the target PST store: $PstPath"
}

# Recursively count items across a folder subtree (for copy validation).
function Get-FolderItemCountRecursive {
    param([Parameter(Mandatory)]$Folder)
    $count = 0
    try { $count = [int]$Folder.Items.Count } catch { }
    try { foreach ($sf in $Folder.Folders) { $count += Get-FolderItemCountRecursive $sf } } catch { }
    return $count
}

# Copy every top-level folder of a source store into the target root. Returns a
# summary with source/copied counts and any per-folder issues.
function Copy-StoreToTarget {
    param(
        [Parameter(Mandatory)]$SourceStore,
        [Parameter(Mandatory)]$TargetRootFolder,
        [scriptblock]$OnProgress
    )
    $srcRoot = $SourceStore.GetRootFolder()
    $foldersCopied = 0; $srcItems = 0; $dstItems = 0
    $issues = [System.Collections.Generic.List[string]]::new()

    foreach ($folder in $srcRoot.Folders) {
        $fname = $folder.Name
        if ($OnProgress) { & $OnProgress 'folder' "$($srcRoot.Name)\$fname" }
        $before = Get-FolderItemCountRecursive $folder
        $srcItems += $before
        try {
            $copied = $folder.CopyTo($TargetRootFolder)   # whole subtree + items
            $foldersCopied++
            $after = Get-FolderItemCountRecursive $copied
            $dstItems += $after
            if ($after -lt $before) {
                $issues.Add("'$fname': copied $after of $before items (some skipped)")
                Write-Log WARN "Folder '$fname': copied $after of $before items"
            }
        } catch {
            $issues.Add("'$fname': copy failed - $($_.Exception.Message)")
            Write-Log ERROR "Folder '$fname' copy failed: $($_.Exception.Message)"
        }
    }

    return [pscustomobject]@{
        FoldersCopied = $foldersCopied
        SourceItems   = $srcItems
        CopiedItems   = $dstItems
        Issues        = @($issues)
    }
}

# Release a COM object quietly.
function Remove-ComRef {
    param($Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
    }
}

# Resolve which OUTLOOK.EXE PID we own (guard #1), by diffing the process snapshot taken
# before launching COM against the one taken after. Exactly one new PID = we started it
# (own it, quit it on cleanup). Zero new = we attached to a running instance (NOT ours,
# leave it alone). More than one new = ambiguous (a race with another launcher) - claim
# nothing rather than risk quitting the wrong Outlook. Replaces the racy
# `$startedByUs = -not (Get-Process OUTLOOK)` heuristic. Pure.
function Resolve-OwnedProcessId {
    param([int[]]$Before = @(), [int[]]$After = @())
    $new = @($After | Where-Object { $_ -notin $Before })
    if ($new.Count -eq 1) { return $new[0] }
    return $null
}

# Verify a PID we quit actually left (guard #1): true when it is absent from the current
# OUTLOOK.EXE snapshot. Pure.
function Test-ProcessExited {
    param([Parameter(Mandatory)][int]$ProcessId, [int[]]$Current = @())
    return ($ProcessId -notin $Current)
}

# Orchestrate the full export. $Config supplies Include* toggles. Returns a summary.
function Export-MailboxToPst {
    param(
        [Parameter(Mandatory)][string]$PstPath,
        [hashtable]$Config = @{},
        [switch]$WaitForSync,
        [int]$SyncTimeoutMinutes = 30,
        [string]$ExpectedSmtp,
        [scriptblock]$OnProgress
    )
    $includeArchive = if ($Config.ContainsKey('IncludeArchive')) { [bool]$Config.IncludeArchive } else { $true }
    $includeShared  = if ($Config.ContainsKey('IncludeSharedMailboxes')) { [bool]$Config.IncludeSharedMailboxes } else { $false }
    $includePublic  = if ($Config.ContainsKey('IncludePublicFolders')) { [bool]$Config.IncludePublicFolders } else { $false }

    # Guard #1 ownership: snapshot OUTLOOK.EXE PIDs before launching COM so we can tell the
    # instance we start from one we merely attach to (Resolve-OwnedProcessId), and quit only
    # the one we own. Replaces the racy `-not (Get-Process OUTLOOK)` heuristic.
    $before = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $outlook = $null; $ns = $null; $targetRoot = $null; $ownedPid = $null
    $allIssues = [System.Collections.Generic.List[string]]::new()
    $storesExported = 0; $foldersCopied = 0; $srcItems = 0; $dstItems = 0
    $success = $false
    $syncStalled = $false; $crossAccountBlocked = $false   # guard #13 outcome inputs

    try {
        $crossAccountBlocked = -not (Test-CrossAccountCopyEnabled)
        if ($crossAccountBlocked) {
            $allIssues.Add('DisableCrossAccountCopy is set; copy may be blocked. Clear it and retry.')
        }

        if ($OnProgress) { & $OnProgress 'init' 'Starting Outlook (COM)...' }
        $outlook = New-Object -ComObject Outlook.Application
        $after = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $ownedPid = Resolve-OwnedProcessId -Before $before -After $after
        $ns = $outlook.GetNamespace('MAPI')

        if ($WaitForSync) {
            if ($OnProgress) { & $OnProgress 'sync' 'Waiting for mailbox sync to finish...' }
            $sync = Wait-OutlookSync -Namespace $ns -TimeoutMinutes $SyncTimeoutMinutes -OnProgress $OnProgress
            $syncStalled = [bool]$sync.Stalled
            if ($syncStalled) { $allIssues.Add("Sync did not stabilize within $SyncTimeoutMinutes min; exported cached state ($($sync.LastCount) items).") }
        }

        # Guard #4b: now that Outlook is fully logged on inside the export runspace, re-read the
        # logged-on owner and compare it to the identity the technician confirmed. A profile/identity
        # switch between confirm and export must abort BEFORE the PST is created, so we never write
        # the wrong mailbox under the confirmed name. Empty -ExpectedSmtp skips the check.
        if ($ExpectedSmtp) {
            $reread = $null
            try {
                $cu = $ns.CurrentUser
                $cae = $null; try { $cae = $cu.AddressEntry } catch { }
                $reread = Get-PrimarySmtpFromEntry $cae
            } catch { Write-Log WARN "Could not re-read CurrentUser for the identity check: $($_.Exception.Message)" }
            $verdict = Compare-ReadbackIdentity -Confirmed $ExpectedSmtp -Reread $reread
            if ($verdict -eq 'mismatch') {
                throw "Logged-on mailbox changed since confirmation (expected '$ExpectedSmtp', now '$reread'); aborted before creating the PST."
            } elseif ($verdict -eq 'unverified') {
                Write-Log WARN "Could not verify the logged-on identity against '$ExpectedSmtp' (re-read empty); continuing."
            } else {
                Write-Log INFO "Identity re-check OK: $ExpectedSmtp."
            }
        }

        if ($OnProgress) { & $OnProgress 'init' "Creating target PST: $PstPath" }
        $target = New-PstStore -Namespace $ns -PstPath $PstPath
        $targetRoot = $target.GetRootFolder()

        $stores = Get-OutlookStores -Namespace $ns
        $plan = Get-StoreExportPlan -Stores $stores -IncludeArchive $includeArchive `
            -IncludeShared $includeShared -IncludePublicFolders $includePublic -TargetPstPath $PstPath

        foreach ($p in $plan) {
            $state = if ($p.Include) { 'INCLUDE' } else { 'skip' }
            Write-Log INFO "Store '$($p.DisplayName)' [$($p.Category)] -> $state ($($p.Reason))"
        }

        foreach ($p in ($plan | Where-Object { $_.Include })) {
            if (-not $p.ComStore) { continue }
            if ($OnProgress) { & $OnProgress 'store' "Exporting: $($p.DisplayName)" }
            $r = Copy-StoreToTarget -SourceStore $p.ComStore -TargetRootFolder $targetRoot -OnProgress $OnProgress
            $storesExported++
            $foldersCopied += $r.FoldersCopied
            $srcItems += $r.SourceItems
            $dstItems += $r.CopiedItems
            foreach ($i in $r.Issues) { $allIssues.Add("[$($p.DisplayName)] $i") }
        }

        $success = $true
        Write-Log INFO "Export done: stores=$storesExported folders=$foldersCopied items $dstItems/$srcItems"
    }
    catch {
        Write-Log FATAL "Export failed: $($_.Exception.Message)"
        $allIssues.Add("Fatal: $($_.Exception.Message)")
    }
    finally {
        # Detach the target PST from the profile (the .pst file stays on disk).
        if ($ns -and $targetRoot) {
            try { $ns.RemoveStore($targetRoot) } catch { Write-Log WARN "Could not unmount target PST: $($_.Exception.Message)" }
        }
        # Quit ONLY an Outlook we started from nothing (guard #1, hardened). If any OUTLOOK.EXE
        # was already running before this export ($before non-empty) - the app auto-launched it
        # (#7) or the user/tech had it open - leave it: .Quit() closes the CONNECTED app regardless
        # of which PID we think we own, so quitting here would kill the user's/app-owned session.
        # The app owns its own launch and quits it on close; an attach leaves the user's Outlook be.
        # Poll briefly so Test-ProcessExited does not false-warn; never force-kill (PST corruption).
        if ($ownedPid -and $outlook -and $before.Count -eq 0) {
            try { $outlook.Quit() } catch { }
            # Poll up to ~30s: after a large export Outlook can take 10-20s+ to flush and
            # close, so a short wait false-warns. Break as soon as it is gone; only warn if
            # it is genuinely stuck. Never force-kill (risks PST corruption).
            $gone = $false
            for ($i = 0; $i -lt 30; $i++) {
                $left = @(Get-Process -Name 'OUTLOOK' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
                if (Test-ProcessExited -ProcessId $ownedPid -Current $left) { $gone = $true; break }
                Start-Sleep -Milliseconds 1000
            }
            if (-not $gone) { Write-Log WARN "Owned Outlook (PID $ownedPid) still running ~30s after Quit; left as-is (no force-kill)." }
        }
        Remove-ComRef $targetRoot
        Remove-ComRef $ns
        Remove-ComRef $outlook
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    # Guard #13: classify the outcome so the result dialog says "Export incomplete" (with reasons)
    # whenever items went missing, sync stalled, or a cross-account copy was blocked - never just
    # "complete" on a partial capture.
    $outcome = Get-ExportOutcome -SourceItems $srcItems -CopiedItems $dstItems `
        -SyncStalled $syncStalled -CrossAccountBlocked $crossAccountBlocked

    return [pscustomobject]@{
        Success        = $success
        PstPath        = $PstPath
        StoresExported = $storesExported
        FoldersCopied  = $foldersCopied
        SourceItems    = $srcItems
        CopiedItems    = $dstItems
        Issues         = @($allIssues)
        Outcome        = $outcome
    }
}

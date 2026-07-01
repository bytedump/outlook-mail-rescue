#Requires -Version 5.1
<#
.SYNOPSIS
    Ensure a classic Outlook profile is signed in and the mailbox has finished its
    Cached Exchange Mode sync before a full export, so the PST is not partial.
.DESCRIPTION
    When classic Outlook opens a live account it re-downloads the whole mailbox from the
    server into the OST. Wait-OutlookSync polls the primary mailbox item count and waits
    until it stops growing (stabilizes) or a timeout elapses. If sign-in / MFA is
    pending, Outlook shows its own dialog and the count simply stays flat - the GUI
    surfaces a "waiting for sign-in" hint.

    The stabilization decision (Test-SyncStabilized) is a pure helper and unit-tested.
    This whole path requires classic Outlook + a live account, so the COM parts are
    validated manually.
#>

# True when the last $StableReadings samples are all equal (count has stopped moving).
# Pure + testable.
function Test-SyncStabilized {
    param([int[]]$Samples, [int]$StableReadings = 3)
    if ($null -eq $Samples -or $Samples.Count -lt $StableReadings) { return $false }
    $tail = $Samples[($Samples.Count - $StableReadings)..($Samples.Count - 1)]
    $first = $tail[0]
    foreach ($x in $tail) { if ($x -ne $first) { return $false } }
    return $true
}

function Test-OutlookProfileExists {
    $info = Get-OutlookInfo
    return $info.HasMapiProfile
}

# Sum the item counts of the primary Exchange mailbox's top-level folders. A cheap,
# decent signal for "is sync still pulling items". Returns 0 if no primary store.
function Get-PrimaryMailboxSampleCount {
    param([Parameter(Mandatory)]$Namespace)
    $total = 0
    foreach ($s in $Namespace.Stores) {
        $type = $null
        try { $type = [int]$s.ExchangeStoreType } catch { }
        if ($type -ne 0) { continue }   # olPrimaryExchangeMailbox
        try {
            $root = $s.GetRootFolder()
            foreach ($f in $root.Folders) { try { $total += [int]$f.Items.Count } catch { } }
        } catch { }
    }
    return $total
}

# Poll until the mailbox item count stabilizes or the timeout elapses.
# Returns { Stabilized; Stalled; LastCount; Elapsed; Samples }.
function Wait-OutlookSync {
    param(
        [Parameter(Mandatory)]$Namespace,
        [int]$TimeoutMinutes = 30,
        [int]$PollSeconds = 5,
        [int]$StableReadings = 3,
        [scriptblock]$OnProgress
    )
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $samples = [System.Collections.Generic.List[int]]::new()
    $start = Get-Date

    while ((Get-Date) -lt $deadline) {
        $count = Get-PrimaryMailboxSampleCount -Namespace $Namespace
        $samples.Add([int]$count)
        if ($OnProgress) { & $OnProgress 'sync' "Syncing mailbox... items so far: $count" }

        if (Test-SyncStabilized -Samples $samples.ToArray() -StableReadings $StableReadings) {
            return [pscustomobject]@{
                Stabilized = $true; Stalled = $false; LastCount = $count
                Elapsed = ((Get-Date) - $start); Samples = $samples.ToArray()
            }
        }
        Start-Sleep -Seconds $PollSeconds
    }

    $last = if ($samples.Count) { $samples[$samples.Count - 1] } else { 0 }
    Write-Log WARN "Sync wait timed out after $TimeoutMinutes min (last count=$last). Exporting what is cached."
    return [pscustomobject]@{
        Stabilized = $false; Stalled = $true; LastCount = $last
        Elapsed = ((Get-Date) - $start); Samples = $samples.ToArray()
    }
}

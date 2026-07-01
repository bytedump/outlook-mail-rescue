#Requires -Version 5.1
<#
.SYNOPSIS
    Whole-disk scan for Outlook data files (*.pst, *.ost) from C:\ root.
.DESCRIPTION
    A manual stack-based directory walk with per-directory try/catch so a single
    access-denied folder never aborts the whole scan, and reparse points (junctions /
    symlinks) are skipped to avoid loops. This deliberately avoids
    [IO.EnumerationOptions], which does NOT exist on .NET Framework (Windows
    PowerShell 5.1) - only on .NET Core / PowerShell 7.

    The scan runs in a separate ELEVATED helper process (Invoke-DiskScanHelper) so it
    can read all of C:\; the non-elevated GUI spawns it (Start-ElevatedScan) and reads
    progress/results back through JSON files (UIPI allows file IPC across integrity
    levels). Pure helpers (Get-DataFileType, Format-FileSize) are unit-tested.
#>

# True only when the current process is elevated (admin).
function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    return ([Security.Principal.WindowsPrincipal]$id).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# Classify a path by extension: 'PST' | 'OST' | $null. Pure.
function Get-DataFileType {
    param([string]$Path)
    switch ([System.IO.Path]::GetExtension($Path).ToLowerInvariant()) {
        '.pst'  { return 'PST' }
        '.ost'  { return 'OST' }
        default { return $null }
    }
}

# Human-readable size. Pure.
function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -lt 0) { return '0 B' }
    $units = 'B', 'KB', 'MB', 'GB', 'TB'
    $i = 0; $n = [double]$Bytes
    while ($n -ge 1024 -and $i -lt ($units.Count - 1)) { $n /= 1024; $i++ }
    $fmt = if ($i -eq 0) { '{0:0} {1}' } else { '{0:0.0} {1}' }
    return ($fmt -f $n, $units[$i])
}

# Is the file held open by another process (e.g. Outlook)? Best-effort, never throws.
function Test-FileLocked {
    param([string]$Path)
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
        $fs.Close(); $fs.Dispose()
        return $false
    } catch [System.IO.IOException] {
        return $true
    } catch {
        return $false  # access-denied etc. - not a lock
    }
}

# Build the record for one found data file.
function New-DataFileRecord {
    param([string]$Path)
    $size = -1L; $mtime = $null
    try {
        $fi = [System.IO.FileInfo]::new($Path)
        $size = $fi.Length
        # Format to a local wall-clock string at the source. The scan result crosses
        # to the GUI as JSON, and Windows PowerShell 5.1's ConvertTo-Json serializes a
        # DateTime as an epoch (\/Date(...)\/) that ConvertFrom-Json revives as UTC, so
        # a raw DateTime would display shifted by the local UTC offset. A string is
        # unambiguous and survives the round-trip untouched.
        $mtime = $fi.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
    } catch { }
    return [pscustomobject]@{
        Path         = $Path
        Type         = (Get-DataFileType $Path)
        SizeBytes    = $size
        Size         = (Format-FileSize $size)
        LastWriteTime = $mtime
        Locked       = (Test-FileLocked $Path)
    }
}

# Walk $Roots, collecting *.pst/*.ost. Tolerant of denied dirs, skips reparse points.
#   OnProgress: optional scriptblock invoked as & $OnProgress $dirCount $currentDir $foundCount
#   CancelCheck: optional scriptblock returning $true to abort early.
function Find-OutlookDataFile {
    param(
        [string[]]$Roots = @('C:\'),
        [scriptblock]$OnProgress,
        [scriptblock]$CancelCheck
    )
    $results = [System.Collections.Generic.List[object]]::new()
    $stack = [System.Collections.Generic.Stack[string]]::new()
    foreach ($r in $Roots) { $stack.Push($r) }
    $dirCount = 0

    while ($stack.Count -gt 0) {
        if ($CancelCheck -and (& $CancelCheck)) { break }
        $dir = $stack.Pop()

        # Files in this directory.
        try {
            foreach ($f in [System.IO.Directory]::EnumerateFiles($dir)) {
                if (Get-DataFileType $f) { $results.Add((New-DataFileRecord $f)) }
            }
        } catch { }

        # Subdirectories (skip reparse points to avoid junction/symlink loops).
        try {
            foreach ($d in [System.IO.Directory]::EnumerateDirectories($dir)) {
                try {
                    $attr = [System.IO.File]::GetAttributes($d)
                    if (($attr -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                } catch { continue }
                $stack.Push($d)
            }
        } catch { }

        $dirCount++
        if ($OnProgress -and ($dirCount % 250 -eq 0)) { & $OnProgress $dirCount $dir $results.Count }
    }

    if ($OnProgress) { & $OnProgress $dirCount '(done)' $results.Count }
    # Comma operator: return the List object itself so an empty result is still a
    # collection (count 0), not $null (PowerShell unwraps a returned empty list).
    return , $results
}

# --- Elevated helper mode (runs as admin child; talks to the GUI via JSON files) ---

# Atomic-ish JSON write: temp then rename, so the GUI never reads a half-written file.
function Write-JsonFile {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Object, [int]$Depth = 5)
    $tmp = "$Path.tmp"
    ($Object | ConvertTo-Json -Depth $Depth) | Set-Content -Path $tmp -Encoding UTF8
    Move-Item -Path $tmp -Destination $Path -Force
}

function Invoke-DiskScanHelper {
    param(
        [Parameter(Mandatory)][string]$ResultPath,
        [string]$ProgressPath,
        [string[]]$Roots = @('C:\')
    )
    $cancelPath = "$ResultPath.cancel"
    Write-Log INFO "Scan helper started (admin=$(Test-IsAdmin)) roots=$($Roots -join ',')"

    $onProgress = {
        param($dirCount, $current, $found)
        if ($ProgressPath) {
            Write-JsonFile -Path $ProgressPath -Object ([pscustomobject]@{
                DirCount = $dirCount; Current = $current; Found = $found; Done = $false
            })
        }
    }
    $cancelCheck = { Test-Path $cancelPath }

    $found = Find-OutlookDataFile -Roots $Roots -OnProgress $onProgress -CancelCheck $cancelCheck
    $cancelled = Test-Path $cancelPath

    Write-JsonFile -Path $ResultPath -Object ([pscustomobject]@{
        Cancelled = $cancelled
        Count     = $found.Count
        Files     = @($found)
    })
    if ($ProgressPath) {
        Write-JsonFile -Path $ProgressPath -Object ([pscustomobject]@{
            DirCount = -1; Current = '(done)'; Found = $found.Count; Done = $true
        })
    }
    Write-Log INFO "Scan helper finished: found=$($found.Count) cancelled=$cancelled"
}

# --- GUI side (non-elevated): spawn the elevated helper, read its output ---

# Spawn the elevated scan child. Returns the process + the IPC paths to poll.
function Start-ElevatedScan {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string[]]$Roots = @('C:\')
    )
    $base = Join-Path $env:TEMP ("mailrescue_scan_{0}" -f ([guid]::NewGuid().ToString('N')))
    $resultPath   = "$base.result.json"
    $progressPath = "$base.progress.json"

    $argList = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-WindowStyle', 'Hidden',
        '-File', $ScriptPath,
        '-ScanHelper',
        '-ScanResultPath', $resultPath,
        '-ScanProgressPath', $progressPath,
        '-ScanRoots'
    ) + $Roots

    # The helper talks to the GUI only through the result/progress JSON files, so its console
    # never needs to be visible. Hide it both at launch (-WindowStyle Hidden on Start-Process)
    # and from inside the child (-WindowStyle Hidden in the arg list) so no oversized elevated
    # PowerShell console flashes up during the scan. UAC still prompts (RunAs is on its own secure
    # desktop, unaffected by the window style).
    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -WindowStyle Hidden -PassThru
    return [pscustomobject]@{
        Process      = $proc
        ResultPath   = $resultPath
        ProgressPath = $progressPath
        CancelPath   = "$resultPath.cancel"
    }
}

# Signal the helper to stop (it checks this file between directories).
function Stop-ElevatedScan {
    param([Parameter(Mandatory)][string]$CancelPath)
    try { New-Item -ItemType File -Path $CancelPath -Force | Out-Null } catch { }
}

function Read-ScanProgress {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json) } catch { return $null }
}

# Read final results. Returns @{ Cancelled; Count; Files=@(...) } or $null if not ready.
function Read-ScanResult {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try {
        $obj = Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json
        # ConvertFrom-Json unwraps a single-element array; force an array for callers.
        $files = @(); if ($obj.PSObject.Properties.Name -contains 'Files' -and $obj.Files) { $files = @($obj.Files) }
        return [pscustomobject]@{ Cancelled = [bool]$obj.Cancelled; Count = [int]$obj.Count; Files = $files }
    } catch { return $null }
}

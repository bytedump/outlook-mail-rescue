#Requires -Version 5.1
<#
.SYNOPSIS
    Outlook Mail Rescue - find an existing PST, or do a full mailbox export to PST
    (classic Outlook via COM), detect/switch the new Outlook, and report orphaned OST.
.DESCRIPTION
    Single entry point with three modes:
      (default)    Launch the WinForms GUI a technician fills in. Non-elevated so it
                   can drive Outlook COM (an elevated host cannot - UIPI). Relaunches
                   itself in the PowerShell host whose bitness matches Outlook's.
      -ScanHelper  Internal: run the C:\ disk scan (needs admin), write results as
                   JSON to -ScanResultPath, then exit. Spawned elevated by the GUI.
      -LoadOnly    Test seam: dot-source the modules (defining their functions) and
                   return without launching anything. Used by tests/unit.
.EXAMPLE
    .\run.bat
    Normal use - launches the GUI.
.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\Invoke-MailRescue.ps1 -LoadOnly
    Dot-source for Pester.
#>
[CmdletBinding()]
param(
    [switch]$LoadOnly,
    [switch]$ScanHelper,
    [string]$ScanResultPath,
    [string]$ScanProgressPath,
    [string[]]$ScanRoots = @('C:\'),
    # Internal guard so the bitness relaunch cannot loop forever.
    [switch]$NoRelaunch
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Load modules (function definitions only; nothing runs on dot-source) ---
$srcDir = Join-Path $PSScriptRoot 'src'
foreach ($mod in 'Logging', 'OutlookDetect', 'Scan', 'NewOutlookToggle', 'ProfileSync', 'ComExport', 'Gui') {
    $p = Join-Path $srcDir "$mod.ps1"
    if (Test-Path $p) { . $p }
}

# Test seam: stop after the functions are defined.
if ($LoadOnly) { return }

# --- Elevated scan helper mode: scan, write JSON, exit ---
if ($ScanHelper) {
    Initialize-Logging -Tag 'scan' | Out-Null
    Invoke-DiskScanHelper -ResultPath $ScanResultPath -ProgressPath $ScanProgressPath -Roots $ScanRoots
    return
}

# --- Bitness self-relaunch: COM Interop needs the host bitness to match Outlook ---
# Only relevant for the GUI/COM path. Skip if already relaunched (NoRelaunch) or if
# Outlook bitness is unknown (not installed / undetectable).
if (-not $NoRelaunch) {
    $needPath = Get-MismatchedPowerShellHost   # returns a powershell.exe path, or $null
    if ($needPath) {
        $argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', $PSCommandPath, '-NoRelaunch')
        Start-Process -FilePath $needPath -ArgumentList $argList
        return
    }
}

Initialize-Logging -Tag 'gui' | Out-Null
Show-MailRescueGui -ScriptPath $PSCommandPath

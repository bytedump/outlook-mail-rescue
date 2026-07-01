#Requires -Version 5.1
<#
.SYNOPSIS
    Structured logging for Outlook Mail Rescue.
.DESCRIPTION
    Defines Write-Log, which writes a timestamped line to a log file, to the host,
    and (when a GUI is attached) to a thread-safe queue the WinForms timer drains
    into the on-screen log pane. Dot-sourced by Invoke-MailRescue.ps1; carries no
    UI dependency so it is safe to load in the elevated scan helper and in tests.
#>

# Thread-safe queue the GUI timer drains. Stays $null in headless/helper runs.
if (-not (Get-Variable -Name MrLogQueue -Scope Script -ErrorAction SilentlyContinue)) {
    $script:MrLogQueue = $null
}
# Accumulates ERROR/FATAL lines so the GUI can show a final summary.
if (-not (Get-Variable -Name MrErrors -Scope Script -ErrorAction SilentlyContinue)) {
    $script:MrErrors = [System.Collections.Generic.List[string]]::new()
}
if (-not (Get-Variable -Name MrLogFile -Scope Script -ErrorAction SilentlyContinue)) {
    $script:MrLogFile = $null
}

# Resolve a per-run log file. Defaults under %LOCALAPPDATA%\OutlookMailRescue\logs.
# Pass -Path to override (e.g. tests writing to a temp dir).
function Initialize-Logging {
    param([string]$Path, [string]$Tag = 'session')
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $dir = Join-Path $env:LOCALAPPDATA 'OutlookMailRescue\logs'
        $null = New-Item -ItemType Directory -Force -Path $dir -ErrorAction SilentlyContinue
        $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $Path = Join-Path $dir "mailrescue_${Tag}_${stamp}.log"
    }
    $script:MrLogFile = $Path
    return $Path
}

# Attach the GUI's thread-safe queue so log lines also surface on screen.
function Set-LogQueue {
    param([System.Collections.Concurrent.ConcurrentQueue[object]]$Queue)
    $script:MrLogQueue = $Queue
}

function Write-Log {
    param(
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'FATAL')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory)][string]$Message
    )
    $ts   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"

    if ($script:MrLogFile) {
        Add-Content -Path $script:MrLogFile -Value $line -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    Write-Host $line

    if ($Level -in 'ERROR', 'FATAL') { $script:MrErrors.Add($line) | Out-Null }

    if ($script:MrLogQueue) {
        try { $script:MrLogQueue.Enqueue([pscustomobject]@{ Level = $Level; Line = $line }) } catch { }
    }
}

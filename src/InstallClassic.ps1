#Requires -Version 5.1
<#
.SYNOPSIS
    Opt-in silent install of classic Outlook via the Office Deployment Tool (ODT).
.DESCRIPTION
    Used only when classic Outlook is absent and the technician explicitly confirms it.
    Heavy: large download, needs a valid Office/M365 license, admin, and network. We do
    NOT bundle or auto-download setup.exe (supply-chain / licensing) - the technician
    points the tool at an ODT setup.exe they already have. The config installs the
    Microsoft 365 Apps suite with everything except Outlook excluded.

    Format-OdtConfigXml is pure and unit-tested; the actual install is validated manually.
#>

# Build an ODT configuration XML that installs only Outlook. Pure.
function Format-OdtConfigXml {
    param(
        [ValidateSet('64', '32')][string]$Bitness = '64',
        [string]$ProductId = 'O365ProPlusRetail',
        [string]$Channel = 'Current',
        # Optional directory ODT writes its install log to; lets the GUI track progress.
        [string]$LogPath
    )
    $excludeApps = 'Access', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'PowerPoint', 'Publisher', 'Word', 'Teams', 'Bing'
    $excludeXml = ($excludeApps | ForEach-Object { "      <ExcludeApp ID=`"$_`" />" }) -join "`n"
    $loggingXml = if (-not [string]::IsNullOrWhiteSpace($LogPath)) { "`n  <Logging Level=`"Standard`" Path=`"$LogPath`" />" } else { '' }
    return @"
<Configuration>
  <!-- Adjust Product ID to match your licensing (e.g. O365ProPlusRetail / O365BusinessRetail). -->
  <Add OfficeClientEdition="$Bitness" Channel="$Channel">
    <Product ID="$ProductId">
      <Language ID="MatchOS" />
$excludeXml
    </Product>
  </Add>$loggingXml
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
</Configuration>
"@
}

# Write the ODT config to disk; returns the path.
function New-OdtConfigXml {
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('64', '32')][string]$Bitness = '64',
        [string]$ProductId = 'O365ProPlusRetail',
        [string]$Channel = 'Current',
        [string]$LogPath
    )
    $xml = Format-OdtConfigXml -Bitness $Bitness -ProductId $ProductId -Channel $Channel -LogPath $LogPath
    Set-Content -Path $Path -Value $xml -Encoding UTF8
    return $Path
}

function Test-OdtAvailable {
    param([string]$OdtSetupPath)
    return [bool]($OdtSetupPath -and (Test-Path $OdtSetupPath))
}

# Run ODT: setup.exe /configure config.xml. Requires admin. Returns { Success; ExitCode }.
function Install-ClassicOutlook {
    param(
        [Parameter(Mandatory)][string]$OdtSetupPath,
        [Parameter(Mandatory)][string]$ConfigPath,
        [scriptblock]$OnProgress
    )
    if (-not (Test-IsAdmin)) { throw 'Installing Outlook requires administrator rights.' }
    if (-not (Test-OdtAvailable $OdtSetupPath)) { throw "ODT setup.exe not found: $OdtSetupPath" }
    if (-not (Test-Path $ConfigPath)) { throw "ODT config not found: $ConfigPath" }

    Write-Log INFO "Installing classic Outlook via ODT: $OdtSetupPath /configure $ConfigPath"
    if ($OnProgress) { & $OnProgress 'install' 'Installing classic Outlook (this can take a while)...' }

    $proc = Start-Process -FilePath $OdtSetupPath -ArgumentList @('/configure', "`"$ConfigPath`"") -Wait -PassThru
    $ok = ($proc.ExitCode -eq 0)
    if ($ok) { Write-Log INFO 'ODT install finished (exit 0).' }
    else { Write-Log ERROR "ODT install failed (exit $($proc.ExitCode))." }

    return [pscustomobject]@{ Success = $ok; ExitCode = $proc.ExitCode }
}

# ---------------- Auto-install engine (winget-sourced ODT; validated manually) ----------------

# Derive a coarse install phase from the tail of the ODT log. Pure + unit-tested.
# The exact ODT log wording is refined against a real run; the keyword buckets keep
# the GUI honest (a phase + elapsed time) instead of faking a percentage.
function Get-OdtInstallPhase {
    param([string[]]$LogLines)
    if (-not $LogLines -or $LogLines.Count -eq 0) { return 'Preparing' }
    $text = ($LogLines -join "`n")
    if ($text -match '(?i)installation\s+succeeded|apply\s+succeeded|ending\s+operation') { return 'Finishing' }
    if ($text -match '(?i)apply\s+stage|installing|finalizing') { return 'Installing' }
    if ($text -match '(?i)download\s+stage|downloading|streaming') { return 'Downloading' }
    return 'Preparing'
}

# Fetch the Office Deployment Tool from the official winget source and self-extract its
# setup.exe into a controlled folder. Returns the setup.exe path or throws.
function Get-OdtSetup {
    param(
        [Parameter(Mandatory)][string]$WorkDir,
        [string]$WingetId = 'Microsoft.OfficeDeploymentTool',
        [scriptblock]$OnProgress
    )
    if ($OnProgress) { & $OnProgress 'Getting the Office Deployment Tool...' }
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $winget) { throw 'winget is not available to fetch the Office Deployment Tool.' }

    $dlDir = Join-Path $WorkDir 'odt-download'
    $exDir = Join-Path $WorkDir 'odt'
    New-Item -ItemType Directory -Force -Path $dlDir, $exDir | Out-Null

    $dlArgs = @('download', '--id', $WingetId, '--source', 'winget', '-d', $dlDir,
        '--accept-source-agreements', '--accept-package-agreements')
    $wp = Start-Process -FilePath $winget.Source -ArgumentList $dlArgs -Wait -PassThru -WindowStyle Hidden
    if ($wp.ExitCode -ne 0) { throw "winget failed to download the ODT (exit $($wp.ExitCode))." }

    $odtExe = Get-ChildItem -Path $dlDir -Filter '*.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $odtExe) { throw 'ODT installer .exe not found after winget download.' }

    # The ODT exe is a self-extractor: /extract:<dir> /quiet drops setup.exe + a sample config.
    $ep = Start-Process -FilePath $odtExe.FullName -ArgumentList @("/extract:$exDir", '/quiet') -Wait -PassThru
    if ($ep.ExitCode -ne 0) { throw "ODT self-extract failed (exit $($ep.ExitCode))." }

    $setup = Join-Path $exDir 'setup.exe'
    if (-not (Test-Path $setup)) { throw "ODT setup.exe not found after extract: $setup" }
    return $setup
}

# Run setup.exe /configure WITHOUT blocking, polling the ODT log directory so the caller
# can surface a live phase. Returns { Success; ExitCode }.
function Invoke-OdtConfigure {
    param(
        [Parameter(Mandatory)][string]$SetupPath,
        [Parameter(Mandatory)][string]$ConfigPath,
        [string]$LogDir,
        [scriptblock]$OnProgress,
        [int]$PollSeconds = 3
    )
    Write-Log INFO "Running ODT: $SetupPath /configure $ConfigPath"
    # Do NOT pass -WindowStyle Hidden: the ODT setup.exe bootstrap stalls indefinitely
    # when launched with a hidden window (it never hands off to Click-to-Run). Letting
    # it show its small bootstrap window lets the install proceed. The actual Office
    # install UI stays suppressed by Display Level="None" in the config.
    $proc = Start-Process -FilePath $SetupPath -ArgumentList @('/configure', "`"$ConfigPath`"") -PassThru
    while (-not $proc.HasExited) {
        $lines = @()
        if ($LogDir -and (Test-Path $LogDir)) {
            $logFile = Get-ChildItem -Path $LogDir -Filter '*.log' -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime | Select-Object -Last 1
            if ($logFile) { $lines = @(Get-Content -Path $logFile.FullName -Tail 60 -ErrorAction SilentlyContinue) }
        }
        if ($OnProgress) { & $OnProgress (Get-OdtInstallPhase -LogLines $lines) }
        Start-Sleep -Seconds $PollSeconds
    }
    $ok = ($proc.ExitCode -eq 0)
    if ($ok) { Write-Log INFO 'ODT configure finished (exit 0).' }
    else { Write-Log ERROR "ODT configure failed (exit $($proc.ExitCode))." }
    return [pscustomobject]@{ Success = $ok; ExitCode = $proc.ExitCode }
}

# Elevated-child entry point: get the ODT, write an Outlook-only config, run it, and
# report phase/elapsed (progress JSON) + a final result JSON the GUI polls. Mirrors
# Invoke-DiskScanHelper. Spawned by Start-ElevatedInstall.
function Invoke-ClassicInstallHelper {
    param(
        [Parameter(Mandatory)][string]$ResultPath,
        [string]$ProgressPath,
        [string]$ProductId = 'O365ProPlusRetail',
        [ValidateSet('64', '32')][string]$Bitness = '64'
    )
    $work = Join-Path $env:TEMP ("mailrescue_install_{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Force -Path $work | Out-Null
    $start = Get-Date
    Write-Log INFO "Install helper started (admin=$(Test-IsAdmin)) product=$ProductId bitness=$Bitness"

    $onProgress = {
        param($detail)
        if ($ProgressPath) {
            $el = (Get-Date) - $start
            Write-JsonFile -Path $ProgressPath -Object ([pscustomobject]@{
                Phase = $detail; Elapsed = ('{0:mm\:ss}' -f $el); Done = $false
            })
        }
    }

    $result = [pscustomobject]@{ Success = $false; ExitCode = $null; Message = '' }
    try {
        if (-not (Test-IsAdmin)) { throw 'Installing Outlook requires administrator rights.' }
        $setup = Get-OdtSetup -WorkDir $work -OnProgress $onProgress
        $logDir = Join-Path $work 'log'
        New-Item -ItemType Directory -Force -Path $logDir | Out-Null
        $cfg = New-OdtConfigXml -Path (Join-Path $work 'config.xml') -Bitness $Bitness -ProductId $ProductId -LogPath $logDir
        $r = Invoke-OdtConfigure -SetupPath $setup -ConfigPath $cfg -LogDir $logDir -OnProgress $onProgress
        $msg = if ($r.Success) { 'Classic Outlook installed.' } else { "ODT exited with code $($r.ExitCode)." }
        $result = [pscustomobject]@{ Success = $r.Success; ExitCode = $r.ExitCode; Message = $msg }
    } catch {
        Write-Log ERROR "Install helper failed: $($_.Exception.Message)"
        $result = [pscustomobject]@{ Success = $false; ExitCode = $null; Message = "$($_.Exception.Message)" }
    } finally {
        if ($ProgressPath) {
            Write-JsonFile -Path $ProgressPath -Object ([pscustomobject]@{
                Phase = 'Done'; Elapsed = ('{0:mm\:ss}' -f ((Get-Date) - $start)); Done = $true
            })
        }
        Write-JsonFile -Path $ResultPath -Object $result
        Write-Log INFO "Install helper finished: success=$($result.Success) msg='$($result.Message)'"
    }
}

# --- GUI side (non-elevated): spawn the elevated install child, read its output ---

# Spawn the elevated install child. Returns the process + the IPC paths to poll.
function Start-ElevatedInstall {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [string]$ProductId = 'O365ProPlusRetail',
        [ValidateSet('64', '32')][string]$Bitness = '64'
    )
    $base = Join-Path $env:TEMP ("mailrescue_install_{0}" -f ([guid]::NewGuid().ToString('N')))
    $resultPath = "$base.result.json"
    $progressPath = "$base.progress.json"

    $argList = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA',
        '-File', $ScriptPath,
        '-InstallHelper',
        '-InstallResultPath', $resultPath,
        '-InstallProgressPath', $progressPath,
        '-InstallProductId', $ProductId,
        '-InstallBitness', $Bitness
    )
    $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -PassThru
    return [pscustomobject]@{
        Process      = $proc
        ResultPath   = $resultPath
        ProgressPath = $progressPath
    }
}

# Read the final install result, or $null if not ready yet.
function Read-InstallResult {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) { return $null }
    try { return (Get-Content -Path $Path -Raw -ErrorAction Stop | ConvertFrom-Json) } catch { return $null }
}

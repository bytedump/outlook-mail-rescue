#Requires -Version 5.1
<#
.SYNOPSIS
    WinForms GUI for Outlook Mail Rescue: the technician fills in metadata, scans C:\,
    and runs the export decision tree.
.DESCRIPTION
    Two async mechanisms feed one UI timer (so the form never freezes):
      - the disk scan runs in a separate ELEVATED process; the timer polls its JSON
        progress/result files;
      - the COM export runs in a background STA runspace; the timer drains a shared
        log queue and reads a synchronized state hashtable.
    The GUI process itself stays non-elevated so Outlook COM works (UIPI).
#>

# Load config.ps1 if present, else config.example.ps1, else built-in defaults.
function Get-MailRescueConfig {
    param([string]$ScriptDir)
    foreach ($f in 'config.ps1', 'config.example.ps1') {
        $p = Join-Path $ScriptDir $f
        if (Test-Path $p) {
            try { return (& $p) } catch { Write-Log WARN "Failed to load ${f}: $($_.Exception.Message)" }
        }
    }
    return [ordered]@{
        OutputFolder = Join-Path $env:USERPROFILE 'Desktop\MailRescue'
        FileNameTemplate = '{username}_{stamp}.pst'
        ScanRoots = @('C:\'); IncludeArchive = $true; IncludeSharedMailboxes = $false
        IncludePublicFolders = $false; SyncTimeoutMinutes = 30; OstConverterPath = ''
    }
}

# Pure decision for the auto-detect prefill (guard #7): may an auto-detected value be
# written into the field, or has the technician taken it over? Returns $true only when the
# field is still "ours" - empty, showing just the cue, or holding the exact value we last
# auto-set (so a later, better detection can upgrade it: none -> displayname -> smtp). Any
# other content means the user typed/pasted something - never overwrite it. Comparison is
# case-sensitive: even a case change counts as a user edit.
function Test-CanAutoFill {
    param([string]$CurrentText, [string]$LastAutoSet, [string]$CueText)
    if ([string]::IsNullOrWhiteSpace($CurrentText)) { return $true }
    $ordinal = [System.StringComparison]::Ordinal
    if ($CueText -and [string]::Equals($CurrentText, $CueText, $ordinal)) { return $true }
    if ([string]::Equals($CurrentText, $LastAutoSet, $ordinal)) { return $true }
    return $false
}

# Pure check (guard #12): does $Path resolve at or under any OneDrive/KFM-synced root?
# Writing a multi-GB PST into a synced folder makes OneDrive upload it while Outlook still
# holds it open -> corruption / sync conflicts. The caller gathers the roots (env /
# registry) and warns or redirects before exporting. Boundary-safe (a sibling like
# '...\OneDriveBackup' is NOT under '...\OneDrive') and case-insensitive. Expects
# absolute, real folder paths (no '..' segments to resolve).
function Test-PathUnderOneDrive {
    param([string]$Path, [string[]]$OneDriveRoots)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not $OneDriveRoots) { return $false }

    $normalize = { param($p)
        if ([string]::IsNullOrWhiteSpace($p)) { return $null }
        return $p.Trim().Replace('/', '\').TrimEnd('\').ToLowerInvariant()
    }

    $target = & $normalize $Path
    if (-not $target) { return $false }

    foreach ($root in $OneDriveRoots) {
        $r = & $normalize $root
        if (-not $r) { continue }
        if ($target -eq $r -or $target.StartsWith("$r\", [System.StringComparison]::Ordinal)) {
            return $true
        }
    }
    return $false
}

# Collect this machine's OneDrive root folders (guard #12): the OneDrive* environment
# variables plus each account's UserFolder under HKCU (covers Known-Folder-Move/business +
# personal). Feeds Test-PathUnderOneDrive so we can warn before writing a multi-GB PST into
# a synced folder, where the sync client can lock or upload-corrupt it mid-write. Live read
# (env/registry), so it is validated by manual e2e; the matching logic above is unit-tested.
function Get-OneDriveRoots {
    $roots = New-Object System.Collections.Generic.List[string]
    foreach ($name in 'OneDrive', 'OneDriveCommercial', 'OneDriveConsumer') {
        $v = [Environment]::GetEnvironmentVariable($name)
        if ($v) { $roots.Add($v) }
    }
    try {
        $base = 'HKCU:\Software\Microsoft\OneDrive\Accounts'
        if (Test-Path $base) {
            foreach ($acct in Get-ChildItem $base -ErrorAction SilentlyContinue) {
                $p = Get-ItemProperty -Path $acct.PSPath -Name 'UserFolder' -ErrorAction SilentlyContinue
                if ($p -and $p.UserFolder) { $roots.Add([string]$p.UserFolder) }
            }
        }
    } catch { }
    return ($roots | Where-Object { $_ } | Select-Object -Unique)
}

# Start the COM export in a background STA runspace. Returns @{ Ps; Async }.
function Start-ExportRunspace {
    param(
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$PstPath,
        [Parameter(Mandatory)][hashtable]$ConfigHash,
        [Parameter(Mandatory)][hashtable]$Ui,
        [string]$LogFile,
        [bool]$WaitForSync = $true,
        [int]$SyncTimeoutMinutes = 30,
        [string]$ExpectedSmtp
    )
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($ScriptPath, $PstPath, $ConfigHash, $Ui, $LogFile, $WaitForSync, $SyncTimeoutMinutes, $ExpectedSmtp)
        . $ScriptPath -LoadOnly
        Set-LogQueue -Queue $Ui.LogQueue
        Set-Variable -Name MrLogFile -Scope Script -Value $LogFile
        $onProgress = { param($phase, $detail) $Ui.Phase = $phase; $Ui.Status = $detail }
        try {
            $res = Export-MailboxToPst -PstPath $PstPath -Config $ConfigHash `
                -WaitForSync:$WaitForSync -SyncTimeoutMinutes $SyncTimeoutMinutes `
                -ExpectedSmtp $ExpectedSmtp -OnProgress $onProgress
            $Ui.Result = $res
        } catch {
            $Ui.Result = [pscustomobject]@{ Success = $false; Issues = @("$($_.Exception.Message)"); PstPath = $PstPath }
        } finally {
            $Ui.Done = $true
        }
    })
    [void]$ps.AddArgument($ScriptPath).AddArgument($PstPath).AddArgument($ConfigHash)
    [void]$ps.AddArgument($Ui).AddArgument($LogFile).AddArgument($WaitForSync).AddArgument($SyncTimeoutMinutes).AddArgument($ExpectedSmtp)

    $async = $ps.BeginInvoke()
    return @{ Ps = $ps; Async = $async; Runspace = $rs }
}

# Run the identity probe in a background STA runspace so the COM/RPC-retry latency never
# freezes the UI (guard #3). Writes the Resolve-DetectedIdentity result to $Probe.Result
# and flips $Probe.Done; the UI timer drains and disposes it. Mirrors Start-ExportRunspace.
function Start-ProbeRunspace {
    param([Parameter(Mandatory)][string]$ScriptPath, [Parameter(Mandatory)][hashtable]$Probe, [string]$LogFile)
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions = 'ReuseThread'
    $rs.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript({
        param($ScriptPath, $Probe, $LogFile)
        . $ScriptPath -LoadOnly
        Set-Variable -Name MrLogFile -Scope Script -Value $LogFile
        try { $Probe.Result = Get-DetectedOwnerIdentity }
        catch { $Probe.Result = $null }
        finally { $Probe.Done = $true }
    })
    [void]$ps.AddArgument($ScriptPath).AddArgument($Probe).AddArgument($LogFile)

    $async = $ps.BeginInvoke()
    return @{ Ps = $ps; Async = $async; Runspace = $rs }
}

# Build the form, controls, handlers and timer; return the form WITHOUT running the
# message loop. Split out from Show-MailRescueGui so it can be constructed in a smoke
# test (construction catches most wiring errors; Application.Run needs a live desktop).
function New-MailRescueForm {
    param([Parameter(Mandatory)][string]$ScriptPath, [switch]$Run)

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    [System.Windows.Forms.Application]::EnableVisualStyles()

    # Native SendMessage for EM_SETCUEBANNER (grey placeholder in the empty Mailbox field,
    # guard #6b). Guarded so a re-entry never re-adds the type; purely cosmetic.
    # NB: Add-Type -MemberDefinition already injects `using System.Runtime.InteropServices;`
    # by default, so DllImport/CharSet resolve without -UsingNamespace; adding it duplicates
    # the using and fails under the PS 5.1 (csc) warning-as-error compiler.
    if (-not ('MailRescue.NativeMethods' -as [type])) {
        Add-Type -Namespace 'MailRescue' -Name 'NativeMethods' -MemberDefinition @'
[DllImport("user32.dll", CharSet = CharSet.Unicode)]
public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
'@
    }

    $scriptDir = Split-Path -Parent $ScriptPath
    $cfg = Get-MailRescueConfig -ScriptDir $scriptDir
    $info = Get-OutlookInfo

    # --- Shared, timer-drained state (script scope so all handlers see the same data) ---
    $script:MrUi = [hashtable]::Synchronized(@{
        LogQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        Status   = ''
        Result   = $null
        Done     = $false
    })
    Set-LogQueue -Queue $script:MrUi.LogQueue
    $script:MrScan   = @{ Active = $false; Info = $null }
    $script:MrExport = @{ Active = $false; Handle = $null; Revert = $null }
    # Classic Outlook the app auto-launched at startup (guard #7). OwnedPid is set only when we
    # started it ourselves; the app quits that one PID on FormClosing and leaves any pre-existing
    # Outlook alone.
    $script:MrOutlook = @{ OwnedPid = $null }
    # Probe result crosses the runspace boundary -> synchronized. LastAutoSet tracks the
    # value we prefilled so Test-CanAutoFill can upgrade it without clobbering typed text.
    $script:MrProbe = [hashtable]::Synchronized(@{
        Active = $false; Handle = $null; Done = $false; Result = $null; LastAutoSet = ''
        LastDetected = $null   # last identity object the probe returned (for the #4 confirm)
    })
    # Placeholder/cue shown in the empty Mailbox field; also the value Test-ValidExportName
    # (guard #6) treats as "not a real name" so it can never become a PST file name.
    $script:MrMailboxCue = 'example@example.com.br'

    # --- Theme (light, flat, modern) ---
    $clrBg       = [System.Drawing.Color]::FromArgb(245, 246, 248)   # window background
    $clrCard     = [System.Drawing.Color]::White                     # panels / inputs
    $clrAccent   = [System.Drawing.Color]::FromArgb(0, 103, 192)     # primary blue
    $clrAccentDk = [System.Drawing.Color]::FromArgb(0, 78, 150)      # primary hover
    $clrAccentLt = [System.Drawing.Color]::FromArgb(232, 241, 251)   # secondary hover wash
    $clrText     = [System.Drawing.Color]::FromArgb(32, 32, 32)      # primary text
    $clrMuted    = [System.Drawing.Color]::FromArgb(107, 107, 107)   # secondary text
    $clrBorder   = [System.Drawing.Color]::FromArgb(220, 222, 226)   # hairline border
    $clrWarn     = [System.Drawing.Color]::FromArgb(176, 0, 32)      # errors / warnings
    $clrWarnBg   = [System.Drawing.Color]::FromArgb(255, 244, 206)   # warning banner wash
    $clrZebra    = [System.Drawing.Color]::FromArgb(247, 248, 250)   # alternate list row

    $fontBase  = New-Object System.Drawing.Font('Segoe UI', 9)
    $fontH1    = New-Object System.Drawing.Font('Segoe UI', 15, [System.Drawing.FontStyle]::Bold)
    $fontBold  = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $fontSmall = New-Object System.Drawing.Font('Segoe UI', 8)

    # Flat single-line input (drops the sunken Fixed3D border that looked "dented").
    function New-FlatInput {
        param([int]$X, [int]$Y, [int]$W)
        $t = New-Object System.Windows.Forms.TextBox
        $t.BorderStyle = 'FixedSingle'
        $t.Location = New-Object System.Drawing.Point($X, $Y)
        $t.Size = New-Object System.Drawing.Size($W, 25)
        return $t
    }

    # Flat button + native hover. Primary = accent fill; secondary = outlined.
    function Set-FlatButton {
        param($Btn, $Back, $Fore, $Hover, $Border, [int]$BorderSize)
        $Btn.FlatStyle = 'Flat'
        $Btn.Font = $fontBold
        $Btn.BackColor = $Back; $Btn.ForeColor = $Fore
        $Btn.FlatAppearance.BorderSize = $BorderSize
        if ($BorderSize -gt 0) { $Btn.FlatAppearance.BorderColor = $Border }
        $Btn.FlatAppearance.MouseOverBackColor = $Hover
        $Btn.Cursor = [System.Windows.Forms.Cursors]::Hand
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Outlook Mail Rescue'
    $form.Size = New-Object System.Drawing.Size(920, 790)
    $form.StartPosition = 'CenterScreen'
    $form.Font = $fontBase
    $form.BackColor = $clrBg
    $form.ForeColor = $clrText
    $form.MinimumSize = New-Object System.Drawing.Size(860, 700)

    # --- Header ---
    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Outlook Mail Rescue'
    $title.Font = $fontH1; $title.ForeColor = $clrAccent
    $title.Location = New-Object System.Drawing.Point(24, 16)
    $title.AutoSize = $true
    $form.Controls.Add($title)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Find or export a user's mailbox to a single PST file."
    $subtitle.Font = $fontBase; $subtitle.ForeColor = $clrMuted
    $subtitle.Location = New-Object System.Drawing.Point(26, 50)
    $subtitle.AutoSize = $true
    $form.Controls.Add($subtitle)

    # --- Golden-rule banner (soft warning card with an accent strip) ---
    $banner = New-Object System.Windows.Forms.Panel
    $banner.Location = New-Object System.Drawing.Point(24, 80)
    $banner.Size = New-Object System.Drawing.Size(872, 54)
    $banner.BackColor = $clrWarnBg
    $banner.Anchor = 'Top,Left,Right'
    $form.Controls.Add($banner)

    $bannerBar = New-Object System.Windows.Forms.Panel
    $bannerBar.Size = New-Object System.Drawing.Size(4, 54)
    $bannerBar.BackColor = $clrWarn
    $bannerBar.Dock = 'Left'
    $banner.Controls.Add($bannerBar)

    $bannerLbl = New-Object System.Windows.Forms.Label
    $bannerLbl.Text = 'Golden rule: run this BEFORE the account is disabled. A fresh export needs the live mailbox; once the account is gone, only local PST/OST files remain.'
    $bannerLbl.ForeColor = $clrText; $bannerLbl.Font = $fontBold
    $bannerLbl.Dock = 'Fill'; $bannerLbl.TextAlign = 'MiddleLeft'
    $bannerLbl.Padding = New-Object System.Windows.Forms.Padding(12, 0, 12, 0)
    $banner.Controls.Add($bannerLbl)

    # --- Case details card ---
    $card = New-Object System.Windows.Forms.Panel
    $card.Location = New-Object System.Drawing.Point(24, 146)
    $card.Size = New-Object System.Drawing.Size(872, 110)
    $card.BackColor = $clrCard; $card.BorderStyle = 'FixedSingle'
    $card.Anchor = 'Top,Left,Right'
    $form.Controls.Add($card)

    $cardHdr = New-Object System.Windows.Forms.Label
    $cardHdr.Text = 'Case details'; $cardHdr.Font = $fontBold; $cardHdr.ForeColor = $clrMuted
    $cardHdr.Location = New-Object System.Drawing.Point(14, 10); $cardHdr.AutoSize = $true
    $card.Controls.Add($cardHdr)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text = 'Mailbox / username'; $lblUser.ForeColor = $clrText
    $lblUser.Location = New-Object System.Drawing.Point(14, 36); $lblUser.AutoSize = $true
    $card.Controls.Add($lblUser)
    $txtUser = New-FlatInput 14 56 360
    $card.Controls.Add($txtUser)
    # Re-evaluate the Export gate (guard #6) on every change, including programmatic
    # prefill from the probe; keeps Export disabled until a usable identity is present.
    $txtUser.Add_TextChanged({ Update-ExportEnabled })
    $btnDetect = New-Object System.Windows.Forms.Button
    $btnDetect.Text = 'Detect owner'; $btnDetect.Location = New-Object System.Drawing.Point(380, 55)
    $btnDetect.Size = New-Object System.Drawing.Size(86, 26)
    Set-FlatButton $btnDetect $clrCard $clrAccent $clrAccentLt $clrBorder 1
    $card.Controls.Add($btnDetect)
    $btnDetect.Add_Click({ Start-IdentityProbe })
    $lblUserHint = New-Object System.Windows.Forms.Label
    $lblUserHint.Text = 'Auto-detected from classic Outlook when available; used to name the PST.'
    $lblUserHint.Font = $fontSmall; $lblUserHint.ForeColor = $clrMuted
    $lblUserHint.Location = New-Object System.Drawing.Point(14, 84); $lblUserHint.AutoSize = $true
    $card.Controls.Add($lblUserHint)

    $lblOut = New-Object System.Windows.Forms.Label
    $lblOut.Text = 'Output folder'; $lblOut.ForeColor = $clrText
    $lblOut.Location = New-Object System.Drawing.Point(470, 36); $lblOut.AutoSize = $true
    $card.Controls.Add($lblOut)
    $txtOut = New-FlatInput 470 56 320
    $txtOut.Text = [string]$cfg.OutputFolder
    $card.Controls.Add($txtOut)
    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = 'Browse'; $btnBrowse.Location = New-Object System.Drawing.Point(796, 55)
    $btnBrowse.Size = New-Object System.Drawing.Size(62, 27)
    Set-FlatButton $btnBrowse $clrCard $clrAccent $clrAccentLt $clrBorder 1
    $card.Controls.Add($btnBrowse)
    $btnBrowse.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq 'OK') { $txtOut.Text = $dlg.SelectedPath }
    })

    # --- Outlook status ---
    $lblStatus = New-Object System.Windows.Forms.Label
    $activeTxt  = (Get-Culture).TextInfo.ToTitleCase([string]$info.ActiveFlavor)
    $classicTxt = if ($info.ClassicInstalled) { "Classic: Yes ($($info.ClassicBitness))" } else { 'Classic: Not installed' }
    $newTxt     = if ($info.NewInstalled) { 'New: Yes' } else { 'New: No' }
    $mapiTxt    = if ($info.HasMapiProfile) { 'Yes' } else { 'No' }
    $lblStatus.Text = "Detected   |   Active: $activeTxt   |   $classicTxt   |   $newTxt   |   MAPI profile: $mapiTxt"
    $lblStatus.Location = New-Object System.Drawing.Point(24, 266); $lblStatus.AutoSize = $true
    $lblStatus.Font = $fontBold
    $lblStatus.ForeColor = if ($info.ClassicInstalled) { $clrAccent } else { $clrWarn }
    $form.Controls.Add($lblStatus)

    # --- Action buttons ---
    # Only scan (secondary/outlined): elevated C:\ sweep for existing PST/OST, no export.
    $btnScan = New-Object System.Windows.Forms.Button
    $btnScan.Text = 'Only scan'; $btnScan.Location = New-Object System.Drawing.Point(24, 294)
    $btnScan.Size = New-Object System.Drawing.Size(120, 34)
    Set-FlatButton $btnScan $clrCard $clrAccent $clrAccentLt $clrBorder 1
    $form.Controls.Add($btnScan)

    $btnScanExport = New-Object System.Windows.Forms.Button
    # UseMnemonic off so the literal '&' renders; otherwise WinForms eats it as an Alt-key marker
    # and the caption shows as "Scan  Export".
    $btnScanExport.UseMnemonic = $false
    $btnScanExport.Text = 'Scan & Export'; $btnScanExport.Location = New-Object System.Drawing.Point(154, 294)
    $btnScanExport.Size = New-Object System.Drawing.Size(200, 34)
    Set-FlatButton $btnScanExport $clrAccent ([System.Drawing.Color]::White) $clrAccentDk $clrAccent 0
    $form.Controls.Add($btnScanExport)
    $btnScanExport.Enabled = $false   # guard #6: enabled only once Test-ValidExportName passes

    $btnCopy = New-Object System.Windows.Forms.Button
    $btnCopy.Text = 'Copy selected PST'; $btnCopy.Location = New-Object System.Drawing.Point(364, 294)
    $btnCopy.Size = New-Object System.Drawing.Size(160, 34)
    Set-FlatButton $btnCopy $clrCard $clrAccent $clrAccentLt $clrBorder 1
    $form.Controls.Add($btnCopy)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel scan'; $btnCancel.Location = New-Object System.Drawing.Point(534, 294)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 34); $btnCancel.Enabled = $false
    Set-FlatButton $btnCancel $clrCard $clrMuted $clrAccentLt $clrBorder 1
    $form.Controls.Add($btnCancel)

    # --- Results list ---
    $list = New-Object System.Windows.Forms.ListView
    $list.View = 'Details'; $list.FullRowSelect = $true; $list.GridLines = $false; $list.CheckBoxes = $false
    $list.BorderStyle = 'FixedSingle'; $list.BackColor = $clrCard
    $list.Location = New-Object System.Drawing.Point(24, 336)
    $list.Size = New-Object System.Drawing.Size(872, 200)
    $list.Anchor = 'Top,Left,Right'
    [void]$list.Columns.Add('Type', 50)
    [void]$list.Columns.Add('Size', 80)
    [void]$list.Columns.Add('Modified', 140)
    [void]$list.Columns.Add('Locked', 60)
    [void]$list.Columns.Add('Path', 530)
    $form.Controls.Add($list)

    # --- Progress + status line ---
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Location = New-Object System.Drawing.Point(24, 546); $bar.Size = New-Object System.Drawing.Size(872, 14)
    $bar.Anchor = 'Top,Left,Right'; $bar.Style = 'Blocks'
    $form.Controls.Add($bar)

    $lblLine = New-Object System.Windows.Forms.Label
    $lblLine.Text = 'Ready.'; $lblLine.Location = New-Object System.Drawing.Point(24, 566); $lblLine.AutoSize = $true
    $lblLine.ForeColor = $clrMuted
    $form.Controls.Add($lblLine)

    # --- Log pane ---
    $log = New-Object System.Windows.Forms.RichTextBox
    $log.Location = New-Object System.Drawing.Point(24, 590); $log.Size = New-Object System.Drawing.Size(872, 148)
    $log.Anchor = 'Top,Bottom,Left,Right'; $log.ReadOnly = $true; $log.BorderStyle = 'FixedSingle'
    $log.Font = New-Object System.Drawing.Font('Consolas', 9); $log.BackColor = $clrCard
    $form.Controls.Add($log)

    function Add-LogLine {
        param($Level, $Line)
        $color = switch ($Level) {
            'ERROR' { $clrWarn } 'FATAL' { $clrWarn } 'WARN' { [System.Drawing.Color]::DarkOrange }
            default { [System.Drawing.Color]::Black }
        }
        $log.SelectionStart = $log.TextLength; $log.SelectionLength = 0
        $log.SelectionColor = $color
        $log.AppendText("$Line`n")
        $log.SelectionColor = $log.ForeColor
        $log.ScrollToCaret()
    }

    function Set-Busy { param([bool]$Busy, [string]$Style = 'Marquee')
        $btnScan.Enabled = -not $Busy; $btnCopy.Enabled = -not $Busy
        # Export re-gates on the identity (guard #6) when idle - a bare scan can leave the
        # Mailbox field empty/at the cue, so it must not blindly re-enable Export.
        if ($Busy) { $btnScanExport.Enabled = $false } else { Update-ExportEnabled }
        $bar.Style = if ($Busy) { $Style } else { 'Blocks' }
        if (-not $Busy) { $bar.Value = 0 }
    }

    function Get-SelectedRecord {
        if ($list.SelectedItems.Count -eq 0) { return $null }
        return $list.SelectedItems[0].Tag
    }

    # Launch the owner-identity probe (guard #5). Serialized against itself and the export
    # (guard #2): never two COM automations on one profile at once. The timer finalizes it.
    function Start-IdentityProbe {
        if ($script:MrProbe.Active -or $script:MrExport.Active) { return }
        $script:MrProbe.Done = $false; $script:MrProbe.Result = $null; $script:MrProbe.Active = $true
        $btnDetect.Enabled = $false
        $lblLine.Text = 'Detecting owner from classic Outlook...'
        $script:MrProbe.Handle = Start-ProbeRunspace -ScriptPath $ScriptPath -Probe $script:MrProbe -LogFile $script:MrLogFile
    }

    # Gate the Export button on a real, usable identity in the Mailbox field (guard #6).
    # Pure Test-ValidExportName decides; Export stays disabled until the name would yield a
    # valid PST file name, never the old 'user' fallback or the cue placeholder.
    function Update-ExportEnabled {
        $btnScanExport.Enabled = [bool](Test-ValidExportName -Name $txtUser.Text -CueText $script:MrMailboxCue)
    }

    $btnCancel.Add_Click({
        if ($script:MrScan.Active -and $script:MrScan.Info) {
            Stop-ElevatedScan -CancelPath $script:MrScan.Info.CancelPath
            $lblLine.Text = 'Cancelling scan...'
        }
    })

    # --- Only scan: elevated C:\ sweep for existing PST/OST, no export, no confirmation ---
    $btnScan.Add_Click({
        if ($script:MrScan.Active -or $script:MrExport.Active) { return }
        $roots = @($cfg.ScanRoots); if (-not $roots) { $roots = @('C:\') }
        Write-Log INFO "Requesting elevated scan of: $($roots -join ', ')"
        try {
            $scanInfo = Start-ElevatedScan -ScriptPath $ScriptPath -Roots $roots
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Elevation was cancelled or failed:`n$($_.Exception.Message)", 'Only scan', 'OK', 'Warning') | Out-Null
            return
        }
        $list.Items.Clear()
        $script:MrScan.Info = $scanInfo; $script:MrScan.Active = $true
        $script:MrScan.StartTime = Get-Date
        $btnCancel.Enabled = $true
        Set-Busy $true 'Marquee'
        $lblLine.Text = 'Scanning C:\ (elevated)...'
    })

    # --- Scan & Export: one confirmation, then a fresh dated export + a disk scan run together ---
    $btnScanExport.Add_Click({
        if ($script:MrScan.Active -or $script:MrExport.Active) { return }
        if ($script:MrProbe.Active) { [System.Windows.Forms.MessageBox]::Show('Owner detection is still running. Wait for it to finish.', 'Scan & Export', 'OK', 'Warning') | Out-Null; return }
        $user = $txtUser.Text.Trim()
        if (-not (Test-ValidExportName -Name $user -CueText $script:MrMailboxCue)) {
            [System.Windows.Forms.MessageBox]::Show('Enter or detect a valid mailbox identity first.', 'Scan & Export', 'OK', 'Warning') | Out-Null; return
        }
        $outDir = $txtOut.Text.Trim()
        if (-not $outDir) { [System.Windows.Forms.MessageBox]::Show('Choose an output folder.', 'Scan & Export', 'OK', 'Warning') | Out-Null; return }

        $live = Get-OutlookInfo
        if (-not $live.ClassicInstalled) {
            [System.Windows.Forms.MessageBox]::Show(
                "Classic Outlook is not installed on this machine, so a client-side export is not possible.`n`nOptions: install classic Outlook (Office Deployment Tool) and retry, or perform a server-side export.",
                'Classic Outlook required', 'OK', 'Error') | Out-Null
            return
        }

        # Name the PST from the identity + today's date: owner@company.DD-MM-YYYY.pst. A same-day
        # re-run gets an extra time suffix so an earlier backup is never overwritten (guard #10).
        $stamp = Get-Date -Format 'dd-MM-yyyy'
        $fileName = Get-DatedPstFileName -Identity $user -DateStamp $stamp
        if (-not $fileName) {
            [System.Windows.Forms.MessageBox]::Show("Could not derive a PST file name from '$user'. Enter a valid mailbox identity.", 'Scan & Export', 'OK', 'Warning') | Out-Null
            return
        }
        $pstPath = Join-Path $outDir $fileName
        if (Test-Path $pstPath) {
            $fileName = Get-DatedPstFileName -Identity $user -DateStamp ("$stamp-" + (Get-Date -Format 'HHmmss'))
            $pstPath = Join-Path $outDir $fileName
        }

        # Guard #4: the single identity confirmation before any export. Echo owner + mailbox +
        # machine + target so the tech catches a wrong profile or wrong PC before a multi-GB write.
        # Fold the OneDrive (guard #12) and new->classic notices in here so the run needs no further
        # dialogs. The export runspace independently re-reads CurrentUser once Outlook is logged on
        # and re-checks this identity (Compare-ReadbackIdentity) before copying.
        $detected = $script:MrProbe.LastDetected
        $ownerName = if ($detected -and $detected.PrimarySmtp -and
            [string]::Equals([string]$detected.PrimarySmtp, $user, [System.StringComparison]::OrdinalIgnoreCase) -and
            $detected.DisplayName) { [string]$detected.DisplayName } else { '(entered manually)' }
        $switching = ($live.ActiveFlavor -eq 'new' -or $live.UseNewOutlook -eq 1)
        $msg = "Scan this PC and export the mailbox to a fresh PST backup?`n`nOwner:    $ownerName`nMailbox:  $user`nComputer: $env:COMPUTERNAME`nSave to:  $pstPath`n`nThe entire classic-Outlook profile on this machine will be copied, and C:\ will be scanned for existing PST/OST files."
        if (Test-PathUnderOneDrive -Path $outDir -OneDriveRoots (Get-OneDriveRoots)) {
            $msg += "`n`nNOTE: the output folder is inside OneDrive. Sync can lock or corrupt a large PST while it is written - a local folder (e.g. C:\Exports) is safer."
        }
        if ($switching) {
            $msg += "`n`nNOTE: the user is on new Outlook. It will be switched to classic for the export (backed up; a revert is offered at the end)."
        }
        $confirm = [System.Windows.Forms.MessageBox]::Show($msg, 'Confirm Scan & Export', 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }

        # From here on: no more app dialogs until the run finishes - only the OS UAC prompt for the
        # elevated scan and any Outlook sign-in prompt remain, which cannot be suppressed.

        # Auto-switch new -> classic (backed up; the timer offers a revert when the export ends).
        $script:MrExport.Revert = $null
        if ($switching) { $script:MrExport.Revert = Switch-ToClassicOutlook }

        # Guard #4b: only enforce the in-runspace identity re-check when the field equals the
        # auto-detected primary SMTP; a manually typed (non-SMTP) name has nothing reliable to
        # compare, so pass empty to SKIP rather than false-abort.
        $expectedSmtp = ''
        if ($detected -and $detected.PrimarySmtp -and
            [string]::Equals([string]$detected.PrimarySmtp, $user, [System.StringComparison]::OrdinalIgnoreCase)) {
            $expectedSmtp = [string]$detected.PrimarySmtp
        }
        $cfgHash = @{
            IncludeArchive = [bool]$cfg.IncludeArchive
            IncludeSharedMailboxes = [bool]$cfg.IncludeSharedMailboxes
            IncludePublicFolders = [bool]$cfg.IncludePublicFolders
        }
        Write-Log INFO "Case: technician='$env:USERNAME' user='$user'"
        Write-Log INFO "Starting export -> $pstPath"

        # Export is the primary deliverable: start it first so a cancelled scan UAC never aborts it.
        $script:MrExport.Handle = Start-ExportRunspace -ScriptPath $ScriptPath -PstPath $pstPath `
            -ConfigHash $cfgHash -Ui $script:MrUi -LogFile $script:MrLogFile `
            -WaitForSync $true -SyncTimeoutMinutes ([int]$cfg.SyncTimeoutMinutes) `
            -ExpectedSmtp $expectedSmtp
        $script:MrExport.Active = $true

        # Then kick off the disk scan. If elevation is cancelled or fails, log and continue - the
        # export still runs; the scan is auxiliary (it only finds already-existing PST/OST files).
        $roots = @($cfg.ScanRoots); if (-not $roots) { $roots = @('C:\') }
        try {
            Write-Log INFO "Requesting elevated scan of: $($roots -join ', ')"
            $scanInfo = Start-ElevatedScan -ScriptPath $ScriptPath -Roots $roots
            $list.Items.Clear()
            $script:MrScan.Info = $scanInfo; $script:MrScan.Active = $true
            $script:MrScan.StartTime = Get-Date
            $btnCancel.Enabled = $true
        } catch {
            Write-Log WARN "Disk scan skipped (elevation cancelled or failed): $($_.Exception.Message). Export continues."
        }

        Set-Busy $true 'Marquee'
        $lblLine.Text = if ($script:MrScan.Active) { 'Scanning + exporting... (sign in if Outlook prompts)' } else { 'Exporting mailbox... (sign in if Outlook prompts)' }
    })

    # --- Copy a found PST ---
    $btnCopy.Add_Click({
        $rec = Get-SelectedRecord
        if (-not $rec -or $rec.Type -ne 'PST') { [System.Windows.Forms.MessageBox]::Show('Select a PST row first.', 'Copy', 'OK', 'Warning') | Out-Null; return }
        $outDir = $txtOut.Text.Trim()
        if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }
        $dest = Join-Path $outDir ([System.IO.Path]::GetFileName($rec.Path))
        try {
            $form.Cursor = 'WaitCursor'; $lblLine.Text = "Copying $($rec.Path)..."
            Copy-Item -Path $rec.Path -Destination $dest -Force
            Write-Log INFO "Copied PST to $dest"
            [System.Windows.Forms.MessageBox]::Show("Copied to:`n$dest", 'Copy', 'OK', 'Information') | Out-Null
        } catch {
            Write-Log ERROR "Copy failed: $($_.Exception.Message)"
            [System.Windows.Forms.MessageBox]::Show("Copy failed:`n$($_.Exception.Message)", 'Copy', 'OK', 'Error') | Out-Null
        } finally { $form.Cursor = 'Default'; $lblLine.Text = 'Ready.' }
    })

    # --- One timer drives both the scan poll and the export finalize + log drain ---
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 300
    $timer.Add_Tick({
        # Drain log queue.
        $item = $null
        while ($script:MrUi.LogQueue.TryDequeue([ref]$item)) { Add-LogLine $item.Level $item.Line }

        # Identity probe finalize + prefill (guard #5/#7).
        if ($script:MrProbe.Active -and $script:MrProbe.Done) {
            $script:MrProbe.Active = $false
            $hp = $script:MrProbe.Handle
            if ($hp) {
                try { $hp.Ps.EndInvoke($hp.Async) } catch { }
                try { $hp.Runspace.Close() } catch { }
                try { $hp.Ps.Dispose() } catch { }
            }
            $script:MrProbe.Handle = $null
            $btnDetect.Enabled = $true
            $res = $script:MrProbe.Result
            if ($res -and $res.HasIdentity) {
                $script:MrProbe.LastDetected = $res   # retain for the #4 confirm dialog
                if (Test-CanAutoFill -CurrentText $txtUser.Text -LastAutoSet $script:MrProbe.LastAutoSet) {
                    $txtUser.Text = $res.NameSourceText
                    $script:MrProbe.LastAutoSet = $res.NameSourceText
                }
                $lblLine.Text = "Detected owner: $($res.NameSourceText)  [$($res.Tier)]"
                Write-Log INFO "Identity detected: $($res.NameSourceText) (tier=$($res.Tier))"
            } else {
                $lblLine.Text = 'No owner auto-detected - enter the mailbox/username manually.'
            }
            $script:MrProbe.Done = $false; $script:MrProbe.Result = $null
        }

        # Export status + finalize.
        if ($script:MrExport.Active) {
            if ($script:MrUi.Status) { $lblLine.Text = $script:MrUi.Status }
            if ($script:MrUi.Done) {
                $script:MrExport.Active = $false
                $res = $script:MrUi.Result
                $h = $script:MrExport.Handle
                if ($h) {
                    try { $h.Ps.EndInvoke($h.Async) } catch { }
                    try { $h.Runspace.Close() } catch { }
                    try { $h.Ps.Dispose() } catch { }
                }
                if (-not ($script:MrScan.Active -or $script:MrExport.Active)) { Set-Busy $false }
                if ($res -and $res.Success) {
                    # Guard #13: title from the classified outcome - "Export complete" only when the
                    # whole mailbox was captured, else "Export incomplete" + the specific reasons, so
                    # a technician never offboards a mailbox that was not fully backed up.
                    $outcome = if ($res.PSObject.Properties.Name -contains 'Outcome') { $res.Outcome } else { $null }
                    $title = if ($outcome) { $outcome.Title } else { 'Export complete' }
                    $icon  = if ($outcome -and $outcome.Degraded) { 'Warning' } else { 'Information' }
                    $msg = "$title.`n`nPST: $($res.PstPath)`nStores: $($res.StoresExported)  Folders: $($res.FoldersCopied)`nItems: $($res.CopiedItems)/$($res.SourceItems)"
                    if ($outcome -and $outcome.Reasons -and $outcome.Reasons.Count) { $msg += "`n`nWhy incomplete:`n - " + ($outcome.Reasons -join "`n - ") }
                    if ($res.Issues -and $res.Issues.Count) { $msg += "`n`nIssues:`n - " + ($res.Issues -join "`n - ") }
                    [System.Windows.Forms.MessageBox]::Show($msg, 'Export', 'OK', $icon) | Out-Null
                } else {
                    $iss = if ($res) { ($res.Issues -join "`n - ") } else { 'unknown error' }
                    [System.Windows.Forms.MessageBox]::Show("Export failed:`n - $iss", 'Export', 'OK', 'Error') | Out-Null
                }
                # Offer to revert the new->classic switch.
                if ($script:MrExport.Revert) {
                    $ans = [System.Windows.Forms.MessageBox]::Show('Restore the user''s new-Outlook setting?', 'Revert', 'YesNo', 'Question')
                    if ($ans -eq 'Yes') { Restore-RegistryValue -Backup $script:MrExport.Revert }
                    $script:MrExport.Revert = $null
                }
                $script:MrUi.Done = $false; $script:MrUi.Result = $null; $script:MrUi.Status = ''
                $lblLine.Text = 'Ready.'
            }
        }

        # Scan poll + finalize.
        if ($script:MrScan.Active -and $script:MrScan.Info) {
            $si = $script:MrScan.Info
            $elapsed = if ($script:MrScan.StartTime) { (Get-Date) - $script:MrScan.StartTime } else { [TimeSpan]::Zero }
            $elTxt = '{0:mm\:ss}' -f $elapsed
            $prog = Read-ScanProgress -Path $si.ProgressPath
            if ($prog) { $lblLine.Text = "Scanning... $elTxt elapsed  |  Dirs: $($prog.DirCount)  |  Found: $($prog.Found)" }
            else { $lblLine.Text = "Scanning... $elTxt elapsed" }
            if (Test-Path $si.ResultPath) {
                $result = Read-ScanResult -Path $si.ResultPath
                if ($result) {
                    foreach ($f in $result.Files) {
                        $it = New-Object System.Windows.Forms.ListViewItem([string]$f.Type)
                        [void]$it.SubItems.Add([string]$f.Size)
                        [void]$it.SubItems.Add(("{0}" -f $f.LastWriteTime))
                        [void]$it.SubItems.Add($(if ($f.Locked) { 'Yes' } else { '' }))
                        [void]$it.SubItems.Add([string]$f.Path)
                        $it.Tag = $f
                        if ($list.Items.Count % 2 -eq 1) { $it.BackColor = $clrZebra }
                        [void]$list.Items.Add($it)
                    }
                    Write-Log INFO "Scan finished: $($result.Count) file(s) found (cancelled=$($result.Cancelled))."
                    $lblLine.Text = "Scan done: $($result.Count) file(s) found in $elTxt."
                }
                $script:MrScan.Active = $false; $btnCancel.Enabled = $false
                if (-not ($script:MrScan.Active -or $script:MrExport.Active)) { Set-Busy $false }
            } elseif ($si.Process -and $si.Process.HasExited) {
                $script:MrScan.Active = $false; $btnCancel.Enabled = $false
                if (-not ($script:MrScan.Active -or $script:MrExport.Active)) { Set-Busy $false }
                $lblLine.Text = 'Scan ended without results.'
            }
        }
    })

    $form.Add_Shown({
        Write-Log INFO 'Outlook Mail Rescue GUI started.'
        $timer.Start()
        Update-ExportEnabled   # guard #6: reflect the initial field state on the button
        # Grey placeholder in the empty Mailbox field (guard #6b). Non-fatal: a failure here
        # is cosmetic and must never block the GUI. wParam=1 keeps the cue visible on focus.
        try {
            [void][MailRescue.NativeMethods]::SendMessage($txtUser.Handle, 0x1501, [IntPtr]1, $script:MrMailboxCue)
        } catch { Write-Log WARN "Could not set Mailbox cue banner: $($_.Exception.Message)" }
        # Guard #7 auto-launch (overrides the old attach-only gate): if classic Outlook is already
        # running, attach + probe. If it is installed but NOT running, launch it and KEEP it open
        # (we own that PID and quit it on FormClosing), then probe the now-running instance. If
        # only the new Outlook is present we cannot launch classic - fall back to "Detect owner".
        if ($info.ClassicRunning) {
            Write-Log INFO 'Startup: classic Outlook already running; attach + probe (own nothing).'
            Start-IdentityProbe
        } elseif ($info.ClassicInstalled -and $info.ClassicPath) {
            Write-Log INFO 'Startup: classic Outlook not running; auto-launching (#7).'
            $lblLine.Text = 'Opening classic Outlook...'
            $script:MrOutlook.OwnedPid = Start-OwnedOutlook -ClassicPath $info.ClassicPath
            Start-IdentityProbe
        } else {
            Write-Log INFO 'Startup: classic Outlook not available; waiting for manual "Detect owner".'
        }
    })
    $form.Add_FormClosing({
        # $args[1] is the FormClosingEventArgs (sender is $args[0]/$this, unused here).
        # Guard #3: never close mid-export. An uncontrolled teardown could corrupt the PST, and the
        # owned-Outlook quit below would race the running export. Block the close and let it finish
        # (there is intentionally no cancel - aborting a COM CopyTo mid-flight is riskier than waiting).
        if ($script:MrExport.Active) {
            $args[1].Cancel = $true
            [System.Windows.Forms.MessageBox]::Show(
                'An export is in progress. Wait for it to finish before closing.',
                'Export in progress', 'OK', 'Warning') | Out-Null
            return
        }
        $timer.Stop()
        $hp = $script:MrProbe.Handle
        if ($hp) {
            try { $hp.Ps.Stop() } catch { }
            try { $hp.Runspace.Close() } catch { }
            try { $hp.Ps.Dispose() } catch { }
        }
        # Guard #7: quit the classic Outlook we auto-launched (KEEP-OPEN ownership ends on close).
        # Only an instance we started; an Outlook that was already running is left untouched.
        if ($script:MrOutlook.OwnedPid) {
            try { Stop-OwnedOutlook -OwnedPid $script:MrOutlook.OwnedPid | Out-Null } catch { }
            $script:MrOutlook.OwnedPid = $null
        }
    })

    # The message loop MUST run inside this function's scope. WinForms event handlers
    # ($timer.Start in Add_Shown, the button clicks, the timer tick) are scriptblocks
    # that resolve their variables against the scope stack live WHEN the event fires.
    # Running Application.Run in a caller after this function returned would tear down
    # this scope first, so locals ($timer, $list, the nested helpers) would be gone and
    # StrictMode would throw ("variable '$timer' ... has not been set"). With -Run we
    # keep the scope on the stack for the whole loop; without it we just return the
    # built form (construction-only test seam).
    if ($Run) {
        [void][System.Windows.Forms.Application]::Run($form)
        return
    }
    return $form
}

function Show-MailRescueGui {
    param([Parameter(Mandatory)][string]$ScriptPath)
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    New-MailRescueForm -ScriptPath $ScriptPath -Run
}

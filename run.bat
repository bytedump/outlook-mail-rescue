@echo off
REM Outlook Mail Rescue launcher.
REM Starts the GUI NON-elevated (so it can drive Outlook COM - an elevated host
REM cannot, UIPI blocks it). -ExecutionPolicy Bypass is per-process only; it does
REM not change the machine policy. -STA is required for WinForms + COM.
REM Elevation happens later, only for the disk scan (one UAC prompt at scan time).
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Invoke-MailRescue.ps1" %*

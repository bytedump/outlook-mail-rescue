#Requires -Version 5.1
<#
.SYNOPSIS
    Default configuration for Outlook Mail Rescue.
.DESCRIPTION
    Copy this file to config.ps1 and adjust. config.ps1 is gitignored. The app loads
    config.ps1 if present, otherwise falls back to these defaults. The file returns an
    ordered hashtable (its last expression), so just edit the values below.
#>

[ordered]@{
    # Where exported PSTs and the run report are written. The GUI lets the technician
    # override this per run. NOTE: a PST holds an entire mailbox and is UNENCRYPTED -
    # keep the output on an access-controlled location, not a shared drive.
    OutputFolder = Join-Path $env:USERPROFILE 'Desktop\MailRescue'

    # Output file name. Tokens: {username} {stamp} (stamp = yyyyMMdd-HHmmss).
    # {ticket} is also supported by the engine if you want a ticket number in the name.
    FileNameTemplate = '{username}_{stamp}.pst'

    # Disk-scan roots. Default is the whole system drive from its root.
    ScanRoots = @('C:\')

    # What to fold into the full export.
    IncludeArchive         = $true   # the user's Online Archive store
    IncludeSharedMailboxes = $false  # delegate/shared mailboxes (often huge)
    IncludePublicFolders   = $false  # rarely wanted, needs Exchange permissions

    # Wait at most this long for a fresh Cached Exchange Mode sync before exporting.
    SyncTimeoutMinutes = 30

    # v2 only: full path to an external/company-approved OST->PST converter executable
    # (or the bundled FOSS extractor). Empty = orphaned OST stays report-only.
    OstConverterPath = ''
}

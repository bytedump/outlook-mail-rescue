#Requires -Version 5.1
<#
.SYNOPSIS
    Switch the user from the new Outlook to classic by toggling UseNewOutlook.
.DESCRIPTION
    UseNewOutlook is a per-user (HKCU) DWORD preference. Setting it to 0 makes classic
    Outlook the active client on next launch. It is supported, reversible, and needs no
    admin. Every change is backed up first so it can be reverted at the end of a run.
    Effective ONLY if classic Outlook is also installed (the caller pre-checks that).

    The generic registry backup/restore helpers are unit-tested against a throwaway
    HKCU key, so the backup/revert semantics are verified for real.
#>

$script:UseNewOutlookPath = 'HKCU:\Software\Microsoft\office\16.0\outlook\preferences'
$script:UseNewOutlookName = 'UseNewOutlook'

# Capture a registry value's current state so it can be restored later. Returns
# { Path; Name; Existed; OldValue }.
function Backup-RegistryValue {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name)
    $existed = $false; $old = $null
    if (Test-Path $Path) {
        $v = Get-RegistryValue $Path $Name
        if ($null -ne $v) { $existed = $true; $old = $v }
    }
    return [pscustomobject]@{ Path = $Path; Name = $Name; Existed = $existed; OldValue = $old }
}

# Restore a value from its backup: put the old value back, or remove it if it was
# absent before we touched it.
function Restore-RegistryValue {
    param([Parameter(Mandatory)][pscustomobject]$Backup)
    if ($Backup.Existed) {
        if (-not (Test-Path $Backup.Path)) { New-Item -Path $Backup.Path -Force | Out-Null }
        Set-ItemProperty -Path $Backup.Path -Name $Backup.Name -Value ([int]$Backup.OldValue) -Type DWord
        Write-Log INFO "Restored $($Backup.Name)=$($Backup.OldValue)"
    } else {
        try {
            Remove-ItemProperty -Path $Backup.Path -Name $Backup.Name -ErrorAction Stop
            Write-Log INFO "Removed $($Backup.Name) (absent before this run)"
        } catch { }
    }
}

function Get-UseNewOutlook {
    return (Get-RegistryValue $script:UseNewOutlookPath $script:UseNewOutlookName)
}

# Set UseNewOutlook and return the backup for later revert. -Path overridable for tests.
function Set-UseNewOutlook {
    param(
        [Parameter(Mandatory)][int]$Value,
        [string]$Path = $script:UseNewOutlookPath,
        [string]$Name = $script:UseNewOutlookName
    )
    $backup = Backup-RegistryValue -Path $Path -Name $Name
    if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type DWord
    Write-Log INFO "Set $Name=$Value (was: existed=$($backup.Existed) old=$($backup.OldValue))"
    return $backup
}

# High-level: flip the user to classic Outlook. Returns the backup object (pass it to
# Restore-RegistryValue to revert). Caller must confirm classic is installed first.
function Switch-ToClassicOutlook {
    Write-Log INFO 'Switching active client to classic Outlook (UseNewOutlook=0).'
    return (Set-UseNewOutlook -Value 0)
}

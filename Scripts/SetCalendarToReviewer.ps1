# Requires Exchange Online PowerShell V2 module

param(
    [switch]$Rollback
)

function Get-UserMailboxes {
    Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
}

function Show-Permissions {
    $mailboxes = Get-UserMailboxes
    foreach ($mbx in $mailboxes) {
        $identity = "$($mbx.PrimarySmtpAddress):\Calendar"
        $perm = Get-MailboxFolderPermission -Identity $identity -User Default -ErrorAction SilentlyContinue
        $rights = if ($perm) { $perm.AccessRights } else { 'None' }
        Write-Host "$($mbx.PrimarySmtpAddress) - Default: $rights"
    }
}

function Set-DefaultPermission {
    param(
        [string]$AccessRight
    )
    $mailboxes = Get-UserMailboxes
    $backup = @()
    foreach ($mbx in $mailboxes) {
        $identity = "$($mbx.PrimarySmtpAddress):\Calendar"
        $perm = Get-MailboxFolderPermission -Identity $identity -User Default -ErrorAction SilentlyContinue
        $old = if ($perm) { $perm.AccessRights } else { 'None' }
        $backup += [PSCustomObject]@{Mailbox=$mbx.PrimarySmtpAddress;OldPermission=$old}
        if ($perm) {
            Set-MailboxFolderPermission -Identity $identity -User Default -AccessRights $AccessRight -ErrorAction Stop
        } else {
            Add-MailboxFolderPermission -Identity $identity -User Default -AccessRights $AccessRight -ErrorAction Stop
        }
    }
    $backup | Export-Csv -Path .\CalendarPermissionsBackup.csv -NoTypeInformation
    Test-Changes -Expected $AccessRight -Mailboxes $mailboxes
}

function Test-Changes {
    param(
        [string]$Expected,
        [array]$Mailboxes
    )
    foreach ($mbx in $Mailboxes) {
        $identity = "$($mbx.PrimarySmtpAddress):\Calendar"
        $perm = Get-MailboxFolderPermission -Identity $identity -User Default -ErrorAction SilentlyContinue
        if ($perm -and ($perm.AccessRights -contains $Expected)) {
            Write-Host "$($mbx.PrimarySmtpAddress) updated successfully" -ForegroundColor Green
        } else {
            Write-Host "$($mbx.PrimarySmtpAddress) failed to update" -ForegroundColor Red
        }
    }
}

function Restore-Permissions {
    if (Test-Path .\CalendarPermissionsBackup.csv) {
        $backup = Import-Csv .\CalendarPermissionsBackup.csv
        foreach ($row in $backup) {
            $identity = "$($row.Mailbox):\Calendar"
            if ($row.OldPermission -eq 'None') {
                Remove-MailboxFolderPermission -Identity $identity -User Default -Confirm:$false -ErrorAction SilentlyContinue
            } else {
                Set-MailboxFolderPermission -Identity $identity -User Default -AccessRights $row.OldPermission -ErrorAction SilentlyContinue
            }
            Write-Host "Restored $($row.Mailbox)" -ForegroundColor Yellow
        }
    } else {
        Write-Host 'No backup found. Cannot rollback.' -ForegroundColor Red
    }
}

if ($Rollback) {
    Restore-Permissions
    exit
}

Connect-ExchangeOnline -ShowBanner:$false

while ($true) {
    Write-Host '1. Show current default permissions'
    Write-Host '2. Set default permission for all users'
    Write-Host '3. Rollback to previous permissions'
    Write-Host '4. Exit'
    $choice = Read-Host 'Select an option'
    switch ($choice) {
        '1' { Show-Permissions }
        '2' {
            $perm = Read-Host 'Enter default permission (e.g., AvailabilityOnly, Reviewer)'
            if ($perm) { Set-DefaultPermission -AccessRight $perm }
        }
        '3' { Restore-Permissions }
        '4' { break }
        default { Write-Host 'Invalid option' }
    }
}

Disconnect-ExchangeOnline -Confirm:$false
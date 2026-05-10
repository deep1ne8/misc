#Requires -Modules ExchangeOnlineManagement
<#
.SYNOPSIS
    Interactive Exchange Online Calendar Permissions Manager
.DESCRIPTION
    Checks and sets 'Default' calendar permission to Reviewer for all user mailboxes
    - Exports backup of current permissions
    - Only changes where not already Reviewer
    - Verifies changes after applying
#>

function Connect-ExchangeOnlineSession {
    Write-Host "`n[+] Connecting to Exchange Online..." -ForegroundColor Cyan
    if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable)) {
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
    }
    Import-Module ExchangeOnlineManagement
    Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter your Exchange admin UPN")
}

function Export-CurrentPermissions {
    Write-Host "`n[+] Exporting current calendar permissions..." -ForegroundColor Yellow
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $exportPath = ".\CalendarPermissions_Backup_$timestamp.csv"

    $results = @()

    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox
    foreach ($mb in $mailboxes) {
        $perm = Get-EXOMailboxFolderPermission -Identity "$($mb.UserPrincipalName):\Calendar" -ErrorAction SilentlyContinue
        if ($perm) {
            foreach ($entry in $perm) {
                $results += [PSCustomObject]@{
                    User                = $mb.UserPrincipalName
                    Folder              = $entry.FolderName
                    UserOrGroup         = $entry.User
                    AccessRights        = $entry.AccessRights -join ','
                    SharingPermission   = $entry.SharingPermissionFlags
                }
            }
        }
    }

    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "[✓] Backup saved to $exportPath" -ForegroundColor Green
    return $results
}

function Set-PermissionsToReviewer {
    Write-Host "`n[+] Setting 'Default' calendar permission to Reviewer if not already set..." -ForegroundColor Yellow
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

    foreach ($mb in $mailboxes) {
        try {
            $defaultPerm = Get-EXOMailboxFolderPermission -Identity "$($mb.UserPrincipalName):\Calendar" -User Default -ErrorAction Stop
            if ($defaultPerm.AccessRights -ne "Reviewer") {
                Set-MailboxFolderPermission -Identity "$($mb.UserPrincipalName):\Calendar" -User Default -AccessRights Reviewer -Confirm:$false
                Write-Host "    → Updated: $($mb.UserPrincipalName)" -ForegroundColor Green
            } else {
                Write-Host "    → Skipped (already Reviewer): $($mb.UserPrincipalName)" -ForegroundColor Gray
            }
        }
        catch {
            Write-Warning "    × Failed or not set: $($mb.UserPrincipalName) - $_"
        }
    }
}

function Verify-Permissions {
    Write-Host "`n[+] Verifying updated calendar permissions..." -ForegroundColor Yellow
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox

    foreach ($mb in $mailboxes) {
        $defaultPerm = Get-EXOMailboxFolderPermission -Identity "$($mb.UserPrincipalName):\Calendar" -User Default -ErrorAction SilentlyContinue
        if ($defaultPerm.AccessRights -eq "Reviewer") {
            Write-Host "    ✓ Verified: $($mb.UserPrincipalName)" -ForegroundColor Green
        } else {
            Write-Warning "    × Issue with: $($mb.UserPrincipalName)"
        }
    }
}

function Show-Menu {
    Clear-Host
    Write-Host "=== Exchange Online Calendar Permission Manager ===" -ForegroundColor Cyan
    Write-Host "1. Connect to Exchange Online"
    Write-Host "2. Export current calendar permissions"
    Write-Host "3. Set Default permissions to Reviewer"
    Write-Host "4. Verify all permissions"
    Write-Host "5. Run All Steps"
    Write-Host "0. Exit"
}

do {
    Show-Menu
    $choice = Read-Host "Select an option"
    switch ($choice) {
        "1" { Connect-ExchangeOnlineSession }
        "2" { Export-CurrentPermissions | Out-Null }
        "3" { Set-PermissionsToReviewer }
        "4" { Verify-Permissions }
        "5" {
            Connect-ExchangeOnlineSession
            Export-CurrentPermissions | Out-Null
            Set-PermissionsToReviewer
            Verify-Permissions
        }
        "0" { Write-Host "Exiting..."; break }
        default { Write-Warning "Invalid option. Please select a valid number." }
    }
    Pause
} while ($true)

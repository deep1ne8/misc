# Calendar Permission Manager Script
# This script sets full calendar permissions for a specified user and verifies the configuration

function Test-ExchangeConnection {
    try {
        # Test if Exchange commands are available
        $null = Get-Command Get-MailboxFolderPermission -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ExchangeIfNeeded {
    if (-not (Test-ExchangeConnection)) {
        Write-Host "Exchange PowerShell module not detected. Attempting to connect..." -ForegroundColor Yellow
        
        try {
            # Try to connect to Exchange Online
            if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) {
                Import-Module ExchangeOnlineManagement
                Connect-ExchangeOnline -ErrorAction Stop
                Write-Host "Connected to Exchange Online." -ForegroundColor Green
            }
            else {
                throw "ExchangeOnlineManagement module not found. Please install it using: Install-Module -Name ExchangeOnlineManagement"
            }
        }
        catch {
            Write-Host "Error connecting to Exchange: $_" -ForegroundColor Red
            exit
        }
    }
}

function Set-CalendarPermissions {
    param (
        [Parameter(Mandatory = $true)]
        [string]$MailboxOwner,
        
        [Parameter(Mandatory = $true)]
        [string]$AccessUser
    )

    # Connect to Exchange if needed
    Connect-ExchangeIfNeeded

    # Get the calendar folder
    $calendarFolder = "$($MailboxOwner):\Calendar"
    
    try {
        # Set full access permission for the specified user
        Add-MailboxFolderPermission -Identity $calendarFolder -User $AccessUser -AccessRights Editor -SharingPermissionFlags Delegate, CanViewPrivateItems -ErrorAction Stop
        Write-Host "Successfully granted full calendar permissions to $AccessUser." -ForegroundColor Green
        
        # Verify the permissions
        $permissions = Get-MailboxFolderPermission -Identity $calendarFolder | Where-Object { $_.User.ToString() -eq $AccessUser }
        
        if ($permissions) {
            Write-Host "`nPermission verification:" -ForegroundColor Cyan
            Write-Host "User: $($permissions.User)" -ForegroundColor White
            Write-Host "Access Rights: $($permissions.AccessRights)" -ForegroundColor White
            Write-Host "Sharing Permission Flags: $($permissions.SharingPermissionFlags)" -ForegroundColor White
            
            if ($permissions.AccessRights -contains "Editor" -and 
                $permissions.SharingPermissionFlags -contains "Delegate" -and 
                $permissions.SharingPermissionFlags -contains "CanViewPrivateItems") {
                Write-Host "`nVerification PASSED: Full calendar permissions are correctly set." -ForegroundColor Green
            }
            else {
                Write-Host "`nVerification WARNING: Permissions were set but may not be complete." -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "`nVerification FAILED: Could not verify permissions. Please check manually." -ForegroundColor Red
        }
    }
    catch {
        # Handle the case where permissions already exist
        if ($_.Exception.Message -like "*already exists*") {
            Write-Host "User $AccessUser already has permissions on this calendar. Updating permissions..." -ForegroundColor Yellow
            
            try {
                # Update the permissions instead
                Set-MailboxFolderPermission -Identity $calendarFolder -User $AccessUser -AccessRights Editor -SharingPermissionFlags Delegate, CanViewPrivateItems
                Write-Host "Successfully updated calendar permissions for $AccessUser." -ForegroundColor Green
                
                # Verify the updated permissions
                $permissions = Get-MailboxFolderPermission -Identity $calendarFolder | Where-Object { $_.User.ToString() -eq $AccessUser }
                
                if ($permissions) {
                    Write-Host "`nPermission verification:" -ForegroundColor Cyan
                    Write-Host "User: $($permissions.User)" -ForegroundColor White
                    Write-Host "Access Rights: $($permissions.AccessRights)" -ForegroundColor White
                    Write-Host "Sharing Permission Flags: $($permissions.SharingPermissionFlags)" -ForegroundColor White
                    
                    if ($permissions.AccessRights -contains "Editor" -and 
                        $permissions.SharingPermissionFlags -contains "Delegate" -and 
                        $permissions.SharingPermissionFlags -contains "CanViewPrivateItems") {
                        Write-Host "`nVerification PASSED: Full calendar permissions are correctly set." -ForegroundColor Green
                    }
                    else {
                        Write-Host "`nVerification WARNING: Permissions were updated but may not be complete." -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Write-Host "Error updating permissions: $_" -ForegroundColor Red
            }
        }
        else {
            Write-Host "Error setting permissions: $_" -ForegroundColor Red
        }
    }
}

# Main script execution
Clear-Host
Write-Host "=== Calendar Permission Manager ===" -ForegroundColor Cyan

# Get mailbox information
$mailboxOwner = Read-Host "Enter the email address of the mailbox owner"
$accessUser = Read-Host "Enter the email address of the user who needs calendar access"

# Set and verify permissions
Set-CalendarPermissions -MailboxOwner $mailboxOwner -AccessUser $accessUser
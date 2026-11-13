# Requires ExchangeOnlineManagement module
# Make sure you're signed in with Connect-ExchangeOnline before running

Connect-ExchangeOnline

function Set-CalendarReviewerPermission {
    # Prompt for UPN (email address)
    $UserPrincipalName = Read-Host "Enter the user's email address (UPN) to test calendar permissions"

    # Validate the mailbox exists
    try {
        Get-EXOMailbox -Identity $UserPrincipalName -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Error: Mailbox '$UserPrincipalName' not found in Exchange Online." -ForegroundColor Red
        return
    }

    $CalendarIdentity = "${UserPrincipalName}:\Calendar"
    $Account = "Default"
    $TargetPermission = "Reviewer"

    # Get current permissions
    try {
        $Permissions = Get-MailboxFolderPermission -Identity $CalendarIdentity -ErrorAction Stop
    } catch {
        Write-Host "Error: Unable to retrieve calendar permissions for '$UserPrincipalName'." -ForegroundColor Red
        return
    }

    $Current = $Permissions | Where-Object { $_.User -eq $Account }

    if ($null -eq $Current) {
        Write-Host "No existing '$Account' permission found. Adding Reviewer rights..." -ForegroundColor Yellow
        Add-MailboxFolderPermission -Identity $CalendarIdentity -User $Account -AccessRights Reviewer | Out-Null
    } elseif ($Current.AccessRights -ne $TargetPermission) {
        Write-Host "'$Account' permission exists but is '$($Current.AccessRights)'. Updating to Reviewer..." -ForegroundColor Cyan
        Set-MailboxFolderPermission -Identity $CalendarIdentity -User $Account -AccessRights Reviewer | Out-Null
    } else {
        Write-Host "Permission already set to Reviewer for '$Account'." -ForegroundColor Green
    }

    Write-Host "Done." -ForegroundColor Green
}

# Run the function
Set-CalendarReviewerPermission

# Disconnect from Exchange Online
#Disconnect-ExchangeOnline

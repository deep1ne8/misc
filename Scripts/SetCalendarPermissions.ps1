#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Sets default calendar permissions for all user mailboxes in Exchange Online
.DESCRIPTION
    Connects to Exchange Online and applies specified calendar permissions to all user mailboxes,
    with options to exclude specific users and comprehensive logging
.PARAMETER AdminUPN
    Admin User Principal Name for Exchange Online connection
.PARAMETER Permission
    Calendar permission level to apply
.PARAMETER ExcludeList
    Comma-separated list of email addresses to exclude
.PARAMETER LogPath
    Custom path for log file (optional)
#>

[CmdletBinding()]
param(
    [string]$AdminUPN,
    [ValidateSet("AvailabilityOnly", "LimitedDetails", "Reviewer", "Editor")]
    [string]$Permission,
    [string]$ExcludeList,
    [string]$LogPath
)

# Error handling and logging functions
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:logFile -Value $logEntry
    
    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message -ForegroundColor Cyan }
    }
}

function Test-ExchangeConnection {
    try {
        Get-ConnectionInformation -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

# Initialize script
$ErrorActionPreference = "Stop"
$script:logFile = if ($LogPath) { $LogPath } else { "CalendarPermissions_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" }

try {
    # Initialize log
    Write-Log "=== Exchange Online Calendar Permissions Script Started ==="
    
    # Get admin UPN if not provided
    if (-not $AdminUPN) {
        $AdminUPN = Read-Host "Enter your admin UPN"
        if ([string]::IsNullOrWhiteSpace($AdminUPN)) {
            throw "Admin UPN is required"
        }
    }
    
    # Get permission level if not provided
    if (-not $Permission) {
        $accessOptions = @("AvailabilityOnly", "LimitedDetails", "Reviewer", "Editor")
        Write-Host "Available permission levels: $($accessOptions -join ', ')" -ForegroundColor Cyan
        $Permission = Read-Host "Enter default calendar permission to set"
        
        if ($accessOptions -notcontains $Permission) {
            throw "Invalid permission level: $Permission. Valid options: $($accessOptions -join ', ')"
        }
    }
    
    # Process exclusion list
    $excludedUsers = @()
    if (-not $ExcludeList) {
        $ExcludeList = Read-Host "Enter comma-separated email addresses to exclude (or press Enter to skip)"
    }
    
    if (-not [string]::IsNullOrWhiteSpace($ExcludeList)) {
        $excludedUsers = $ExcludeList.Split(",") | ForEach-Object { 
            $email = $_.Trim()
            if ($email -match '^[^@]+@[^@]+\.[^@]+$') {
                $email.ToLower()
            }
            else {
                Write-Log "Skipping invalid email format: $email" "WARNING"
            }
        } | Where-Object { $_ }
        
        Write-Log "Excluded users: $($excludedUsers -join ', ')"
    }
    
    # Connect to Exchange Online
    Write-Log "Connecting to Exchange Online as $AdminUPN..."
    
    if (-not (Test-ExchangeConnection)) {
        Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowProgress:$false
        
        if (-not (Test-ExchangeConnection)) {
            throw "Failed to establish Exchange Online connection"
        }
    }
    
    Write-Log "Successfully connected to Exchange Online" "SUCCESS"
    
    # Get all user mailboxes with progress tracking
    Write-Log "Retrieving user mailboxes..."
    $mailboxes = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox)
    $totalCount = $mailboxes.Count
    Write-Log "Found $totalCount user mailboxes to process"
    
    # Initialize counters
    $successCount = 0
    $errorCount = 0
    $excludedCount = 0
    $processed = 0
    
    # Process each mailbox
    foreach ($mailbox in $mailboxes) {
        $processed++
        $user = $mailbox.PrimarySmtpAddress.ToString().ToLower()
        $progressPercent = [math]::Round(($processed / $totalCount) * 100, 1)
        
        Write-Progress -Activity "Processing Calendar Permissions" -Status "Processing $user ($processed of $totalCount)" -PercentComplete $progressPercent
        
        if ($excludedUsers -contains $user) {
            Write-Log "Skipping excluded user: $user" "WARNING"
            $excludedCount++
            continue
        }
        
        try {
            # Check if calendar folder exists and get current permissions
            $calendarPath = "$user`:Calendar"
            $currentPerms = Get-MailboxFolderPermission -Identity $calendarPath -User Default -ErrorAction SilentlyContinue
            
            if ($currentPerms -and $currentPerms.AccessRights -eq $Permission) {
                Write-Log "User $user already has $Permission permission - skipping"
                $successCount++
                continue
            }
            
            # Set or update permission
            if ($currentPerms) {
                Set-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights $Permission -Confirm:$false
                Write-Log "Updated permission for $user to $Permission" "SUCCESS"
            }
            else {
                Add-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights $Permission -Confirm:$false
                Write-Log "Added permission for $user as $Permission" "SUCCESS"
            }
            
            $successCount++
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Failed to set permission for $user`: $errorMsg" "ERROR"
            $errorCount++
        }
    }
    
    Write-Progress -Activity "Processing Calendar Permissions" -Completed
    
    # Summary
    Write-Log "=== Operation Complete ==="
    Write-Log "Total mailboxes processed: $totalCount"
    Write-Log "Successful updates: $successCount" "SUCCESS"
    Write-Log "Errors encountered: $errorCount" $(if ($errorCount -gt 0) { "ERROR" } else { "INFO" })
    Write-Log "Excluded users: $excludedCount" "INFO"
    Write-Log "Log file saved to: $script:logFile" "SUCCESS"
    
    if ($errorCount -gt 0) {
        Write-Log "Review the log file for detailed error information" "WARNING"
    }
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
    exit 1
}
finally {
    Write-Log "=== Script Execution Ended ==="
}

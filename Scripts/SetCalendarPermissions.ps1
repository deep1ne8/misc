#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Exchange Online Calendar Permissions Management Tool
.DESCRIPTION
    Multi-mode tool for managing Exchange Online calendar permissions:
    - Mode 1: List and export all active mailboxes to CSV
    - Mode 2: Import CSV and apply calendar permissions with safety checks
.PARAMETER Mode
    Operation mode: List, Apply, or Interactive
.PARAMETER AdminUPN
    Admin User Principal Name for Exchange Online connection
.PARAMETER Permission
    Calendar permission level to apply (AvailabilityOnly, LimitedDetails, Reviewer, Editor)
.PARAMETER ImportPath
    Path to CSV file containing mailboxes to process
.PARAMETER ExportPath
    Path for exporting mailbox list CSV
.PARAMETER WhatIf
    Preview changes without applying them
#>

[CmdletBinding()]
param(
    [ValidateSet("List", "Apply", "Interactive")]
    [string]$Mode = "Interactive",
    [string]$AdminUPN,
    [string]$Permission,
    [string]$ImportPath,
    [string]$ExportPath,
    [switch]$WhatIf
)

# Global variables
$script:logFile = "ExchangeCalendarTool_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:validPermissions = @("AvailabilityOnly", "LimitedDetails", "Reviewer", "Editor")

# Enhanced logging and progress functions
function Write-Log {
    param(
        [string]$Message, 
        [ValidateSet("INFO", "SUCCESS", "WARNING", "ERROR", "VERBOSE")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $script:logFile -Value $logEntry -ErrorAction SilentlyContinue
    
    switch ($Level) {
        "ERROR" { Write-Host "❌ $Message" -ForegroundColor Red }
        "WARNING" { Write-Host "⚠️  $Message" -ForegroundColor Yellow }
        "SUCCESS" { Write-Host "✅ $Message" -ForegroundColor Green }
        "VERBOSE" { Write-Host "ℹ️  $Message" -ForegroundColor Cyan }
        default { Write-Host "📝 $Message" -ForegroundColor White }
    }
}

function Write-Progress-Enhanced {
    param(
        [string]$Activity,
        [string]$Status,
        [int]$Current,
        [int]$Total,
        [string]$CurrentItem = ""
    )
    
    $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 1) } else { 0 }
    $statusText = "$Status - $Current of $Total ($percent%)"
    
    if ($CurrentItem) {
        $statusText += " - $CurrentItem"
    }
    
    Write-Progress -Activity $Activity -Status $statusText -PercentComplete $percent
    Write-Log "Progress: $statusText" "VERBOSE"
}

function Test-ExchangeConnection {
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Connect-ToExchangeOnline {
    param([string]$UPN)
    
    Write-Log "Checking Exchange Online connection..." "VERBOSE"
    
    if (Test-ExchangeConnection) {
        Write-Log "Already connected to Exchange Online" "SUCCESS"
        return $true
    }
    
    if (-not $UPN) {
        $UPN = Read-Host "Enter your admin UPN for Exchange Online"
        if ([string]::IsNullOrWhiteSpace($UPN)) {
            Write-Log "Admin UPN is required for connection" "ERROR"
            return $false
        }
    }
    
    Write-Log "Connecting to Exchange Online as $UPN..." "INFO"
    
    try {
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress:$false -ErrorAction Stop
        
        if (Test-ExchangeConnection) {
            Write-Log "Successfully connected to Exchange Online" "SUCCESS"
            
            # Display connection info
            $connectionInfo = Get-ConnectionInformation
            Write-Log "Connected to: $($connectionInfo.Name) ($($connectionInfo.UserPrincipalName))" "VERBOSE"
            return $true
        }
        else {
            Write-Log "Connection verification failed" "ERROR"
            return $false
        }
    }
    catch {
        Write-Log "Failed to connect to Exchange Online: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-AllActiveMailboxes {
    param([string]$ExportPath)
    
    Write-Log "=== MAILBOX DISCOVERY MODE ===" "INFO"
    Write-Log "Retrieving all active user mailboxes from Exchange Online..." "INFO"
    
    try {
        # Get mailboxes with detailed properties
        Write-Log "Querying Exchange Online for user mailboxes..." "VERBOSE"
        $mailboxes = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox | 
            Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, WhenCreated, 
                         @{Name='MailboxSize';Expression={(Get-MailboxStatistics $_.Identity -ErrorAction SilentlyContinue).TotalItemSize}},
                         @{Name='LastLogonTime';Expression={(Get-MailboxStatistics $_.Identity -ErrorAction SilentlyContinue).LastLogonTime}})
        
        $totalCount = $mailboxes.Count
        Write-Log "Found $totalCount active user mailboxes" "SUCCESS"
        
        if ($totalCount -eq 0) {
            Write-Log "No mailboxes found. Please verify your permissions and connection." "WARNING"
            return @()
        }
        
        # Display sample mailboxes
        Write-Log "Sample mailboxes found:" "INFO"
        $mailboxes | Select-Object -First 5 | ForEach-Object {
            Write-Log "  - $($_.DisplayName) ($($_.PrimarySmtpAddress))" "VERBOSE"
        }
        
        if ($totalCount -gt 5) {
            Write-Log "  ... and $($totalCount - 5) more mailboxes" "VERBOSE"
        }
        
        # Export to CSV if requested
        if ($ExportPath -or ($Mode -eq "Interactive")) {
            if (-not $ExportPath) {
                $defaultPath = "ExchangeMailboxes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                $ExportPath = Read-Host "Enter CSV export path (or press Enter for '$defaultPath')"
                if ([string]::IsNullOrWhiteSpace($ExportPath)) {
                    $ExportPath = $defaultPath
                }
            }
            
            try {
                $mailboxes | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Log "Mailbox list exported to: $ExportPath" "SUCCESS"
                Write-Log "CSV contains $totalCount mailbox records with columns: DisplayName, PrimarySmtpAddress, UserPrincipalName, WhenCreated, MailboxSize, LastLogonTime" "INFO"
            }
            catch {
                Write-Log "Failed to export CSV: $($_.Exception.Message)" "ERROR"
            }
        }
        
        return $mailboxes
    }
    catch {
        Write-Log "Failed to retrieve mailboxes: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Import-MailboxesFromCSV {
    param([string]$Path)
    
    if (-not $Path) {
        $Path = Read-Host "Enter path to CSV file containing mailboxes to process"
    }
    
    if (-not (Test-Path $Path)) {
        Write-Log "CSV file not found: $Path" "ERROR"
        return @()
    }
    
    Write-Log "Importing mailboxes from CSV: $Path" "INFO"
    
    try {
        $csvData = Import-Csv -Path $Path
        $importedCount = $csvData.Count
        
        Write-Log "CSV file contains $importedCount records" "VERBOSE"
        
        # Validate required columns
        $requiredColumns = @("PrimarySmtpAddress")
        $csvColumns = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $csvColumns }
        if ($missingColumns) {
            Write-Log "CSV missing required columns: $($missingColumns -join ', ')" "ERROR"
            Write-Log "Available columns: $($csvColumns -join ', ')" "INFO"
            return @()
        }
        
        # Extract and validate email addresses
        $emailAddresses = @()
        foreach ($row in $csvData) {
            $email = $row.PrimarySmtpAddress
            if ([string]::IsNullOrWhiteSpace($email)) {
                Write-Log "Skipping row with empty email address" "WARNING"
                continue
            }
            
            if ($email -match '^[^@]+@[^@]+\.[^@]+$') {
                $emailAddresses += $email.ToLower().Trim()
            }
            else {
                Write-Log "Skipping invalid email format: $email" "WARNING"
            }
        }
        
        $validCount = $emailAddresses.Count
        Write-Log "Imported $validCount valid email addresses from CSV" "SUCCESS"
        
        if ($validCount -gt 0) {
            Write-Log "Sample imported addresses:" "INFO"
            $emailAddresses | Select-Object -First 3 | ForEach-Object {
                Write-Log "  - $_" "VERBOSE"
            }
            if ($validCount -gt 3) {
                Write-Log "  ... and $($validCount - 3) more addresses" "VERBOSE"
            }
        }
        
        return $emailAddresses
    }
    catch {
        Write-Log "Failed to import CSV: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Set-CalendarPermissions {
    param(
        [string[]]$EmailAddresses,
        [string]$Permission,
        [switch]$WhatIf
    )
    
    if (-not $Permission) {
        Write-Log "Available permission levels: $($script:validPermissions -join ', ')" "INFO"
        do {
            $Permission = Read-Host "Enter calendar permission level to apply"
        } while ($Permission -notin $script:validPermissions)
    }
    
    if ($Permission -notin $script:validPermissions) {
        Write-Log "Invalid permission level: $Permission. Valid options: $($script:validPermissions -join ', ')" "ERROR"
        return
    }
    
    $totalCount = $EmailAddresses.Count
    Write-Log "=== APPLYING CALENDAR PERMISSIONS ===" "INFO"
    Write-Log "Permission level: $Permission" "INFO"
    Write-Log "Target mailboxes: $totalCount" "INFO"
    Write-Log "What-If mode: $($WhatIf.IsPresent)" "INFO"
    
    if ($WhatIf) {
        Write-Log "PREVIEW MODE - No changes will be applied" "WARNING"
    }
    else {
        Write-Log "⚠️  LIVE MODE - Changes will be applied to mailboxes" "WARNING"
        $confirm = Read-Host "Do you want to continue? (y/N)"
        if ($confirm -ne 'y' -and $confirm -ne 'Y') {
            Write-Log "Operation cancelled by user" "INFO"
            return
        }
    }
    
    # Initialize counters
    $successCount = 0
    $errorCount = 0
    $skippedCount = 0
    $processed = 0
    
    foreach ($email in $EmailAddresses) {
        $processed++
        Write-Progress-Enhanced -Activity "Setting Calendar Permissions" -Status "Processing mailboxes" -Current $processed -Total $totalCount -CurrentItem $email
        
        try {
            $calendarPath = "$($email.ToLower()):Calendar"
            
            # Check current permissions
            Write-Log "Checking current permissions for $email..." "VERBOSE"
            $currentPerms = Get-MailboxFolderPermission -Identity $calendarPath -User Default -ErrorAction SilentlyContinue
            
            if ($currentPerms -and $currentPerms.AccessRights -contains $Permission) {
                Write-Log "Mailbox $email already has $Permission permission - skipping" "INFO"
                $skippedCount++
                continue
            }
            
            if ($WhatIf) {
                if ($currentPerms) {
                    Write-Log "[PREVIEW] Would update permission for $email from $($currentPerms.AccessRights) to $Permission" "INFO"
                }
                else {
                    Write-Log "[PREVIEW] Would add permission for $email as $Permission" "INFO"
                }
                $successCount++
            }
            else {
                # Apply the permission change
                if ($currentPerms) {
                    Set-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights $Permission -Confirm:$false
                    Write-Log "Updated permission for $email to $Permission" "SUCCESS"
                }
                else {
                    Add-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights $Permission -Confirm:$false
                    Write-Log "Added permission for $email as $Permission" "SUCCESS"
                }
                $successCount++
            }
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Log "Failed to process $email`: $errorMsg" "ERROR"
            $errorCount++
        }
        
        # Brief pause to avoid throttling
        if ($processed % 10 -eq 0) {
            Start-Sleep -Milliseconds 100
        }
    }
    
    Write-Progress -Activity "Setting Calendar Permissions" -Completed
    
    # Final summary
    Write-Log "=== OPERATION SUMMARY ===" "INFO"
    Write-Log "Total mailboxes processed: $totalCount" "INFO"
    Write-Log "Successful operations: $successCount" "SUCCESS"
    Write-Log "Skipped (already configured): $skippedCount" "INFO"
    Write-Log "Errors encountered: $errorCount" $(if ($errorCount -gt 0) { "ERROR" } else { "INFO" })
    
    if ($WhatIf) {
        Write-Log "This was a preview run - no actual changes were made" "WARNING"
    }
    
    if ($errorCount -gt 0) {
        Write-Log "Review the log file for detailed error information: $script:logFile" "WARNING"
    }
}

function Show-InteractiveMenu {
    Write-Host ""
    Write-Host "=== Exchange Online Calendar Permissions Tool ===" -ForegroundColor Cyan
    Write-Host "1. List and export all active mailboxes to CSV" -ForegroundColor White
    Write-Host "2. Import CSV and apply calendar permissions" -ForegroundColor White  
    Write-Host "3. Preview changes (What-If mode)" -ForegroundColor Yellow
    Write-Host "4. Exit" -ForegroundColor Gray
    Write-Host ""
}

# Main execution logic
try {
    Write-Log "=== Exchange Online Calendar Permissions Tool Started ===" "INFO"
    Write-Log "Mode: $Mode | WhatIf: $($WhatIf.IsPresent)" "VERBOSE"
    Write-Log "Log file: $script:logFile" "INFO"
    
    # Connect to Exchange Online
    if (-not (Connect-ToExchangeOnline -UPN $AdminUPN)) {
        throw "Failed to establish Exchange Online connection"
    }
    
    switch ($Mode) {
        "List" {
            $null = Get-AllActiveMailboxes -ExportPath $ExportPath
        }
        
        "Apply" {
            if (-not $ImportPath) {
                throw "ImportPath parameter is required for Apply mode"
            }
            
            $emailAddresses = Import-MailboxesFromCSV -Path $ImportPath
            if ($emailAddresses.Count -eq 0) {
                throw "No valid email addresses found in CSV"
            }
            
            Set-CalendarPermissions -EmailAddresses $emailAddresses -Permission $Permission -WhatIf:$WhatIf
        }
        
        "Interactive" {
            do {
                Show-InteractiveMenu
                $choice = Read-Host "Select an option (1-4)"
                
                switch ($choice) {
                    "1" {
                        Write-Host ""
                        $null = Get-AllActiveMailboxes
                    }
                    "2" {
                        Write-Host ""
                        $emailAddresses = Import-MailboxesFromCSV
                        if ($emailAddresses.Count -gt 0) {
                            Set-CalendarPermissions -EmailAddresses $emailAddresses
                        }
                    }
                    "3" {
                        Write-Host ""
                        $emailAddresses = Import-MailboxesFromCSV
                        if ($emailAddresses.Count -gt 0) {
                            Set-CalendarPermissions -EmailAddresses $emailAddresses -WhatIf
                        }
                    }
                    "4" {
                        Write-Log "Exiting application" "INFO"
                        break
                    }
                    default {
                        Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Red
                    }
                }
                
                if ($choice -ne "4" -and $choice -in @("1","2","3")) {
                    Write-Host ""
                    Read-Host "Press Enter to continue"
                }
                
            } while ($choice -ne "4")
        }
    }
    
    Write-Log "Script execution completed successfully" "SUCCESS"
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
    Write-Host ""
    Write-Host "For troubleshooting help:" -ForegroundColor Yellow
    Write-Host "1. Verify you have Exchange Online PowerShell permissions" -ForegroundColor Gray
    Write-Host "2. Check your network connection and proxy settings" -ForegroundColor Gray
    Write-Host "3. Review the log file: $script:logFile" -ForegroundColor Gray
    exit 1
}
finally {
    Write-Log "=== Script Execution Ended ===" "INFO"
}

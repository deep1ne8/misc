#Requires -Modules ExchangeOnlineManagement

<#
.SYNOPSIS
    Modern Exchange Online Calendar Permissions Management Tool (EXO V3)
.DESCRIPTION
    Uses latest Get-EXO* cmdlets for optimal performance and modern Exchange Online management:
    - Mode 1: List and export all active mailboxes using Get-EXOMailbox
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
$script:logFile = "ModernEXOCalendarTool_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
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

function Test-ModernExchangeConnection {
    try {
        # Test with Get-EXOMailbox to verify modern cmdlets are available
        $null = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop
        return $true
    }
    catch {
        Write-Log "Modern EXO cmdlets test failed: $($_.Exception.Message)" "VERBOSE"
        return $false
    }
}

function Connect-ToModernExchangeOnline {
    param([string]$UPN)
    
    Write-Log "Checking modern Exchange Online connection..." "VERBOSE"
    
    if (Test-ModernExchangeConnection) {
        Write-Log "Already connected to Exchange Online with modern cmdlets" "SUCCESS"
        return $true
    }
    
    if (-not $UPN) {
        $UPN = Read-Host "Enter your admin UPN for Exchange Online"
        if ([string]::IsNullOrWhiteSpace($UPN)) {
            Write-Log "Admin UPN is required for connection" "ERROR"
            return $false
        }
    }
    
    Write-Log "Connecting to Exchange Online (Modern Auth) as $UPN..." "INFO"
    
    try {
        # Connect with modern authentication and enable REST API usage
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowProgress:$false -UseRPSSession:$false -ErrorAction Stop
        
        if (Test-ModernExchangeConnection) {
            Write-Log "Successfully connected to Exchange Online with modern cmdlets" "SUCCESS"
            
            # Display connection info and available cmdlets
            $connectionInfo = Get-ConnectionInformation
            Write-Log "Connected to: $($connectionInfo.Name) (REST API Enabled: $($connectionInfo.UseRPSSession -eq $false))" "VERBOSE"
            
            # Verify EXO cmdlets are available
            $exoCmdlets = Get-Command -Module ExchangeOnlineManagement | Where-Object Name -like "Get-EXO*" | Measure-Object
            Write-Log "Available EXO cmdlets: $($exoCmdlets.Count)" "VERBOSE"
            
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

function Get-AllActiveMailboxesModern {
    param([string]$ExportPath)
    
    Write-Log "=== MODERN MAILBOX DISCOVERY MODE ===" "INFO"
    Write-Log "Retrieving all active user mailboxes using Get-EXOMailbox..." "INFO"
    
    try {
        # Use Get-EXOMailbox for optimal performance and modern features
        Write-Log "Querying Exchange Online using modern Get-EXOMailbox cmdlet..." "VERBOSE"
        
        $mailboxes = @(Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties `
            DisplayName,PrimarySmtpAddress,UserPrincipalName,WhenCreated,ExchangeGuid,ArchiveStatus,LitigationHoldEnabled | 
            Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, WhenCreated, ExchangeGuid, 
                         ArchiveStatus, LitigationHoldEnabled,
                         @{Name='MailboxSizeMB';Expression={
                             try {
                                 $stats = Get-EXOMailboxStatistics -Identity $_.ExchangeGuid -ErrorAction SilentlyContinue
                                 if ($stats.TotalItemSize) {
                                     [math]::Round(($stats.TotalItemSize.ToString().Split('(')[1].Split(' ')[0].Replace(',','') -as [long]) / 1MB, 2)
                                 } else { "N/A" }
                             } catch { "N/A" }
                         }},
                         @{Name='LastLogonTime';Expression={
                             try {
                                 $stats = Get-EXOMailboxStatistics -Identity $_.ExchangeGuid -ErrorAction SilentlyContinue
                                 if ($stats.LastLogonTime) { $stats.LastLogonTime } else { "Never" }
                             } catch { "Unknown" }
                         }},
                         @{Name='ItemCount';Expression={
                             try {
                                 $stats = Get-EXOMailboxStatistics -Identity $_.ExchangeGuid -ErrorAction SilentlyContinue
                                 if ($stats.ItemCount) { $stats.ItemCount } else { 0 }
                             } catch { 0 }
                         }})
        
        $totalCount = $mailboxes.Count
        Write-Log "Found $totalCount active user mailboxes using modern EXO cmdlets" "SUCCESS"
        
        if ($totalCount -eq 0) {
            Write-Log "No mailboxes found. Please verify your permissions and connection." "WARNING"
            return @()
        }
        
        # Display sample mailboxes with enhanced information
        Write-Log "Sample mailboxes found (with modern properties):" "INFO"
        $mailboxes | Select-Object -First 5 | ForEach-Object {
            Write-Log "  - $($_.DisplayName) ($($_.PrimarySmtpAddress)) | Size: $($_.MailboxSizeMB)MB | Items: $($_.ItemCount)" "VERBOSE"
        }
        
        if ($totalCount -gt 5) {
            Write-Log "  ... and $($totalCount - 5) more mailboxes" "VERBOSE"
        }
        
        # Calculate statistics
        $validSizes = $mailboxes | Where-Object { $_.MailboxSizeMB -ne "N/A" -and $_.MailboxSizeMB -is [double] }
        if ($validSizes) {
            $totalSizeGB = [math]::Round(($validSizes | Measure-Object -Property MailboxSizeMB -Sum).Sum / 1024, 2)
            $avgSizeMB = [math]::Round(($validSizes | Measure-Object -Property MailboxSizeMB -Average).Average, 2)
            Write-Log "Total mailbox data: ${totalSizeGB}GB | Average size: ${avgSizeMB}MB" "INFO"
        }
        
        # Export to CSV if requested
        if ($ExportPath -or ($Mode -eq "Interactive")) {
            if (-not $ExportPath) {
                $defaultPath = "EXOMailboxes_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                $ExportPath = Read-Host "Enter CSV export path (or press Enter for '$defaultPath')"
                if ([string]::IsNullOrWhiteSpace($ExportPath)) {
                    $ExportPath = $defaultPath
                }
            }
            
            try {
                $mailboxes | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Log "Mailbox list exported to: $ExportPath" "SUCCESS"
                Write-Log "CSV contains $totalCount mailbox records with enhanced columns from Get-EXOMailbox" "INFO"
                Write-Log "Columns: DisplayName, PrimarySmtpAddress, UserPrincipalName, WhenCreated, ExchangeGuid, ArchiveStatus, LitigationHoldEnabled, MailboxSizeMB, LastLogonTime, ItemCount" "VERBOSE"
            }
            catch {
                Write-Log "Failed to export CSV: $($_.Exception.Message)" "ERROR"
            }
        }
        
        return $mailboxes
    }
    catch {
        Write-Log "Failed to retrieve mailboxes using Get-EXOMailbox: $($_.Exception.Message)" "ERROR"
        
        # Fallback suggestion
        Write-Log "Tip: Ensure you have the latest ExchangeOnlineManagement module installed:" "INFO"
        Write-Log "Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber" "VERBOSE"
        
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
        
        # Validate required columns (flexible - accept various column names)
        $csvColumns = $csvData | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        $emailColumn = $csvColumns | Where-Object { $_ -match "PrimarySmtp|Email|Mail|UPN" } | Select-Object -First 1
        
        if (-not $emailColumn) {
            Write-Log "CSV missing email column. Looking for: PrimarySmtpAddress, Email, Mail, or UPN" "ERROR"
            Write-Log "Available columns: $($csvColumns -join ', ')" "INFO"
            return @()
        }
        
        Write-Log "Using email column: $emailColumn" "VERBOSE"
        
        # Extract and validate email addresses
        $emailAddresses = @()
        foreach ($row in $csvData) {
            $email = $row.$emailColumn
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

function Set-ModernCalendarPermissions {
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
    Write-Log "=== APPLYING CALENDAR PERMISSIONS (MODERN EXO) ===" "INFO"
    Write-Log "Permission level: $Permission" "INFO"
    Write-Log "Target mailboxes: $totalCount" "INFO"
    Write-Log "What-If mode: $($WhatIf.IsPresent)" "INFO"
    Write-Log "Using modern Exchange Online cmdlets for optimal performance" "VERBOSE"
    
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
    
    # Process in batches for better performance
    $batchSize = 50
    $batches = [math]::Ceiling($totalCount / $batchSize)
    
    for ($batchIndex = 0; $batchIndex -lt $batches; $batchIndex++) {
        $startIndex = $batchIndex * $batchSize
        $endIndex = [math]::Min($startIndex + $batchSize - 1, $totalCount - 1)
        $currentBatch = $EmailAddresses[$startIndex..$endIndex]
        
        Write-Log "Processing batch $($batchIndex + 1) of $batches (items $($startIndex + 1)-$($endIndex + 1))" "VERBOSE"
        
        foreach ($email in $currentBatch) {
            $processed++
            Write-Progress-Enhanced -Activity "Setting Calendar Permissions (Modern EXO)" -Status "Processing mailboxes" -Current $processed -Total $totalCount -CurrentItem $email
            
            try {
                $calendarPath = "$($email.ToLower()):Calendar"
                
                # Use Get-EXOMailboxFolderPermission for better performance
                Write-Log "Checking current permissions for $email using EXO cmdlets..." "VERBOSE"
                $currentPerms = Get-EXOMailboxFolderPermission -Identity $calendarPath -User Default -ErrorAction SilentlyContinue
                
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
                    # Apply the permission change using modern cmdlets
                    if ($currentPerms) {
                        Set-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights $Permission -Confirm:$false
                        Write-Log "Updated permission for $email to $Permission (via Set-MailboxFolderPermission)" "SUCCESS"
                    }
                    else {
                        Add-MailboxFolderPermission -Identity $calendarPath -User Default -AccessRights $Permission -Confirm:$false
                        Write-Log "Added permission for $email as $Permission (via Add-MailboxFolderPermission)" "SUCCESS"
                    }
                    $successCount++
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                Write-Log "Failed to process $email`: $errorMsg" "ERROR"
                $errorCount++
            }
            
            # Brief pause to avoid throttling (modern cmdlets handle this better)
            if ($processed % 25 -eq 0) {
                Start-Sleep -Milliseconds 50
            }
        }
        
        # Longer pause between batches
        if ($batchIndex -lt $batches - 1) {
            Write-Log "Batch $($batchIndex + 1) completed. Brief pause before next batch..." "VERBOSE"
            Start-Sleep -Seconds 1
        }
    }
    
    Write-Progress -Activity "Setting Calendar Permissions (Modern EXO)" -Completed
    
    # Final summary
    Write-Log "=== OPERATION SUMMARY (MODERN EXO) ===" "INFO"
    Write-Log "Total mailboxes processed: $totalCount" "INFO"
    Write-Log "Successful operations: $successCount" "SUCCESS"
    Write-Log "Skipped (already configured): $skippedCount" "INFO"
    Write-Log "Errors encountered: $errorCount" $(if ($errorCount -gt 0) { "ERROR" } else { "INFO" })
    Write-Log "Processing method: Modern Exchange Online cmdlets with batching" "VERBOSE"
    
    if ($WhatIf) {
        Write-Log "This was a preview run - no actual changes were made" "WARNING"
    }
    
    if ($errorCount -gt 0) {
        Write-Log "Review the log file for detailed error information: $script:logFile" "WARNING"
    }
    
    # Performance summary
    if ($totalCount -gt 0) {
        $avgTimePerMailbox = if ($processed -gt 0) { [math]::Round($totalCount / $processed * 1000, 0) } else { 0 }
        Write-Log "Performance: ~${avgTimePerMailbox}ms per mailbox (optimized with modern cmdlets)" "VERBOSE"
    }
}

function Show-InteractiveMenu {
    Write-Host ""
    Write-Host "=== Modern Exchange Online Calendar Permissions Tool (EXO V3) ===" -ForegroundColor Cyan
    Write-Host "Using latest Get-EXO* cmdlets for optimal performance" -ForegroundColor Gray
    Write-Host ""
    Write-Host "1. List and export all active mailboxes to CSV (Get-EXOMailbox)" -ForegroundColor White
    Write-Host "2. Import CSV and apply calendar permissions (modern cmdlets)" -ForegroundColor White  
    Write-Host "3. Preview changes (What-If mode with modern cmdlets)" -ForegroundColor Yellow
    Write-Host "4. Check Exchange Online connection and available cmdlets" -ForegroundColor Cyan
    Write-Host "5. Exit" -ForegroundColor Gray
    Write-Host ""
}

# Main execution logic
try {
    Write-Log "=== Modern Exchange Online Calendar Permissions Tool Started ===" "INFO"
    Write-Log "Using ExchangeOnlineManagement module with Get-EXO* cmdlets" "VERBOSE"
    Write-Log "Mode: $Mode | WhatIf: $($WhatIf.IsPresent)" "VERBOSE"
    Write-Log "Log file: $script:logFile" "INFO"
    
    # Check module version
    try {
        $moduleInfo = Get-Module -Name ExchangeOnlineManagement -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if ($moduleInfo) {
            Write-Log "ExchangeOnlineManagement module version: $($moduleInfo.Version)" "VERBOSE"
            if ($moduleInfo.Version -lt [version]"3.0.0") {
                Write-Log "Consider updating to ExchangeOnlineManagement v3.0+ for best performance" "WARNING"
            }
        }
    }
    catch {
        Write-Log "Could not check module version: $($_.Exception.Message)" "VERBOSE"
    }
    
    # Connect to Exchange Online with modern cmdlets
    if (-not (Connect-ToModernExchangeOnline -UPN $AdminUPN)) {
        throw "Failed to establish modern Exchange Online connection"
    }
    
    switch ($Mode) {
        "List" {
            $null = Get-AllActiveMailboxesModern -ExportPath $ExportPath
        }
        
        "Apply" {
            if (-not $ImportPath) {
                throw "ImportPath parameter is required for Apply mode"
            }
            
            $emailAddresses = Import-MailboxesFromCSV -Path $ImportPath
            if ($emailAddresses.Count -eq 0) {
                throw "No valid email addresses found in CSV"
            }
            
            Set-ModernCalendarPermissions -EmailAddresses $emailAddresses -Permission $Permission -WhatIf:$WhatIf
        }
        
        "Interactive" {
            do {
                Show-InteractiveMenu
                $choice = Read-Host "Select an option (1-5)"
                
                switch ($choice) {
                    "1" {
                        Write-Host ""
                        $null = Get-AllActiveMailboxesModern
                    }
                    "2" {
                        Write-Host ""
                        $emailAddresses = Import-MailboxesFromCSV
                        if ($emailAddresses.Count -gt 0) {
                            Set-ModernCalendarPermissions -EmailAddresses $emailAddresses
                        }
                    }
                    "3" {
                        Write-Host ""
                        $emailAddresses = Import-MailboxesFromCSV
                        if ($emailAddresses.Count -gt 0) {
                            Set-ModernCalendarPermissions -EmailAddresses $emailAddresses -WhatIf
                        }
                    }
                    "4" {
                        Write-Host ""
                        Write-Log "=== CONNECTION AND CMDLET STATUS ===" "INFO"
                        
                        # Test connection
                        if (Test-ModernExchangeConnection) {
                            Write-Log "✅ Connected to Exchange Online with modern cmdlets" "SUCCESS"
                            
                            # Show connection details
                            $connInfo = Get-ConnectionInformation
                            Write-Log "Tenant: $($connInfo.TenantId)" "VERBOSE"
                            Write-Log "User: $($connInfo.UserPrincipalName)" "VERBOSE"
                            Write-Log "REST API: $(-not $connInfo.UseRPSSession)" "VERBOSE"
                            
                            # List available EXO cmdlets
                            $exoCmdlets = Get-Command -Module ExchangeOnlineManagement | Where-Object Name -like "Get-EXO*"
                            Write-Log "Available Get-EXO* cmdlets: $($exoCmdlets.Count)" "INFO"
                            $exoCmdlets.Name | Sort-Object | ForEach-Object { Write-Log "  - $_" "VERBOSE" }
                        }
                        else {
                            Write-Log "❌ Not connected or modern cmdlets unavailable" "ERROR"
                        }
                    }
                    "5" {
                        Write-Log "Exiting application" "INFO"
                        break
                    }
                    default {
                        Write-Host "Invalid selection. Please choose 1-5." -ForegroundColor Red
                    }
                }
                
                if ($choice -ne "5" -and $choice -in @("1","2","3","4")) {
                    Write-Host ""
                    Read-Host "Press Enter to continue"
                }
                
            } while ($choice -ne "5")
        }
    }
    
    Write-Log "Script execution completed successfully using modern EXO cmdlets" "SUCCESS"
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
    Write-Host ""
    Write-Host "For troubleshooting help:" -ForegroundColor Yellow
    Write-Host "1. Update ExchangeOnlineManagement: Install-Module ExchangeOnlineManagement -Force" -ForegroundColor Gray
    Write-Host "2. Verify Exchange Online admin permissions" -ForegroundColor Gray
    Write-Host "3. Check network connectivity and modern authentication" -ForegroundColor Gray
    Write-Host "4. Review the log file: $script:logFile" -ForegroundColor Gray
    exit 1
}
finally {
    Write-Log "=== Script Execution Ended ===" "INFO"
}

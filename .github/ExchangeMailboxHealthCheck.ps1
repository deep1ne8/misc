# Exchange Settings Health Check Script
# Version 1.0
# This script performs a comprehensive health check on Exchange mailbox settings
# focusing on calendar sharing, permissions, and Teams integration issues

function Start-ExchangeHealthCheck {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Mailboxes,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportCSV,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath = ".\ExchangeHealthCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss')",
        
        [Parameter(Mandatory = $false)]
        [switch]$FixIssues
    )
    
    begin {
        # Create output directory if it doesn't exist
        if (!(Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }
        
        # Initialize log file
        $logFile = Join-Path -Path $OutputPath -ChildPath "HealthCheck.log"
        $null = New-Item -Path $logFile -ItemType File -Force
        
        function Write-Log {
            param (
                [string]$Message,
                [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
                [string]$Level = "INFO"
            )

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logEntry = "[$timestamp] [$Level] $Message"
            Add-Content -Path $logFile -Value $logEntry

            # Also output to console with color
            switch ($Level) {
                "INFO" { Write-Host $logEntry -ForegroundColor Cyan }
                "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
                "ERROR" { Write-Host $logEntry -ForegroundColor Red }
                "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
            }
        }

        
        Write-Log "Starting Exchange Health Check for mailboxes: $($Mailboxes -join ', ')" "INFO"
        
        # Verify Exchange connection
        try {
            $exchangeSession = Get-PSSession | Where-Object { $_.ConfigurationName -eq "Microsoft.Exchange" -and $_.State -eq "Opened" }
            if (!$exchangeSession) {
                throw "No active Exchange Online PowerShell session found."
            }
            Write-Log "Active Exchange Online PowerShell session detected" "SUCCESS"
        }
        catch {
            Write-Log "Error verifying Exchange Online connection: $_" "ERROR"
            Write-Log "Please connect to Exchange Online using Connect-ExchangeOnline before running this script" "ERROR"
            return
        }
        
        # Initialize results collection
        $healthResults = @()
    }
    
    process {
        foreach ($mailbox in $Mailboxes) {
            Write-Log "Processing mailbox: $mailbox" "INFO"
            
            # Create result object for this mailbox
            $mailboxHealth = [PSCustomObject]@{
                Mailbox = $mailbox
                ExistsInExchange = $false
                MailboxType = ""
                RecipientTypeDetails = ""
                HiddenFromAddressLists = $null
                CalendarFolderExists = $false
                CalendarFolderPath = ""
                CalendarProcessingEnabled = $false
                DefaultCalendarPermission = ""
                UserCalendarPermissions = @()
                CASMailboxEnabled = $false
                ActiveSyncEnabled = $null
                OWAEnabled = $null
                PopEnabled = $null
                ImapEnabled = $null
                EwsEnabled = $null
                TeamsMeetingAddinDisabled = $null
                ThrottlingPolicy = ""
                ELCMailboxFlags = ""
                RetentionHoldEnabled = $null
                LitigationHoldEnabled = $null
                ArchiveStatus = ""
                ForwardingAddress = ""
                ForwardingSmtpAddress = ""
                DeliverToMailboxAndForward = $null
                IssuesDetected = @()
                FixesApplied = @()
            }
            
            # 1. Check if mailbox exists
            try {
                $mbx = Get-Mailbox -Identity $mailbox -ErrorAction Stop
                $mailboxHealth.ExistsInExchange = $true
                $mailboxHealth.MailboxType = $mbx.RecipientTypeDetails
                $mailboxHealth.RecipientTypeDetails = $mbx.RecipientTypeDetails
                $mailboxHealth.HiddenFromAddressLists = $mbx.HiddenFromAddressListsEnabled
                Write-Log "Mailbox exists: $mailbox ($($mbx.RecipientTypeDetails))" "SUCCESS"
                
                # Check forwarding settings
                $mailboxHealth.ForwardingAddress = $mbx.ForwardingAddress
                $mailboxHealth.ForwardingSmtpAddress = $mbx.ForwardingSmtpAddress
                $mailboxHealth.DeliverToMailboxAndForward = $mbx.DeliverToMailboxAndForward
                
                if ($mbx.ForwardingAddress -or $mbx.ForwardingSmtpAddress) {
                    $forwardDest = if ($mbx.ForwardingAddress) { $mbx.ForwardingAddress } else { $mbx.ForwardingSmtpAddress }
                    $deliverCopy = if ($mbx.DeliverToMailboxAndForward) { "keeping a copy" } else { "not keeping a copy" }
                    $mailboxHealth.IssuesDetected += "Mailbox is forwarding to $forwardDest ($deliverCopy)"
                    Write-Log "Mailbox forwarding detected for $mailbox to $forwardDest ($deliverCopy)" "WARNING"
                }
                
                # Check hold status
                $mailboxHealth.RetentionHoldEnabled = $mbx.RetentionHoldEnabled
                $mailboxHealth.LitigationHoldEnabled = $mbx.LitigationHoldEnabled
                $mailboxHealth.ArchiveStatus = if ($mbx.ArchiveStatus) { $mbx.ArchiveStatus } else { "Not Enabled" }
                
                if ($mbx.RetentionHoldEnabled -or $mbx.LitigationHoldEnabled) {
                    $holdType = @()
                    if ($mbx.RetentionHoldEnabled) { $holdType += "Retention Hold" }
                    if ($mbx.LitigationHoldEnabled) { $holdType += "Litigation Hold" }
                    $mailboxHealth.IssuesDetected += "Mailbox is on hold: $($holdType -join ', ')"
                    Write-Log "Mailbox $mailbox is on hold: $($holdType -join ', ')" "WARNING"
                }
                
            }
            catch {
                Write-Log "Mailbox not found: $mailbox - $_" "ERROR"
                $mailboxHealth.IssuesDetected += "Mailbox does not exist or cannot be accessed"
                $healthResults += $mailboxHealth
                continue
            }
            
            # 2. Check calendar folder
            try {
                $calendarStats = Get-MailboxFolderStatistics $mailbox -FolderScope Calendar -ErrorAction Stop
                if ($calendarStats) {
                    $defaultCalendar = $calendarStats | Where-Object { $_.FolderType -eq "Calendar" } | Select-Object -First 1
                    if ($defaultCalendar) {
                        $mailboxHealth.CalendarFolderExists = $true
                        $mailboxHealth.CalendarFolderPath = $defaultCalendar.FolderPath
                        Write-Log "Calendar folder exists for $mailbox: $($defaultCalendar.FolderPath)" "SUCCESS"
                    }
                    else {
                        $mailboxHealth.IssuesDetected += "Default calendar folder not found"
                        Write-Log "Default calendar folder not found for $mailbox" "ERROR"
                    }
                }
            }
            catch {
                Write-Log "Error retrieving calendar folder for $mailbox: $_" "ERROR"
                $mailboxHealth.IssuesDetected += "Cannot access calendar folder"
            }
            
            # 3. Check calendar processing settings
            if ($mailboxHealth.CalendarFolderExists) {
                try {
                    $calProc = Get-CalendarProcessing -Identity $mailbox -ErrorAction Stop
                    $mailboxHealth.CalendarProcessingEnabled = $true
                    
                    # Check for potential calendar processing issues
                    if ($calProc.RemovePrivateProperty -eq $true) {
                        $mailboxHealth.IssuesDetected += "Calendar processing is removing private flag from meetings"
                        Write-Log "Calendar processing is removing private flag from meetings for $mailbox" "WARNING"
                        
                        if ($FixIssues) {
                            try {
                                Set-CalendarProcessing -Identity $mailbox -RemovePrivateProperty $false
                                $mailboxHealth.FixesApplied += "Set RemovePrivateProperty to False"
                                Write-Log "Fixed: Set RemovePrivateProperty to False for $mailbox" "SUCCESS"
                            }
                            catch {
                                Write-Log "Failed to update RemovePrivateProperty for $mailbox: $_" "ERROR"
                            }
                        }
                    }
                }
                catch {
                    Write-Log "Error retrieving calendar processing settings for $mailbox: $_" "ERROR"
                    $mailboxHealth.IssuesDetected += "Calendar processing settings inaccessible"
                }
                
                # 4. Check calendar permissions
                try {
                    $calendarPermissions = Get-MailboxFolderPermission -Identity "$($mailbox):\Calendar" -ErrorAction Stop
                    
                    # Get Default permission
                    $defaultPerm = $calendarPermissions | Where-Object { $_.User.DisplayName -eq "Default" }
                    if ($defaultPerm) {
                        $mailboxHealth.DefaultCalendarPermission = $defaultPerm.AccessRights -join ', '
                        Write-Log "Default calendar permission for $mailbox is: $($mailboxHealth.DefaultCalendarPermission)" "INFO"
                        
                        # Check if default permissions are too restrictive or too permissive
                        if ($defaultPerm.AccessRights -contains "None") {
                            $mailboxHealth.IssuesDetected += "Default calendar permission is None (too restrictive for organization)"
                            Write-Log "Default calendar permission is None (too restrictive) for $mailbox" "WARNING"
                            
                            if ($FixIssues) {
                                try {
                                    Set-MailboxFolderPermission -Identity "$($mailbox):\Calendar" -User Default -AccessRights AvailabilityOnly
                                    $mailboxHealth.FixesApplied += "Set Default calendar permission to AvailabilityOnly"
                                    $mailboxHealth.DefaultCalendarPermission = "AvailabilityOnly"
                                    Write-Log "Fixed: Set Default calendar permission to AvailabilityOnly for $mailbox" "SUCCESS"
                                }
                                catch {
                                    Write-Log "Failed to update Default calendar permission for $mailbox: $_" "ERROR"
                                }
                            }
                        }
                        elseif ($defaultPerm.AccessRights -contains "Owner" -or $defaultPerm.AccessRights -contains "PublishingEditor") {
                            $mailboxHealth.IssuesDetected += "Default calendar permission is too permissive ($($defaultPerm.AccessRights -join ', '))"
                            Write-Log "Default calendar permission is too permissive for $mailbox: $($defaultPerm.AccessRights -join ', ')" "WARNING"
                            
                            if ($FixIssues) {
                                try {
                                    Set-MailboxFolderPermission -Identity "$($mailbox):\Calendar" -User Default -AccessRights LimitedDetails
                                    $mailboxHealth.FixesApplied += "Set Default calendar permission to LimitedDetails (from too permissive)"
                                    $mailboxHealth.DefaultCalendarPermission = "LimitedDetails"
                                    Write-Log "Fixed: Set Default calendar permission to LimitedDetails for $mailbox" "SUCCESS"
                                }
                                catch {
                                    Write-Log "Failed to update Default calendar permission for $mailbox: $_" "ERROR"
                                }
                            }
                        }
                    }
                    else {
                        $mailboxHealth.IssuesDetected += "No Default calendar permission found"
                        Write-Log "No Default calendar permission found for $mailbox" "WARNING"
                    }
                    
                    # Check user permissions
                    $userPermissions = $calendarPermissions | Where-Object { 
                        $_.User.DisplayName -ne "Default" -and 
                        $_.User.DisplayName -ne "Anonymous" -and
                        -not ($_.User.DisplayName -like "Exchange*")
                    }
                    
                    if ($userPermissions) {
                        foreach ($perm in $userPermissions) {
                            $permObj = [PSCustomObject]@{
                                User = $perm.User.DisplayName
                                AccessRights = $perm.AccessRights -join ', '
                                SharingFlags = $perm.SharingPermissionFlags -join ', '
                            }
                            $mailboxHealth.UserCalendarPermissions += $permObj
                            
                            # Check for Team integration issues (Reviewer permission is often insufficient)
                            if ($perm.AccessRights -contains "Reviewer" -and -not ($perm.AccessRights -contains "Editor" -or $perm.AccessRights -contains "Owner")) {
                                $mailboxHealth.IssuesDetected += "User $($perm.User.DisplayName) has Reviewer permission which may be insufficient for Teams calendar integration"
                                Write-Log "User $($perm.User.DisplayName) has Reviewer permission on $mailbox calendar - may cause Teams integration issues" "WARNING"
                                
                                if ($FixIssues) {
                                    try {
                                        $userEmail = $perm.User.ADRecipient.PrimarySmtpAddress
                                        if (-not $userEmail) {
                                            $userEmail = $perm.User.DisplayName
                                        }
                                        Set-MailboxFolderPermission -Identity "$($mailbox):\Calendar" -User $userEmail -AccessRights Editor
                                        $mailboxHealth.FixesApplied += "Upgraded $($perm.User.DisplayName) permission from Reviewer to Editor"
                                        Write-Log "Fixed: Upgraded $($perm.User.DisplayName) permission from Reviewer to Editor on $mailbox calendar" "SUCCESS"
                                    }
                                    catch {
                                        Write-Log "Failed to upgrade permission for $($perm.User.DisplayName) on $mailbox calendar: $_" "ERROR"
                                    }
                                }
                            }
                        }
                    }
                    else {
                        Write-Log "No user-specific calendar permissions found for $mailbox" "INFO"
                    }
                }
                catch {
                    Write-Log "Error retrieving calendar permissions for $mailbox: $_" "ERROR"
                    $mailboxHealth.IssuesDetected += "Cannot access calendar permissions"
                }
            }
            
            # 5. Check CAS mailbox settings
            try {
                $casMailbox = Get-CASMailbox -Identity $mailbox -ErrorAction Stop
                $mailboxHealth.CASMailboxEnabled = $true
                $mailboxHealth.ActiveSyncEnabled = $casMailbox.ActiveSyncEnabled
                $mailboxHealth.OWAEnabled = $casMailbox.OWAEnabled
                $mailboxHealth.PopEnabled = $casMailbox.PopEnabled
                $mailboxHealth.ImapEnabled = $casMailbox.ImapEnabled
                $mailboxHealth.EwsEnabled = $casMailbox.EwsEnabled
                $mailboxHealth.TeamsMeetingAddinDisabled = $casMailbox.TeamsMeetingAddInDisabled
                
                Write-Log "CAS mailbox settings retrieved for $mailbox" "SUCCESS"
                
                # Check for Teams integration issues
                if ($casMailbox.TeamsMeetingAddInDisabled -eq $true) {
                    $mailboxHealth.IssuesDetected += "Teams Meeting Add-in is disabled"
                    Write-Log "Teams Meeting Add-in is disabled for $mailbox" "WARNING"
                    
                    if ($FixIssues) {
                        try {
                            Set-CASMailbox -Identity $mailbox -TeamsMeetingAddInDisabled $false
                            $mailboxHealth.FixesApplied += "Enabled Teams Meeting Add-in"
                            $mailboxHealth.TeamsMeetingAddinDisabled = $false
                            Write-Log "Fixed: Enabled Teams Meeting Add-in for $mailbox" "SUCCESS"
                        }
                        catch {
                            Write-Log "Failed to enable Teams Meeting Add-in for $mailbox: $_" "ERROR"
                        }
                    }
                }
                
                # Check EWS settings (needed for Teams calendar integration)
                if ($casMailbox.EwsEnabled -eq $false) {
                    $mailboxHealth.IssuesDetected += "Exchange Web Services (EWS) is disabled - required for Teams integration"
                    Write-Log "Exchange Web Services (EWS) is disabled for $mailbox - required for Teams integration" "WARNING"
                    
                    if ($FixIssues) {
                        try {
                            Set-CASMailbox -Identity $mailbox -EwsEnabled $true
                            $mailboxHealth.FixesApplied += "Enabled Exchange Web Services (EWS)"
                            $mailboxHealth.EwsEnabled = $true
                            Write-Log "Fixed: Enabled Exchange Web Services (EWS) for $mailbox" "SUCCESS"
                        }
                        catch {
                            Write-Log "Failed to enable Exchange Web Services (EWS) for $mailbox: $_" "ERROR"
                        }
                    }
                }
                
                # Check EWS application policies
                if ($casMailbox.EwsApplicationAccessPolicy -eq "BlockList") {
                    $mailboxHealth.IssuesDetected += "EWS Application Access Policy is set to BlockList - may block Teams"
                    Write-Log "EWS Application Access Policy is set to BlockList for $mailbox - may block Teams" "WARNING"
                }
                elseif ($casMailbox.EwsApplicationAccessPolicy -eq "AllowList") {
                    $mailboxHealth.IssuesDetected += "EWS Application Access Policy is set to AllowList - ensure Teams is allowed"
                    Write-Log "EWS Application Access Policy is set to AllowList for $mailbox - ensure Teams is allowed" "WARNING"
                }
            }
            catch {
                Write-Log "Error retrieving CAS mailbox settings for $mailbox: $_" "ERROR"
                $mailboxHealth.IssuesDetected += "CAS mailbox settings inaccessible"
            }
            
            # 6. Check throttling policy
            try {
                $throttlingPolicy = Get-ThrottlingPolicy -Identity $mbx.ThrottlingPolicy -ErrorAction SilentlyContinue
                if ($throttlingPolicy) {
                    $mailboxHealth.ThrottlingPolicy = $mbx.ThrottlingPolicy
                    
                    # Check for restrictive EWS policies
                    if ($throttlingPolicy.EwsMaxConcurrency -lt 10 -or 
                        $throttlingPolicy.EwsMaxSubscriptions -lt 20 -or 
                        $throttlingPolicy.EwsMaxBurst -lt 300) {
                        $mailboxHealth.IssuesDetected += "Throttling policy may be too restrictive for Teams integration"
                        Write-Log "Throttling policy $($mbx.ThrottlingPolicy) may be too restrictive for Teams integration" "WARNING"
                    }
                }
                else {
                    $mailboxHealth.ThrottlingPolicy = "Default"
                }
            }
            catch {
                Write-Log "Error retrieving throttling policy for $mailbox: $_" "WARNING"
                $mailboxHealth.ThrottlingPolicy = "Unknown"
            }
            
            # 7. Check mailbox limits and quotas
            try {
                if ($mbx.ProhibitSendQuota -and $mbx.ProhibitSendReceiveQuota) {
                    $sendQuota = [math]::Round([double]($mbx.ProhibitSendQuota.ToString().Split("(")[1].Split(" ")[0]) / 1GB, 2)
                    $sendReceiveQuota = [math]::Round([double]($mbx.ProhibitSendReceiveQuota.ToString().Split("(")[1].Split(" ")[0]) / 1GB, 2)
                    
                    $currentSize = 0
                    $stats = Get-MailboxStatistics -Identity $mailbox -ErrorAction SilentlyContinue
                    if ($stats) {
                        $currentSize = [math]::Round([double]($stats.TotalItemSize.Value.ToString().Split("(")[1].Split(" ")[0]) / 1GB, 2)
                        
                        # Check if approaching quota limits
                        $percentUsed = ($currentSize / $sendQuota) * 100
                        if ($percentUsed -gt 85) {
                            $mailboxHealth.IssuesDetected += "Mailbox is at $($percentUsed.ToString("0.0"))% of quota ($currentSize GB / $sendQuota GB)"
                            Write-Log "Mailbox $mailbox is at $($percentUsed.ToString("0.0"))% of quota ($currentSize GB / $sendQuota GB)" "WARNING"
                        }
                    }
                }
            }
            catch {
                Write-Log "Error checking mailbox quotas for $mailbox: $_" "WARNING"
            }
            
            # Summarize findings
            if ($mailboxHealth.IssuesDetected.Count -eq 0) {
                Write-Log "No issues detected for $mailbox" "SUCCESS"
            }
            else {
                Write-Log "Found $($mailboxHealth.IssuesDetected.Count) potential issues for $mailbox" "WARNING"
                foreach ($issue in $mailboxHealth.IssuesDetected) {
                    Write-Log "  - $issue" "WARNING"
                }
            }
            
            if ($mailboxHealth.FixesApplied.Count -gt 0) {
                Write-Log "Applied $($mailboxHealth.FixesApplied.Count) fixes for $mailbox" "SUCCESS"
                foreach ($fix in $mailboxHealth.FixesApplied) {
                    Write-Log "  - $fix" "SUCCESS"
                }
            }
            
            # Add to results collection
            $healthResults += $mailboxHealth
        }
    }
    
    end {
        # Export results to CSV if requested
        if ($ExportCSV) {
            $csvPath = Join-Path -Path $OutputPath -ChildPath "HealthCheckResults.csv"
            
            # Prepare data for CSV export
            $csvResults = $healthResults | ForEach-Object {
                $result = $_
                $userPerms = if ($result.UserCalendarPermissions) { 
                    ($result.UserCalendarPermissions | ForEach-Object { "$($_.User):$($_.AccessRights)" }) -join '; ' 
                } else { 
                    "None" 
                }
                
                $issues = if ($result.IssuesDetected) { $result.IssuesDetected -join '; ' } else { "None" }
                $fixes = if ($result.FixesApplied) { $result.FixesApplied -join '; ' } else { "None" }
                
                [PSCustomObject]@{
                    Mailbox = $result.Mailbox
                    ExistsInExchange = $result.ExistsInExchange
                    MailboxType = $result.MailboxType
                    HiddenFromAddressLists = $result.HiddenFromAddressLists
                    CalendarFolderExists = $result.CalendarFolderExists
                    DefaultCalendarPermission = $result.DefaultCalendarPermission
                    UserCalendarPermissions = $userPerms
                    TeamsMeetingAddinDisabled = $result.TeamsMeetingAddinDisabled
                    EwsEnabled = $result.EwsEnabled
                    ForwardingEnabled = if ($result.ForwardingAddress -or $result.ForwardingSmtpAddress) { $true } else { $false }
                    IssuesDetected = $issues
                    FixesApplied = $fixes
                }
            }
            
            $csvResults | Export-Csv -Path $csvPath -NoTypeInformation
            Write-Log "Results exported to CSV: $csvPath" "SUCCESS"
        }
        
        # Generate HTML report
        $htmlPath = Join-Path -Path $OutputPath -ChildPath "HealthCheckReport.html"
        
        $htmlHeader = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Exchange Health Check Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2, h3 { color: #0078D4; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th, td { padding: 8px; text-align: left; border: 1px solid #ddd; }
        th { background-color: #0078D4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .warning { color: orange; }
        .error { color: red; }
        .success { color: green; }
        .summary { background-color: #f0f0f0; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .details { margin-top: 20px; }
        .issue-list { margin-left: 20px; }
    </style>
</head>
<body>
    <h1>Exchange Health Check Report</h1>
    <p>Generated on: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
"@
        
        $htmlSummary = @"
    <div class="summary">
        <h2>Summary</h2>
        <p>Total mailboxes checked: $($healthResults.Count)</p>
        <p>Mailboxes with issues: $($healthResults | Where-Object { $_.IssuesDetected.Count -gt 0 } | Measure-Object).Count</p>
        <p>Fixes applied: $($healthResults | Where-Object { $_.FixesApplied.Count -gt 0 } | Measure-Object).Count</p>
    </div>
"@
        
        $htmlMailboxes = ""
        foreach ($result in $healthResults) {
            $issuesHtml = ""
            if ($result.IssuesDetected.Count -gt 0) {
                $issuesHtml = "<h4>Issues Detected:</h4><ul class='issue-list'>"
                foreach ($issue in $result.IssuesDetected) {
                    $issuesHtml += "<li class='warning'>$issue</li>"
                }
                $issuesHtml += "</ul>"
            }
            
            $fixesHtml = ""
            if ($result.FixesApplied.Count -gt 0) {
                $fixesHtml = "<h4>Fixes Applied:</h4><ul class='issue-list'>"
                foreach ($fix in $result.FixesApplied) {
                    $fixesHtml += "<li class='success'>$fix</li>"
                }
                $fixesHtml += "</ul>"
            }
            
            $userPermsHtml = "<h4>Calendar Permissions:</h4><ul>"
            if ($result.UserCalendarPermissions.Count -gt 0) {
                foreach ($perm in $result.UserCalendarPermissions) {
                    $userPermsHtml += "<li>$($perm.User): $($perm.AccessRights)</li>"
                }
            }
            else {
                $userPermsHtml += "<li>No user-specific permissions found</li>"
            }
            $userPermsHtml += "</ul>"
            
            $statusClass = if ($result.IssuesDetected.Count -gt 0) { "warning" } else { "success" }
            
            $htmlMailboxes += @"
    <div class="details">
        <h3>Mailbox: $($result.Mailbox)</h3>
        <p class="$statusClass">Status: $(if ($result.IssuesDetected.Count -gt 0) { "Issues Found" } else { "Healthy" })</p>
        
        <h4>Basic Information:</h4>
        <table>
            <tr><th>Property</th><th>Value</th></tr>
            <tr><td>Mailbox Type</td><td>$($result.MailboxType)</td></tr>
            <tr><td>Hidden From Address Lists</td><td>$($result.HiddenFromAddressLists)</td></tr>
            <tr><td>Calendar Folder Exists</td><td>$($result.CalendarFolderExists)</td></tr>
            <tr><td>Default Calendar Permission</td><td>$($result.DefaultCalendarPermission)</td></tr>
            <tr><td>Teams Meeting Add-in Disabled</td><td>$($result.TeamsMeetingAddinDisabled)</td></tr>
            <tr><td>EWS Enabled</td><td>$($result.EwsEnabled)</td></tr>
            <tr><td>Forwarding</td><td>$(if ($result.ForwardingAddress -or $result.ForwardingSmtpAddress) { "Enabled" } else { "Disabled" })</td></tr>
        </table>
        
        $userPermsHtml
        $issuesHtml
        $fixesHtml
    </div>
"@
        }
        
        $htmlFooter = @"
</body>
</html>
"@
        
        $htmlReport = $htmlHeader + $htmlSummary + $htmlMailboxes + $htmlFooter
        $htmlReport | Out-File -FilePath $htmlPath -Encoding utf8
        
        Write-Log "HTML report generated: $htmlPath" "SUCCESS"
        Write-Log "Exchange Health Check completed. Results can be found in: $OutputPath" "SUCCESS"
        
        # Return results object
        return $healthResults
    }
}

# Example Usage:
# Start-ExchangeHealthCheck -Mailboxes "seth@distributedsun.com", "alexa.moore@distributedsun.com" -ExportCSV
# Start-ExchangeHealthCheck -Mailboxes "seth@distributedsun.com", "alexa.moore@distributedsun.com" -FixIssues
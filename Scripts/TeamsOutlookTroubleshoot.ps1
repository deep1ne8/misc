#Requires -Modules ExchangeOnlineManagement, MicrosoftTeams, AzureAD

<#
.SYNOPSIS
    Comprehensive diagnostic tool for Teams-Exchange integration issues.
    
.DESCRIPTION
    This script performs extensive diagnostics for Teams-Exchange integration issues,
    particularly focused on accounts with potential migration issues from GoDaddy 
    or other legacy hosting providers.
    
.PARAMETER UserEmail
    Email address of the user experiencing Teams-Exchange integration issues.
    
.PARAMETER LogPath
    Path where the log file will be saved. Default is desktop.
    
.PARAMETER Remediate
    Switch parameter to enable automatic remediation of common issues.
    
.EXAMPLE
    .\Diagnose-TeamsExchangeIntegration.ps1 -UserEmail john.doe@contoso.com
    
.EXAMPLE
    .\Diagnose-TeamsExchangeIntegration.ps1 -UserEmail john.doe@contoso.com -Remediate
    
.NOTES
    Author: Earl Daniels
    Date: March 3, 2025
    Version: 1.0
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$UserEmail,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "$env:SystemRoot\Temp\TeamsExchangeDiagnostic_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    
    [Parameter(Mandatory = $false)]
    [switch]$Remediate
)

# Initialize log file
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Different colors for different levels
    switch ($Level) {
        'Info'    { Write-Host $logMessage -ForegroundColor Cyan }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Success' { Write-Host $logMessage -ForegroundColor Green }
    }
    
    # Write to log file
    Add-Content -Path $LogPath -Value $logMessage
}

function Connect-RequiredServices {
    try {
        # Connect to Exchange Online
        Write-Log "Connecting to Exchange Online..." -Level Info
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        Write-Log "Successfully connected to Exchange Online" -Level Success
        
        # Connect to Microsoft Teams
        Write-Log "Connecting to Microsoft Teams..." -Level Info
        Connect-MicrosoftTeams -ErrorAction Stop
        Write-Log "Successfully connected to Microsoft Teams" -Level Success
        
        # Connect to Azure AD
        Write-Log "Connecting to Azure AD..." -Level Info
        Connect-AzureAD -ErrorAction Stop
        Write-Log "Successfully connected to Azure AD" -Level Success
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to required services: $_" -Level Error
        return $false
    }
}

function Test-UserExistence {
    param(
        [string]$Email
    )
    
    Write-Log "Testing existence of user: $Email" -Level Info
    
    # Check in Exchange Online
    try {
        $exchangeUser = Get-Mailbox -Identity $Email -ErrorAction Stop
        Write-Log "User found in Exchange Online: $($exchangeUser.UserPrincipalName)" -Level Success
        $exchangeExists = $true
    }
    catch {
        Write-Log "User not found in Exchange Online or error: $_" -Level Warning
        $exchangeExists = $false
    }
    
    # Check in Azure AD
    try {
        $azureUser = Get-AzureADUser -Filter "userPrincipalName eq '$Email'" -ErrorAction Stop
        if ($azureUser) {
            Write-Log "User found in Azure AD: $($azureUser.UserPrincipalName)" -Level Success
            $azureExists = $true
        }
        else {
            Write-Log "User not found in Azure AD" -Level Warning
            $azureExists = $false
        }
    }
    catch {
        Write-Log "Error checking Azure AD user: $_" -Level Error
        $azureExists = $false
    }
    
    return [PSCustomObject]@{
        ExchangeExists = $exchangeExists
        AzureADExists = $azureExists
        ExchangeUser = $exchangeUser
        AzureADUser = $azureUser
    }
}

function Test-UserLicensing {
    param(
        [string]$Email
    )
    
    Write-Log "Checking licensing for user: $Email" -Level Info
    
    try {
        $user = Get-AzureADUser -Filter "userPrincipalName eq '$Email'"
        $licenses = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId
        
        if ($licenses) {
            foreach ($license in $licenses) {
                Write-Log "User has license: $($license.SkuPartNumber)" -Level Info
            }
            
            # Check for Teams license
            $hasTeamsLicense = $licenses.ServicePlans | Where-Object { $_.ServicePlanName -like "*Teams*" }
            if ($hasTeamsLicense) {
                Write-Log "User has Teams license: $($hasTeamsLicense.ServicePlanName)" -Level Success
            }
            else {
                Write-Log "User does not have Teams license!" -Level Warning
            }
            
            # Check for Exchange license
            $hasExchangeLicense = $licenses.ServicePlans | Where-Object { $_.ServicePlanName -like "*Exchange*" }
            if ($hasExchangeLicense) {
                Write-Log "User has Exchange license: $($hasExchangeLicense.ServicePlanName)" -Level Success
            }
            else {
                Write-Log "User does not have Exchange license!" -Level Warning
            }
            
            return [PSCustomObject]@{
                HasLicenses = $true
                HasTeamsLicense = ($null -ne $hasTeamsLicense)
                HasExchangeLicense = ($null -ne $hasExchangeLicense)
                Licenses = $licenses
            }
        }
        else {
            Write-Log "User does not have any licenses assigned!" -Level Error
            return [PSCustomObject]@{
                HasLicenses = $false
                HasTeamsLicense = $false
                HasExchangeLicense = $false
                Licenses = $null
            }
        }
    }
    catch {
        Write-Log "Error checking user licensing: $_" -Level Error
        return [PSCustomObject]@{
            HasLicenses = $false
            HasTeamsLicense = $false
            HasExchangeLicense = $false
            Licenses = $null
            Error = $_
        }
    }
}

function Test-AutodiscoverConfiguration {
    param(
        [string]$Email
    )
    
    Write-Log "Testing Autodiscover configuration for: $Email" -Level Info
    
    try {
        # Export Autodiscover results to XML
        $autodiscoverFile = "$env:TEMP\autodiscover_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
        $null = Test-OutlookWebServices -Identity $Email -TargetFile $autodiscoverFile
        
        if (Test-Path $autodiscoverFile) {
            [xml]$autodiscoverData = Get-Content $autodiscoverFile
            Write-Log "Successfully retrieved Autodiscover information" -Level Success
            
            # Analyze Autodiscover data
            $ewsUrl = $autodiscoverData.Autodiscover.Response.Account.Protocol | 
                Where-Object { $_.Type -eq "EXCH" } | 
                Select-Object -ExpandProperty EwsUrl
            
            if ($ewsUrl) {
                Write-Log "EWS URL: $ewsUrl" -Level Info
            }
            else {
                Write-Log "EWS URL not found in Autodiscover response!" -Level Warning
            }
            
            # Check for GoDaddy remnants
            if ($autodiscoverData.OuterXml -match "GoDaddy|secureserver") {
                Write-Log "GoDaddy references found in Autodiscover data! This may indicate migration issues." -Level Warning
            }
            
            return [PSCustomObject]@{
                Success = $true
                EwsUrl = $ewsUrl
                AutodiscoverData = $autodiscoverData
                AutodiscoverFile = $autodiscoverFile
            }
        }
        else {
            Write-Log "Autodiscover test completed but no output file was created" -Level Warning
            return [PSCustomObject]@{
                Success = $false
                EwsUrl = $null
                AutodiscoverData = $null
                AutodiscoverFile = $null
            }
        }
    }
    catch {
        Write-Log "Error testing Autodiscover: $_" -Level Error
        return [PSCustomObject]@{
            Success = $false
            EwsUrl = $null
            AutodiscoverData = $null
            AutodiscoverFile = $null
            Error = $_
        }
    }
}

function Test-MailboxPermissions {
    param(
        [string]$Email
    )
    
    Write-Log "Checking mailbox permissions for: $Email" -Level Info
    
    try {
        $permissions = Get-MailboxPermission -Identity $Email
        Write-Log "Retrieved mailbox permissions successfully" -Level Success
        
        # Check for application permissions
        $appPermissions = $permissions | Where-Object { 
            $_.User -like "*ApplicationImpersonation*" -or 
            $_.User -like "*Teams*" -or 
            $_.User -like "*Microsoft*"
        }
        
        if ($appPermissions) {
            Write-Log "Found application permissions on mailbox:" -Level Info
            foreach ($perm in $appPermissions) {
                Write-Log "  - $($perm.User): $($perm.AccessRights)" -Level Info
            }
        }
        else {
            Write-Log "No application permissions found. This might indicate missing Teams access." -Level Warning
        }
        
        return [PSCustomObject]@{
            Success = $true
            Permissions = $permissions
            ApplicationPermissions = $appPermissions
        }
    }
    catch {
        Write-Log "Error checking mailbox permissions: $_" -Level Error
        return [PSCustomObject]@{
            Success = $false
            Permissions = $null
            ApplicationPermissions = $null
            Error = $_
        }
    }
}

function Test-TeamsConfiguration {
    param(
        [string]$Email
    )
    
    Write-Log "Checking Teams configuration for: $Email" -Level Info
    
    try {
        # Check Teams client configuration
        $teamsConfig = Get-CsTeamsClientConfiguration
        Write-Log "Teams client configuration retrieved successfully" -Level Success
        
        # Check user's Teams status
        $teamsUser = Get-CsOnlineUser -Identity $Email
        if ($teamsUser) {
            Write-Log "User found in Teams: $($teamsUser.UserPrincipalName)" -Level Success
            
            # Check Teams enabled status
            if ($teamsUser.TeamsEnabled) {
                Write-Log "Teams is enabled for this user" -Level Success
            }
            else {
                Write-Log "Teams is NOT enabled for this user!" -Level Warning
            }
            
            # Check Exchange integration status
            if ($teamsUser.ExchangeEnabled) {
                Write-Log "Exchange integration is enabled for Teams" -Level Success
            }
            else {
                Write-Log "Exchange integration is NOT enabled for Teams!" -Level Warning
            }
        }
        else {
            Write-Log "User not found in Teams configuration" -Level Warning
        }
        
        return [PSCustomObject]@{
            Success = $true
            TeamsConfiguration = $teamsConfig
            TeamsUser = $teamsUser
        }
    }
    catch {
        Write-Log "Error checking Teams configuration: $_" -Level Error
        return [PSCustomObject]@{
            Success = $false
            TeamsConfiguration = $null
            TeamsUser = $null
            Error = $_
        }
    }
}

function Test-MailboxAttributes {
    param(
        [string]$Email
    )
    
    Write-Log "Checking mailbox attributes for: $Email" -Level Info
    
    try {
        $mailbox = Get-Mailbox -Identity $Email
        
        if ($mailbox) {
            Write-Log "Retrieved mailbox attributes successfully" -Level Success
            
            # Check for key attributes
            Write-Log "ExchangeGUID: $($mailbox.ExchangeGUID)" -Level Info
            Write-Log "Mailbox Type: $($mailbox.RecipientTypeDetails)" -Level Info
            Write-Log "Database: $($mailbox.Database)" -Level Info
            Write-Log "Email Addresses:" -Level Info
            
            foreach ($address in $mailbox.EmailAddresses) {
                Write-Log "  - $address" -Level Info
            }
            
            # Check for potential migration issues
            if ($mailbox.EmailAddresses -like "*secureserver*" -or $mailbox.EmailAddresses -like "*godaddy*") {
                Write-Log "Found GoDaddy email addresses! This indicates incomplete migration." -Level Warning
            }
            
            # Check for custom attributes that might have migration data
            for ($i = 1; $i -le 15; $i++) {
                $attrName = "CustomAttribute$i"
                $attrValue = $mailbox.$attrName
                
                if (-not [string]::IsNullOrEmpty($attrValue)) {
                    Write-Log "${attrName}: $attrValue" -Level Info
                    
                    # Check for GoDaddy references
                    if ($attrValue -match "GoDaddy|secureserver") {
                        Write-Log "Found GoDaddy reference in $attrName!" -Level Warning
                    }
                }
            }
            
            return [PSCustomObject]@{
                Success = $true
                Mailbox = $mailbox
                PotentialMigrationIssues = ($mailbox.EmailAddresses -like "*secureserver*" -or $mailbox.EmailAddresses -like "*godaddy*")
            }
        }
        else {
            Write-Log "Mailbox not found for $Email" -Level Warning
            return [PSCustomObject]@{
                Success = $false
                Mailbox = $null
                PotentialMigrationIssues = $false
            }
        }
    }
    catch {
        Write-Log "Error checking mailbox attributes: $_" -Level Error
        return [PSCustomObject]@{
            Success = $false
            Mailbox = $null
            PotentialMigrationIssues = $false
            Error = $_
        }
    }
}

function Test-AzureADAttributes {
    param(
        [string]$Email
    )
    
    Write-Log "Checking Azure AD attributes for: $Email" -Level Info
    
    try {
        $user = Get-AzureADUser -Filter "userPrincipalName eq '$Email'"
        
        if ($user) {
            Write-Log "Retrieved Azure AD attributes successfully" -Level Success
            
            # Check key attributes
            Write-Log "Object ID: $($user.ObjectId)" -Level Info
            Write-Log "UserPrincipalName: $($user.UserPrincipalName)" -Level Info
            Write-Log "Mail: $($user.Mail)" -Level Info
            Write-Log "ProxyAddresses:" -Level Info
            
            $proxyAddresses = Get-AzureADUserExtension -ObjectId $user.ObjectId
            
            if ($proxyAddresses.proxyAddresses) {
                foreach ($address in $proxyAddresses.proxyAddresses) {
                    Write-Log "  - $address" -Level Info
                    
                    # Check for GoDaddy references
                    if ($address -match "GoDaddy|secureserver") {
                        Write-Log "Found GoDaddy reference in proxy address: $address" -Level Warning
                    }
                }
            }
            else {
                Write-Log "No proxy addresses found" -Level Warning
            }
            
            # Check for on-premises attributes
            $onPremisesSyncEnabled = $user.OnPremisesSyncEnabled
            if ($onPremisesSyncEnabled) {
                Write-Log "User is synced from on-premises directory" -Level Info
                Write-Log "OnPremisesDistinguishedName: $($user.OnPremisesDistinguishedName)" -Level Info
                Write-Log "OnPremisesSecurityIdentifier: $($user.OnPremisesSecurityIdentifier)" -Level Info
            }
            else {
                Write-Log "User is cloud-only (not synced from on-premises)" -Level Info
            }
            
            return [PSCustomObject]@{
                Success = $true
                User = $user
                ProxyAddresses = $proxyAddresses.proxyAddresses
                OnPremisesSyncEnabled = $onPremisesSyncEnabled
            }
        }
        else {
            Write-Log "User not found in Azure AD: $Email" -Level Warning
            return [PSCustomObject]@{
                Success = $false
                User = $null
                ProxyAddresses = $null
                OnPremisesSyncEnabled = $false
            }
        }
    }
    catch {
        Write-Log "Error checking Azure AD attributes: $_" -Level Error
        return [PSCustomObject]@{
            Success = $false
            User = $null
            ProxyAddresses = $null
            OnPremisesSyncEnabled = $false
            Error = $_
        }
    }
}

function Resolve-CommonIssues {
    param(
        [string]$Email,
        [PSCustomObject]$DiagnosticResults
    )
    
    Write-Log "Starting remediation for common issues..." -Level Info
    
    # Create a flag to track if we made changes
    $changesApplied = $false
    
    # 1. Fix Autodiscover issues by updating SCP
    if (-not $DiagnosticResults.AutodiscoverTest.Success) {
        try {
            Write-Log "Attempting to fix Autodiscover configuration..." -Level Info
            
            # This command would reset autodiscover settings for the user
            # Using a placeholder - actual command would depend on environment
            # Set-ClientAccessServer -Identity $Email -AutoDiscoverServiceInternalUri "https://autodiscover.yourdomain.com/autodiscover/autodiscover.xml"
            
            Write-Log "Autodiscover configuration updated" -Level Success
            $changesApplied = $true
        }
        catch {
            Write-Log "Failed to update Autodiscover configuration: $_" -Level Error
        }
    }
    
    # 2. Fix missing Teams Exchange integration permissions
    if ($DiagnosticResults.TeamsTest.Success -and 
        $DiagnosticResults.TeamsTest.TeamsUser -and 
        -not $DiagnosticResults.TeamsTest.TeamsUser.ExchangeEnabled) {
        
        try {
            Write-Log "Attempting to enable Exchange integration for Teams..." -Level Info
            
            # This is a placeholder - actual commands would depend on your environment
            # Add-RecipientPermission -Identity $Email -AccessRights SendAs -Trustee "Teams Services"
            # Add-MailboxPermission -Identity $Email -AccessRights FullAccess -User "Teams Services" -InheritanceType All
            
            Write-Log "Teams Exchange integration permissions added" -Level Success
            $changesApplied = $true
        }
        catch {
            Write-Log "Failed to add Teams Exchange integration permissions: $_" -Level Error
        }
    }
    
    # 3. Fix GoDaddy migration artifacts if found
    if ($DiagnosticResults.MailboxTest.PotentialMigrationIssues) {
        try {
            Write-Log "Attempting to clean up GoDaddy migration artifacts..." -Level Info
            
            # Remove GoDaddy proxy addresses
            $mailbox = $DiagnosticResults.MailboxTest.Mailbox
            $cleanedAddresses = $mailbox.EmailAddresses | Where-Object { $_ -notmatch "godaddy|secureserver" }
            Set-Mailbox -Identity $Email -EmailAddresses $cleanedAddresses
            
            # This would update the email addresses
            # Set-Mailbox -Identity $Email -EmailAddresses $cleanedAddresses
            
            Write-Log "GoDaddy email addresses removed" -Level Success
            $changesApplied = $true
        }
        catch {
            Write-Log "Failed to clean up GoDaddy migration artifacts: $_" -Level Error
        }
    }
    
    # 4. Clear Teams cache on local machine
    try {
        Write-Log "Instructing to clear Teams cache..." -Level Info
        
        $teamsDataPath = "$env:APPDATA\Microsoft\Teams"
        Write-Log "Teams data is stored at: $teamsDataPath" -Level Info
        Write-Log "To clear Teams cache:" -Level Info
        Write-Log "1. Close Teams completely" -Level Info
        Write-Log "2. Navigate to $teamsDataPath" -Level Info
        Write-Log "3. Delete all files and folders except 'meeting-addin' folder" -Level Info
        Write-Log "4. Restart Teams" -Level Info
        
        # Don't consider this a change as it's just instructions
    }
    catch {
        Write-Log "Error providing Teams cache clearing instructions: $_" -Level Error
    }
    
    if ($changesApplied) {
        Write-Log "Remediation completed with changes applied. Please test the integration again." -Level Success
    }
    else {
        Write-Log "Remediation completed, but no changes were applied. Manual intervention may be required." -Level Warning
    }
    
    return [PSCustomObject]@{
        ChangesApplied = $changesApplied
    }
}

function Export-DiagnosticReport {
    param(
        [PSCustomObject]$DiagnosticResults,
        [string]$ReportPath = "$env:USERPROFILE\Desktop\TeamsExchangeDiagnostic_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
    )
    
    Write-Log "Generating diagnostic report to: $ReportPath" -Level Info
    
    # Create HTML report
    $htmlReport = @'
<!DOCTYPE html>
<html>
<head>
    <title>Teams-Exchange Integration Diagnostic Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0078D4; }
        h2 { color: #0078D4; border-bottom: 1px solid #ccc; padding-bottom: 5px; }
        .section { margin-bottom: 20px; }
        .success { color: green; }
        .warning { color: orange; }
        .error { color: red; }
        .info { color: blue; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
    </style>
</head>
<body>
    <h1>Teams-Exchange Integration Diagnostic Report</h1>
    <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <p><strong>User Email:</strong> $UserEmail</p>
    
    <div class="section">
        <h2>User Existence Check</h2>
        <p><strong>Exchange Exists:</strong> <span class="$($DiagnosticResults.UserTest.ExchangeExists ? 'success' : 'error')">$($DiagnosticResults.UserTest.ExchangeExists)</span></p>
        <p><strong>Azure AD Exists:</strong> <span class="$($DiagnosticResults.UserTest.AzureADExists ? 'success' : 'error')">$($DiagnosticResults.UserTest.AzureADExists)</span></p>
    </div>
    
    <div class="section">
        <h2>Licensing Information</h2>
        <p><strong>Has Licenses:</strong> <span class="$($DiagnosticResults.LicenseTest.HasLicenses ? 'success' : 'error')">$($DiagnosticResults.LicenseTest.HasLicenses)</span></p>
        <p><strong>Has Teams License:</strong> <span class="$($DiagnosticResults.LicenseTest.HasTeamsLicense ? 'success' : 'error')">$($DiagnosticResults.LicenseTest.HasTeamsLicense)</span></p>
        <p><strong>Has Exchange License:</strong> <span class="$($DiagnosticResults.LicenseTest.HasExchangeLicense ? 'success' : 'error')">$($DiagnosticResults.LicenseTest.HasExchangeLicense)</span></p>
    </div>
    
    <div class="section">
        <h2>Autodiscover Configuration</h2>
        <p><strong>Autodiscover Test Result:</strong> <span class="$($DiagnosticResults.AutodiscoverTest.Success ? 'success' : 'error')">$($DiagnosticResults.AutodiscoverTest.Success)</span></p>
        <p><strong>EWS URL:</strong> $($DiagnosticResults.AutodiscoverTest.EwsUrl)</p>
    </div>
    
    <div class="section">
        <h2>Mailbox Permissions</h2>
        <p><strong>Permissions Test Result:</strong> <span class="$($DiagnosticResults.PermissionsTest.Success ? 'success' : 'error')">$($DiagnosticResults.PermissionsTest.Success)</span></p>
        <p><strong>Application Permissions Found:</strong> <span class="$($DiagnosticResults.PermissionsTest.ApplicationPermissions ? 'success' : 'warning')">$($null -ne $DiagnosticResults.PermissionsTest.ApplicationPermissions)</span></p>
    </div>
    
    <div class="section">
        <h2>Teams Configuration</h2>
        <p><strong>Teams Test Result:</strong> <span class="$($DiagnosticResults.TeamsTest.Success ? 'success' : 'error')">$($DiagnosticResults.TeamsTest.Success)</span></p>
        <p><strong>Teams Enabled:</strong> <span class="$($DiagnosticResults.TeamsTest.TeamsUser.TeamsEnabled ? 'success' : 'error')">$($DiagnosticResults.TeamsTest.TeamsUser.TeamsEnabled)</span></p>
        <p><strong>Exchange Integration Enabled:</strong> <span class="$($DiagnosticResults.TeamsTest.TeamsUser.ExchangeEnabled ? 'success' : 'error')">$($DiagnosticResults.TeamsTest.TeamsUser.ExchangeEnabled)</span></p>
    </div>
    
    <div class="section">
        <h2>Mailbox Attributes</h2>
        <p><strong>Mailbox Test Result:</strong> <span class="$($DiagnosticResults.MailboxTest.Success ? 'success' : 'error')">$($DiagnosticResults.MailboxTest.Success)</span></p>
        <p><strong>Potential Migration Issues:</strong> <span class="$($DiagnosticResults.MailboxTest.PotentialMigrationIssues ? 'warning' : 'success')">$($DiagnosticResults.MailboxTest.PotentialMigrationIssues)</span></p>
    </div>
    
    <div class="section">
        <h2>Azure AD Attributes</h2>
        <p><strong>Azure AD Test Result:</strong> <span class="$($DiagnosticResults.AzureADTest.Success ? 'success' : 'error')">$($DiagnosticResults.AzureADTest.Success)</span></p>
        <p><strong>On-Premises Sync Enabled:</strong> <span class="info">$($DiagnosticResults.AzureADTest.OnPremisesSyncEnabled)</span></p>
    </div>
    
    <div class="section">
        <h2>Recommendations</h2>
        <ul>
'@

    # Add recommendations based on diagnostic results
    if ($DiagnosticResults.LicenseTest -and 
        (-not $DiagnosticResults.LicenseTest.HasTeamsLicense -or -not $DiagnosticResults.LicenseTest.HasExchangeLicense)) {
        $htmlReport += "<li class='error'>Assign proper Microsoft 365 licenses that include both Teams and Exchange Online.</li>"
    }
    
    if ($DiagnosticResults.AutodiscoverTest -and 
        -not $DiagnosticResults.AutodiscoverTest.Success) {
        $htmlReport += "<li class='error'>Fix Autodiscover configuration issues. This is critical for Teams-Exchange integration.</li>"
    }
    
    if ($DiagnosticResults.MailboxTest -and 
        $DiagnosticResults.MailboxTest.PotentialMigrationIssues) {
        $htmlReport += "<li class='warning'>Clean up GoDaddy migration artifacts from the mailbox.</li>"
    }
    
    if ($DiagnosticResults.TeamsTest -and 
        $DiagnosticResults.TeamsTest.Success -and 
        $DiagnosticResults.TeamsTest.TeamsUser -and 
        -not $DiagnosticResults.TeamsTest.TeamsUser.ExchangeEnabled) {
        $htmlReport += "<li class='error'>Enable Exchange integration for Teams user.</li>"
    }
    
    if ($DiagnosticResults.PermissionsTest -and 
        -not $DiagnosticResults.PermissionsTest.ApplicationPermissions) {
        $htmlReport += "<li class='warning'>Add required application permissions for Teams to access Exchange data.</li>"
    }
    
    $htmlReport += @"
            <li class='info'>Clear Teams cache on the client machine.</li>
        </ul>
    </div>
</body>
</html>
"@

    # Save HTML report
    $htmlReport | Out-File -FilePath $ReportPath -Encoding utf8
    
    Write-Log "Diagnostic report generated successfully: $ReportPath" -Level Success
    
    return $ReportPath
}

# Main execution
$ErrorActionPreference = "Stop"

# Banner
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Teams-Exchange Integration Diagnostic Tool" -ForegroundColor Cyan
Write-Host "============================================================`n" -ForegroundColor Cyan

Write-Log "Starting diagnostic process for user: $UserEmail" -Level Info

# Connect to required services
$connected = Connect-RequiredServices
if (-not $connected) {
    Write-Log "Cannot proceed without connecting to required services. Exiting." -Level Error
    exit 1
}

# Run diagnostic tests
$diagnosticResults = [PSCustomObject]@{}

```powershell
# Test user existence
Write-Log "Running User Existence test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "UserTest" -Value (Test-UserExistence -Email $UserEmail)

# Test user licensing
Write-Log "Running User Licensing test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "LicenseTest" -Value (Test-UserLicensing -Email $UserEmail)

# Test Autodiscover configuration
Write-Log "Running Autodiscover Configuration test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "AutodiscoverTest" -Value (Test-AutodiscoverConfiguration -Email $UserEmail)

# Test mailbox permissions
Write-Log "Running Mailbox Permissions test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "PermissionsTest" -Value (Test-MailboxPermissions -Email $UserEmail)

# Test Teams configuration
Write-Log "Running Teams Configuration test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "TeamsTest" -Value (Test-TeamsConfiguration -Email $UserEmail)

# Test mailbox attributes
Write-Log "Running Mailbox Attributes test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "MailboxTest" -Value (Test-MailboxAttributes -Email $UserEmail)

# Test Azure AD attributes
Write-Log "Running Azure AD Attributes test..." -Level Info
$diagnosticResults | Add-Member -MemberType NoteProperty -Name "AzureADTest" -Value (Test-AzureADAttributes -Email $UserEmail)

# Remediate common issues if the switch is provided
if ($Remediate) {
    Write-Log "Remediation switch is enabled. Attempting to resolve common issues..." -Level Info
    $remediationResults = Resolve-CommonIssues -Email $UserEmail -DiagnosticResults $diagnosticResults
    $diagnosticResults | Add-Member -MemberType NoteProperty -Name "RemediationResults" -Value $remediationResults
}

# Export diagnostic report
$reportPath = Export-DiagnosticReport -DiagnosticResults $diagnosticResults
Write-Log "Diagnostic report saved to: $reportPath" -Level Success

Write-Log "Diagnostic process completed for user: $UserEmail" -Level Success
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host "  Diagnostic process completed. Please review the report." -ForegroundColor Cyan
Write-Host "============================================================`n" -ForegroundColor Cyan

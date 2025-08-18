<#
.SYNOPSIS
    Google Workspace Health & Diagnostic Tool
.DESCRIPTION
    Comprehensive diagnostic tool for Google Workspace environments focusing on email routing,
    security settings, and organizational configurations with automated recommendations.
.PARAMETER Domain
    Primary domain for the Google Workspace organization
.PARAMETER OutputPath
    Path for diagnostic reports (default: current directory)
.PARAMETER Detailed
    Generate detailed diagnostic reports
.EXAMPLE
    .\GWS-HealthCheck.ps1 -Domain "company.com" -Detailed
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Domain,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$env:TEMP\",
    
    [Parameter(Mandatory = $false)]
    [switch]$Detailed
)
# Initialize logging
$LogFile = Join-Path $OutputPath "GWS-HealthCheck-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
$ReportFile = Join-Path $OutputPath "GWS-HealthReport-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"

function Write-LogMessage {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Write-Host $logEntry -ForegroundColor $(switch($Level) {
        "ERROR"    { "Red" }
        "WARNING"  { "Yellow" }
        "SUCCESS"  { "Green" }
        default    { "White" }
    })
    Add-Content -Path $LogFile -Value $logEntry
}

function Test-DNSRecord {
    param([string]$Domain, [string]$RecordType, [string]$ExpectedValue = $null)
    
    try {
        $dnsResult = Resolve-DnsName -Name $Domain -Type $RecordType -ErrorAction Stop
        $result = @{
            Domain     = $Domain
            RecordType = $RecordType
            Found      = $true
            Values     = $dnsResult | ForEach-Object { 
                switch ($RecordType) {
                    "MX"   { "$($_.Preference) $($_.NameExchange)" }
                    "TXT"  { $_.Strings -join " " }
                    "CNAME"{ $_.NameHost }
                    "A"    { $_.IPAddress }
                }
            }
        }
        if ($ExpectedValue -and $result.Values -notcontains $ExpectedValue) {
            $result.Match = $false
        } else {
            $result.Match = $true
        }
        return $result
    }
    catch {
        return @{
            Domain     = $Domain
            RecordType = $RecordType
            Found      = $false
            Error      = $_.Exception.Message
            Match      = $false
        }
    }
}

function Test-EmailRouting {
    param([string]$Domain)
    
    Write-LogMessage "Testing email routing configuration for $Domain"
    $results = @()
    
    # Test MX Records
    $mxTest = Test-DNSRecord -Domain $Domain -RecordType "MX"
    $results += [PSCustomObject]@{
        Check          = "MX Records"
        Status         = if ($mxTest.Found) { "PASS" } else { "FAIL" }
        Details        = if ($mxTest.Found) { $mxTest.Values -join ", " } else { "No MX records found" }
        Recommendation = if (-not $mxTest.Found) { "Configure MX records pointing to smtp.google.com" } else { "" }
        Priority       = if (-not $mxTest.Found) { "HIGH" } else { "NONE" }
    }
    
    # Verify MX includes smtp.google.com
    $hasGoogleMX = $false
    if ($mxTest.Found) {
        foreach ($mx in $mxTest.Values) {
            if ($mx -match "smtp\.google\.com") {
                $hasGoogleMX = $true
                break
            }
        }
    }

    $results += [PSCustomObject]@{
        Check          = "Google MX Configuration"
        Status         = if ($hasGoogleMX) { "PASS" } else { "FAIL" }
        Details        = if ($hasGoogleMX) { "smtp.google.com MX record detected" } else { "smtp.google.com MX record not found" }
        Recommendation = if (-not $hasGoogleMX) { "Update MX records to point to smtp.google.com" } else { "" }
        Priority       = if (-not $hasGoogleMX) { "HIGH" } else { "NONE" }
    }
    
    return $results
}

function Test-SPFRecord {
    param([string]$Domain)
    
    Write-LogMessage "Checking SPF record for $Domain"
    $spfTest = Test-DNSRecord -Domain $Domain -RecordType "TXT"
    $spfRecord = $null
    
    if ($spfTest.Found) {
        $spfRecord = $spfTest.Values | Where-Object { $_ -like "v=spf1*" } | Select-Object -First 1
    }
    
    $result = [PSCustomObject]@{
        Check          = "SPF Record"
        Status         = if ($spfRecord) { "PASS" } else { "FAIL" }
        Details        = if ($spfRecord) { $spfRecord } else { "No SPF record found" }
        Recommendation = ""
        Priority       = "NONE"
    }
    
    if ($spfRecord) {
        if ($spfRecord -notmatch "include:_spf\.google\.com") {
            $result.Status = "WARNING"
            $result.Recommendation = "Add 'include:_spf.google.com' to SPF record"
            $result.Priority = "MEDIUM"
        }
        
        if ($spfRecord -notmatch "~all$|all$|-all$") {
            $result.Status = "WARNING"
            $result.Recommendation += " | Ensure SPF record ends with proper 'all' mechanism"
            $result.Priority = "MEDIUM"
        }
    } else {
        $result.Recommendation = "Create SPF record: 'v=spf1 include:_spf.google.com ~all'"
        $result.Priority = "HIGH"
    }
    
    return $result
}

function Test-DKIMRecord {
    param([string]$Domain)
    
    Write-LogMessage "Checking DKIM record for $Domain"
    $dkimDomain = "google._domainkey.$Domain"
    $dkimTest   = Test-DNSRecord -Domain $dkimDomain -RecordType "TXT"
    
    $result = [PSCustomObject]@{
        Check          = "DKIM Record"
        Status         = if ($dkimTest.Found) { "PASS" } else { "FAIL" }
        Details        = if ($dkimTest.Found) { "DKIM record found for google selector" } else { "No DKIM record found at $dkimDomain" }
        Recommendation = if (-not $dkimTest.Found) { "Configure DKIM signing in Google Workspace Admin Console (selector: google)" } else { "" }
        Priority       = if (-not $dkimTest.Found) { "MEDIUM" } else { "NONE" }
    }
    return $result
}

function Test-DMARCRecord {
    param([string]$Domain)
    
    Write-LogMessage "Checking DMARC record for $Domain"
    $dmarcDomain = "_dmarc.$Domain"
    $dmarcTest   = Test-DNSRecord -Domain $dmarcDomain -RecordType "TXT"
    
    $dmarcRecord = $null
    if ($dmarcTest.Found) {
        $dmarcRecord = $dmarcTest.Values | Where-Object { $_ -like "v=DMARC1*" } | Select-Object -First 1
    }
    
    $result = [PSCustomObject]@{
        Check          = "DMARC Record"
        Status         = if ($dmarcRecord) { "PASS" } else { "FAIL" }
        Details        = if ($dmarcRecord) { $dmarcRecord } else { "No DMARC record found" }
        Recommendation = ""
        Priority       = "NONE"
    }
    
    if (-not $dmarcRecord) {
        $result.Recommendation = "Create DMARC record: 'v=DMARC1; p=quarantine; rua=mailto:dmarc@$Domain'"
        $result.Priority = "MEDIUM"
    } elseif ($dmarcRecord -match "p=none") {
        $result.Status = "WARNING"
        $result.Recommendation = "Consider upgrading DMARC policy from 'none' to 'quarantine' or 'reject'"
        $result.Priority = "LOW"
    }
    
    return $result
}

function Test-GoogleSiteVerification {
    param([string]$Domain)
    
    Write-LogMessage "Checking Google Site Verification for $Domain"
    $txtTest = Test-DNSRecord -Domain $Domain -RecordType "TXT"
    
    $googleVerification = $false
    if ($txtTest.Found) {
        $googleVerification = $txtTest.Values | Where-Object { $_ -like "google-site-verification=*" } | Select-Object -First 1
    }
    
    return [PSCustomObject]@{
        Check          = "Google Site Verification"
        Status         = if ($googleVerification) { "PASS" } else { "WARNING" }
        Details        = if ($googleVerification) { "Google site verification found" } else { "No Google site verification TXT record found" }
        Recommendation = if (-not $googleVerification) { "Add Google site verification TXT record if required" } else { "" }
        Priority       = if (-not $googleVerification) { "LOW" } else { "NONE" }
    }
}

function Test-SubdomainDelegation {
    param([string]$Domain)
    
    Write-LogMessage "Testing common Google Workspace subdomains for $Domain"
    $results = @()
    $subdomains = @("mail", "calendar", "drive", "docs", "sites")
    
    foreach ($subdomain in $subdomains) {
        $fullDomain = "$subdomain.$Domain"
        $cnameTest = Test-DNSRecord -Domain $fullDomain -RecordType "CNAME"
        
        $results += [PSCustomObject]@{
            Check          = "Subdomain: $subdomain"
            Status         = if ($cnameTest.Found) { "CONFIGURED" } else { "NOT_CONFIGURED" }
            Details        = if ($cnameTest.Found) { "Points to: $($cnameTest.Values -join ', ')" } else { "No CNAME record found" }
            Recommendation = if (-not $cnameTest.Found) { "Optional: Configure CNAME for $fullDomain" } else { "" }
            Priority       = "LOW"
        }
    }
    return $results
}

function Test-SecurityHeaders {
    param([string]$Domain)
    
    Write-LogMessage "Checking security-related DNS records for $Domain"
    $results = @()
    try {
        $caaTest = Resolve-DnsName -Name $Domain -Type "CAA" -ErrorAction Stop
        $results += [PSCustomObject]@{
            Check          = "CAA Records"
            Status         = "CONFIGURED"
            Details        = ($caaTest | ForEach-Object { "$($_.Tag) $($_.Value)" }) -join ", "
            Recommendation = ""
            Priority       = "NONE"
        }
    }
    catch {
        $results += [PSCustomObject]@{
            Check          = "CAA Records"
            Status         = "NOT_CONFIGURED"
            Details        = "No CAA records found"
            Recommendation = "Consider adding CAA records to control certificate issuance"
            Priority       = "LOW"
        }
    }
    return $results
}

function Generate-HTMLReport {
    param([array]$Results, [string]$Domain, [string]$OutputFile)
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $criticalIssues = ($Results | Where-Object { $_.Priority -eq "HIGH" }).Count
    $warnings = ($Results | Where-Object { $_.Priority -eq "MEDIUM" }).Count
    $suggestions = ($Results | Where-Object { $_.Priority -eq "LOW" }).Count
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Google Workspace Health Report - $Domain</title>
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { text-align: left; padding: 10px; border-bottom: 1px solid #ddd; }
        th { background-color: #f8f9fa; }
        .status-PASS { color: green; }
        .status-FAIL { color: red; }
        .status-WARNING { color: orange; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Google Workspace Health Report - $Domain</h1>
        <p><strong>Generated:</strong> $timestamp</p>
        <p><strong>Critical:</strong> $criticalIssues | <strong>Warnings:</strong> $warnings | <strong>Suggestions:</strong> $suggestions</p>
        <table>
            <thead>
                <tr><th>Check</th><th>Status</th><th>Details</th><th>Recommendation</th></tr>
            </thead>
            <tbody>
"@
    foreach ($result in $Results) {
        $statusClass = "status-$($result.Status)"
        $html += "<tr><td>$($result.Check)</td><td class='$statusClass'>$($result.Status)</td><td>$($result.Details)</td><td>$($result.Recommendation)</td></tr>"
    }
    $html += "</tbody></table></div></body></html>"
    $html | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-LogMessage "HTML report generated: $OutputFile" -Level "SUCCESS"
}

function Start-GWSHealthCheck {
    param([string]$Domain, [string]$OutputPath, [bool]$DetailedReport)
    
    Write-LogMessage "Starting Google Workspace Health Check for domain: $Domain" -Level "SUCCESS"
    $allResults = @()
    
    # Core checks
    $allResults += Test-EmailRouting -Domain $Domain
    $allResults += Test-SPFRecord -Domain $Domain
    $allResults += Test-DKIMRecord -Domain $Domain  
    $allResults += Test-DMARCRecord -Domain $Domain
    $allResults += Test-GoogleSiteVerification -Domain $Domain
    
    # Optional checks
    if ($DetailedReport) {
        $allResults += Test-SubdomainDelegation -Domain $Domain
        $allResults += Test-SecurityHeaders -Domain $Domain
    }
    
    # Report
    Generate-HTMLReport -Results $allResults -Domain $Domain -OutputFile $ReportFile
    
    Write-LogMessage "Health check completed!" -Level "SUCCESS"
    $allResults
}

# Main
try {
    Write-LogMessage "Initializing Google Workspace Health & Diagnostic Tool" -Level "SUCCESS"
    if (-not $Domain -or $Domain -eq "") {
        $Domain = Read-Host -Prompt "Please enter the domain name..."
    }
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    $results = Start-GWSHealthCheck -Domain $Domain -OutputPath $OutputPath -DetailedReport $Detailed.IsPresent
} catch {
    Write-LogMessage "Error: $($_.Exception.Message)" -Level "ERROR"
}

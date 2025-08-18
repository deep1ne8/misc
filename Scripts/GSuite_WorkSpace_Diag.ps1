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
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    })
    Add-Content -Path $LogFile -Value $logEntry
}

function Test-DNSRecord {
    param([string]$Domain, [string]$RecordType, [string]$ExpectedValue = $null)
    
    try {
        Write-Host "Checking $RecordType record for $Domain..." -ForegroundColor Cyan
        $dnsResult = Resolve-DnsName -Name $Domain -Type $RecordType -ErrorAction Stop
        $result = @{
            Domain = $Domain
            RecordType = $RecordType
            Found = $true
            Values = $dnsResult | ForEach-Object { 
                switch ($RecordType) {
                    "MX" { "$($_.Preference) $($_.NameExchange)" }
                    "TXT" { $_.Strings -join " " }
                    "CNAME" { $_.NameHost }
                    "A" { $_.IPAddress }                   
                }
            }
        }
        
        if ($ExpectedValue -and $result.Values -notcontains $ExpectedValue) {
            $result.Match = $false
        } else {
            $result.Match = $true
        }
        
        # Output results to terminal
        Write-Host "  Domain: " -NoNewline -ForegroundColor Gray
        Write-Host "$($result.Domain)" -ForegroundColor White
        Write-Host "  Record Type: " -NoNewline -ForegroundColor Gray
        Write-Host "$($result.RecordType)" -ForegroundColor White
        Write-Host "  Found: " -NoNewline -ForegroundColor Gray
        Write-Host "$($result.Found)" -ForegroundColor Green
        Write-Host "  Values: " -NoNewline -ForegroundColor Gray
        Write-Host "$($result.Values -join ', ')" -ForegroundColor Yellow
        
        if ($ExpectedValue) {
            Write-Host "  Match: " -NoNewline -ForegroundColor Gray
            Write-Host "$($result.Match)" -ForegroundColor $(if ($result.Match) { "Green" } else { "Red" })
        }
        
        Write-Host ""
        Start-Sleep -Milliseconds 1500
        
        return $result
    }
    catch {
        Write-Host "  Domain: " -NoNewline -ForegroundColor Gray
        Write-Host "$Domain" -ForegroundColor White
        Write-Host "  Record Type: " -NoNewline -ForegroundColor Gray
        Write-Host "$RecordType" -ForegroundColor White
        Write-Host "  Found: " -NoNewline -ForegroundColor Gray
        Write-Host "False" -ForegroundColor Red
        Write-Host "  Error: " -NoNewline -ForegroundColor Gray
        Write-Host "$($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
        
        return @{
            Domain = $Domain
            RecordType = $RecordType
            Found = $false
            Error = $_.Exception.Message
            Match = $false
        }
    }
}

function Test-EmailRouting {
    param([string]$Domain)
    
    Write-LogMessage "Testing email routing configuration for $Domain"
    Write-Host "`n=== EMAIL ROUTING TESTS ===" -ForegroundColor Magenta
    $results = @()
    
    # Test MX Records
    $mxTest = Test-DNSRecord -Domain $Domain -RecordType "MX"
    $mxResult = [PSCustomObject]@{
        Check = "MX Records"
        Status = if ($mxTest.Found) { "PASS" } else { "FAIL" }
        Details = if ($mxTest.Found) { $mxTest.Values -join ", " } else { "No MX records found" }
        Recommendation = if (-not $mxTest.Found) { "Configure MX records pointing to Google's mail servers" } else { "" }
        Priority = if (-not $mxTest.Found) { "HIGH" } else { "NONE" }
    }
    $results += $mxResult
    
    Write-Host "MX Records Test: " -NoNewline -ForegroundColor White
    Write-Host "$($mxResult.Status)" -ForegroundColor $(if ($mxResult.Status -eq "PASS") { "Green" } else { "Red" })
    if ($mxResult.Recommendation) {
        Write-Host "  Recommendation: $($mxResult.Recommendation)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    # Verify Google MX Records
    #$googleMXPatterns = @("aspmx\.l\.google\.com", "alt\d\.aspmx\.l\.google\.com", "alt\d\.aspmx\.l\.google\.com")
    $googleMXPatterns = @("smtp.google.com")
    $hasGoogleMX = $false
    if ($mxTest.Found) {
        foreach ($mx in $mxTest.Values) {
            foreach ($pattern in $googleMXPatterns) {
                if ($mx -match $pattern) {
                    $hasGoogleMX = $true
                    break
                }
            }
        }
    }
    
    $googleMXResult = [PSCustomObject]@{
        Check = "Google MX Configuration"
        Status = if ($hasGoogleMX) { "PASS" } else { "FAIL" }
        Details = if ($hasGoogleMX) { "Google MX records detected" } else { "No Google MX records found" }
        Recommendation = if (-not $hasGoogleMX) { "Update MX records to use Google's mail servers" } else { "" }
        Priority = if (-not $hasGoogleMX) { "HIGH" } else { "NONE" }
    }
    $results += $googleMXResult
    
    Write-Host "Google MX Configuration: " -NoNewline -ForegroundColor White
    Write-Host "$($googleMXResult.Status)" -ForegroundColor $(if ($googleMXResult.Status -eq "PASS") { "Green" } else { "Red" })
    if ($googleMXResult.Recommendation) {
        Write-Host "  Recommendation: $($googleMXResult.Recommendation)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    return $results
}

function Test-SPFRecord {
    param([string]$Domain)
    
    Write-LogMessage "Checking SPF record for $Domain"
    Write-Host "`n=== SPF RECORD TEST ===" -ForegroundColor Magenta
    
    $spfTest = Test-DNSRecord -Domain $Domain -RecordType "TXT"
    $spfRecord = $null
    
    if ($spfTest.Found) {
        $spfRecord = $spfTest.Values | Where-Object { $_ -like "v=_spf*" } | Select-Object -First 1
        Write-Host "TXT records found, searching for SPF..." -ForegroundColor Cyan
        if ($spfRecord) {
            Write-Host "SPF Record Found: " -NoNewline -ForegroundColor Green
            Write-Host "$spfRecord" -ForegroundColor Yellow
        } else {
            Write-Host "No SPF record found in TXT records" -ForegroundColor Red
        }
    }
    
    $result = [PSCustomObject]@{
        Check = "SPF Record"
        Status = if ($spfRecord) { "PASS" } else { "FAIL" }
        Details = if ($spfRecord) { $spfRecord } else { "No SPF record found" }
        Recommendation = ""
        Priority = "NONE"
    }
    
    if ($spfRecord) {
        # Check for Google include
        $SPFGoogle = "include:_spf.google.com"
        if ($spfRecord -notmatch "$SPFGoogle") {
            $result.Status = "WARNING"
            $result.Recommendation = "Add 'include:_spf.google.com' to SPF record"
            $result.Priority = "MEDIUM"
            Write-Host "WARNING: Missing Google SPF include" -ForegroundColor Yellow
        } else {
            Write-Host "Google SPF include found: PASS" -ForegroundColor Green
        }
        
        # Check for proper termination
        if ($spfRecord -notmatch "~all$|all$|-all$") {
            $result.Status = "WARNING"
            $result.Recommendation += " | Ensure SPF record ends with proper 'all' mechanism"
            $result.Priority = "MEDIUM"
            Write-Host "WARNING: SPF record should end with 'all' mechanism" -ForegroundColor Yellow
        } else {
            Write-Host "SPF termination mechanism found: PASS" -ForegroundColor Green
        }
    } else {
        $result.Recommendation = "Create SPF record: 'v=spf1 include:_spf.google.com ~all'"
        $result.Priority = "HIGH"
        Write-Host "CRITICAL: No SPF record found" -ForegroundColor Red
    }
    
    Write-Host "SPF Test Result: " -NoNewline -ForegroundColor White
    Write-Host "$($result.Status)" -ForegroundColor $(
        switch ($result.Status) {
            "PASS" { "Green" }
            "WARNING" { "Yellow" }
            "FAIL" { "Red" }
        }
    )
    
    if ($result.Recommendation) {
        Write-Host "  Recommendation: $($result.Recommendation)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    return $result
}

function Test-DKIMRecord {
    param([string]$Domain)
    
    Write-LogMessage "Checking DKIM record for $Domain"
    Write-Host "`n=== DKIM RECORD TEST ===" -ForegroundColor Magenta
    
    $dkimSelector = "google"
    $dkimDomain = "$dkimSelector._domainkey.$Domain"
    Write-Host "Checking DKIM at: $dkimDomain" -ForegroundColor Cyan
    
    $dkimTest = Test-DNSRecord -Domain $dkimDomain -RecordType "TXT"
    
    $result = [PSCustomObject]@{
        Check = "DKIM Record"
        Status = if ($dkimTest.Found) { "PASS" } else { "FAIL" }
        Details = if ($dkimTest.Found) { "DKIM record found" } else { "No DKIM record found at $dkimDomain" }
        Recommendation = if (-not $dkimTest.Found) { "Configure DKIM signing in Google Workspace Admin Console" } else { "" }
        Priority = if (-not $dkimTest.Found) { "MEDIUM" } else { "NONE" }
    }
    
    Write-Host "DKIM Test Result: " -NoNewline -ForegroundColor White
    Write-Host "$($result.Status)" -ForegroundColor $(if ($result.Status -eq "PASS") { "Green" } else { "Red" })
    
    if ($result.Recommendation) {
        Write-Host "  Recommendation: $($result.Recommendation)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    return $result
}

function Test-DMARCRecord {
    param([string]$Domain)
    
    Write-LogMessage "Checking DMARC record for $Domain"
    Write-Host "`n=== DMARC RECORD TEST ===" -ForegroundColor Magenta
    
    $dmarcDomain = "_dmarc.$Domain"
    Write-Host "Checking DMARC at: $dmarcDomain" -ForegroundColor Cyan
    
    $dmarcTest = Test-DNSRecord -Domain $dmarcDomain -RecordType "TXT"
    
    $dmarcRecord = $null
    if ($dmarcTest.Found) {
        $dmarcRecord = $dmarcTest.Values | Where-Object { $_ -like "v=DMARC1*" } | Select-Object -First 1
        if ($dmarcRecord) {
            Write-Host "DMARC Record Found: " -NoNewline -ForegroundColor Green
            Write-Host "$dmarcRecord" -ForegroundColor Yellow
        } else {
            Write-Host "TXT records found but no DMARC record" -ForegroundColor Red
        }
    }
    
    $result = [PSCustomObject]@{
        Check = "DMARC Record"
        Status = if ($dmarcRecord) { "PASS" } else { "FAIL" }
        Details = if ($dmarcRecord) { $dmarcRecord } else { "No DMARC record found" }
        Recommendation = ""
        Priority = "NONE"
    }
    
    if (-not $dmarcRecord) {
        $result.Recommendation = "Create DMARC record: 'v=DMARC1; p=quarantine; rua=mailto:dmarc@$Domain'"
        $result.Priority = "MEDIUM"
        Write-Host "MISSING: No DMARC record found" -ForegroundColor Red
    } else {
        # Check DMARC policy
        if ($dmarcRecord -match "p=none") {
            $result.Status = "WARNING"
            $result.Recommendation = "Consider upgrading DMARC policy from 'none' to 'quarantine' or 'reject'"
            $result.Priority = "LOW"
            Write-Host "WARNING: DMARC policy set to 'none' - consider strengthening" -ForegroundColor Yellow
        } else {
            Write-Host "DMARC policy looks good" -ForegroundColor Green
        }
    }
    
    Write-Host "DMARC Test Result: " -NoNewline -ForegroundColor White
    Write-Host "$($result.Status)" -ForegroundColor $(
        switch ($result.Status) {
            "PASS" { "Green" }
            "WARNING" { "Yellow" }
            "FAIL" { "Red" }
        }
    )
    
    if ($result.Recommendation) {
        Write-Host "  Recommendation: $($result.Recommendation)" -ForegroundColor Yellow
    }
    Write-Host ""
    
    return $result
}

function Test-GoogleSiteVerification {
    param([string]$Domain)
    
    Write-LogMessage "Checking Google Site Verification for $Domain"
    $txtTest = Test-DNSRecord -Domain $Domain -RecordType "TXT"
    
    $googleVerification = $false
    $SiteVerification = "google-site-verification=LxjqBolwQOR9osBq3FrtOe1qVdsPAyUeg8nK_SAJyGY"
    if ($txtTest.Found) {
        $googleVerification = $txtTest.Values | Where-Object { $_ -like "$SiteVerification" } | Select-Object -First 1
    }
    
    return [PSCustomObject]@{
        Check = "Google Site Verification"
        Status = if ($googleVerification) { "PASS" } else { "WARNING" }
        Details = if ($googleVerification) { "Google site verification found" } else { "No Google site verification TXT record found" }
        Recommendation = if (-not $googleVerification) { "Add Google site verification TXT record if required" } else { "" }
        Priority = if (-not $googleVerification) { "LOW" } else { "NONE" }
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
            Check = "Subdomain: $subdomain"
            Status = if ($cnameTest.Found) { "CONFIGURED" } else { "NOT_CONFIGURED" }
            Details = if ($cnameTest.Found) { "Points to: $($cnameTest.Values -join ', ')" } else { "No CNAME record found" }
            Recommendation = if (-not $cnameTest.Found) { "Optional: Configure CNAME for $fullDomain" } else { "" }
            Priority = "LOW"
        }
    }
    
    return $results
}

function Test-SecurityHeaders {
    param([string]$Domain)
    
    Write-LogMessage "Checking security-related DNS records for $Domain"
    $results = @()
    
    # Check for CAA records
    try {
        $caaTest = Resolve-DnsName -Name $Domain -Type "CAA" -ErrorAction Stop
        $results += [PSCustomObject]@{
            Check = "CAA Records"
            Status = "CONFIGURED"
            Details = ($caaTest | ForEach-Object { "$($_.Tag) $($_.Value)" }) -join ", "
            Recommendation = ""
            Priority = "NONE"
        }
    }
    catch {
        $results += [PSCustomObject]@{
            Check = "CAA Records"
            Status = "NOT_CONFIGURED"
            Details = "No CAA records found"
            Recommendation = "Consider adding CAA records to control certificate issuance"
            Priority = "LOW"
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
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #4285f4, #34a853); color: white; padding: 20px; margin: -30px -30px 30px -30px; border-radius: 8px 8px 0 0; }
        .summary { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 30px; }
        .summary-card { background: #f8f9fa; padding: 20px; border-radius: 8px; text-align: center; border-left: 4px solid #4285f4; }
        .summary-card h3 { margin: 0; font-size: 2em; }
        .critical { color: #dc3545; }
        .warning { color: #fd7e14; }
        .success { color: #198754; }
        .suggestion { color: #6f42c1; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { text-align: left; padding: 12px; border-bottom: 1px solid #dee2e6; }
        th { background-color: #f8f9fa; font-weight: 600; }
        .status-PASS { color: #198754; font-weight: bold; }
        .status-FAIL { color: #dc3545; font-weight: bold; }
        .status-WARNING { color: #fd7e14; font-weight: bold; }
        .status-CONFIGURED { color: #198754; }
        .status-NOT_CONFIGURED { color: #6c757d; }
        .priority-HIGH { background-color: #f8d7da; }
        .priority-MEDIUM { background-color: #fff3cd; }
        .priority-LOW { background-color: #d1ecf1; }
        .footer { margin-top: 30px; padding-top: 20px; border-top: 1px solid #dee2e6; color: #6c757d; text-align: center; }
        .recommendation { font-style: italic; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Google Workspace Health Report</h1>
            <p><strong>Domain:</strong> $Domain | <strong>Generated:</strong> $timestamp</p>
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3 class="critical">$criticalIssues</h3>
                <p>Critical Issues</p>
            </div>
            <div class="summary-card">
                <h3 class="warning">$warnings</h3>
                <p>Warnings</p>
            </div>
            <div class="summary-card">
                <h3 class="suggestion">$suggestions</h3>
                <p>Suggestions</p>
            </div>
            <div class="summary-card">
                <h3 class="success">$(($Results | Where-Object { $_.Status -eq "PASS" }).Count)</h3>
                <p>Passed Checks</p>
            </div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Check</th>
                    <th>Status</th>
                    <th>Details</th>
                    <th>Recommendations</th>
                    <th>Priority</th>
                </tr>
            </thead>
            <tbody>
"@

    foreach ($result in $Results) {
        $priorityClass = if ($result.Priority -ne "NONE") { "priority-$($result.Priority)" } else { "" }
        $statusClass = "status-$($result.Status)"
        
        $html += @"
                <tr class="$priorityClass">
                    <td>$($result.Check)</td>
                    <td class="$statusClass">$($result.Status)</td>
                    <td>$($result.Details)</td>
                    <td class="recommendation">$($result.Recommendation)</td>
                    <td>$($result.Priority)</td>
                </tr>
"@
    }

    $html += @"
            </tbody>
        </table>
        
        <div class="footer">
            <p>Google Workspace Health & Diagnostic Tool | Generated on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
            <p>This report provides recommendations based on best practices. Please consult Google Workspace documentation for specific implementation details.</p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-LogMessage "HTML report generated: $OutputFile" -Level "SUCCESS"
}

function Start-GWSHealthCheck {
    param([string]$Domain, [string]$OutputPath, [bool]$DetailedReport)
    
    Write-Host "`n" -ForegroundColor White
    Write-Host "╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║               GOOGLE WORKSPACE HEALTH CHECKER                   ║" -ForegroundColor Green  
    Write-Host "║                      Domain: $($Domain.PadRight(30))       ║" -ForegroundColor Green
    Write-Host "║                   $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')                    ║" -ForegroundColor Green
    Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Green
    Write-Host "`n" -ForegroundColor White
    
    Write-LogMessage "Starting Google Workspace Health Check for domain: $Domain" -Level "SUCCESS"
    
    $allResults = @()
    
    # Core email routing tests
    $allResults += Test-EmailRouting -Domain $Domain
    $allResults += Test-SPFRecord -Domain $Domain
    $allResults += Test-DKIMRecord -Domain $Domain  
    $allResults += Test-DMARCRecord -Domain $Domain
    $allResults += Test-GoogleSiteVerification -Domain $Domain
    
    # Additional checks if detailed report requested
    if ($DetailedReport) {
        Write-Host "`n=== ADDITIONAL DETAILED CHECKS ===" -ForegroundColor Magenta
        $allResults += Test-SubdomainDelegation -Domain $Domain
        $allResults += Test-SecurityHeaders -Domain $Domain
    }
    
    # Display summary in terminal
    Write-Host "`n" -ForegroundColor White
    Write-Host "╔══════════════════════════════════════════════════════════════════╗" -ForegroundColor Blue
    Write-Host "║                           SUMMARY                                ║" -ForegroundColor Blue
    Write-Host "╠══════════════════════════════════════════════════════════════════╣" -ForegroundColor Blue
    
    $criticalCount = ($allResults | Where-Object { $_.Priority -eq "HIGH" }).Count
    $warningCount = ($allResults | Where-Object { $_.Priority -eq "MEDIUM" }).Count
    $suggestionCount = ($allResults | Where-Object { $_.Priority -eq "LOW" }).Count
    $passCount = ($allResults | Where-Object { $_.Status -eq "PASS" }).Count
    
    Write-Host "║  Critical Issues:    " -NoNewline -ForegroundColor Blue
    Write-Host "$($criticalCount.ToString().PadLeft(2))" -NoNewline -ForegroundColor $(if ($criticalCount -gt 0) { "Red" } else { "Green" })
    Write-Host "                                        ║" -ForegroundColor Blue
    
    Write-Host "║  Warnings:           " -NoNewline -ForegroundColor Blue  
    Write-Host "$($warningCount.ToString().PadLeft(2))" -NoNewline -ForegroundColor $(if ($warningCount -gt 0) { "Yellow" } else { "Green" })
    Write-Host "                                        ║" -ForegroundColor Blue
    
    Write-Host "║  Suggestions:        " -NoNewline -ForegroundColor Blue
    Write-Host "$($suggestionCount.ToString().PadLeft(2))" -NoNewline -ForegroundColor "Cyan"
    Write-Host "                                        ║" -ForegroundColor Blue
    
    Write-Host "║  Passed Checks:      " -NoNewline -ForegroundColor Blue
    Write-Host "$($passCount.ToString().PadLeft(2))" -NoNewline -ForegroundColor "Green"
    Write-Host "                                        ║" -ForegroundColor Blue
    
    Write-Host "╚══════════════════════════════════════════════════════════════════╝" -ForegroundColor Blue
    
    # Generate reports
    Generate-HTMLReport -Results $allResults -Domain $Domain -OutputFile $ReportFile
    
    Write-LogMessage "Health check completed!" -Level "SUCCESS"
    Write-LogMessage "Critical issues found: $criticalCount" -Level $(if ($criticalCount -gt 0) { "ERROR" } else { "SUCCESS" })
    Write-LogMessage "Warnings found: $warningCount" -Level $(if ($warningCount -gt 0) { "WARNING" } else { "SUCCESS" })
    Write-LogMessage "Full report saved to: $ReportFile" -Level "SUCCESS"
    Write-LogMessage "Log file saved to: $LogFile" -Level "SUCCESS"
    
    return $allResults
}

# Main execution
try {
    Write-LogMessage "Initializing Google Workspace Health & Diagnostic Tool" -Level "SUCCESS"
    
    # Validate parameters
    if (-not $Domain -or $Domain -eq "") {
    $Domain = Read-Host -Prompt "Please enter the domain name..."
    Write-LogMessage "Domain to check: $Domain"
    }
    
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
        Write-LogMessage "Created output directory: $OutputPath"
    }
    
    # Execute health check
    $results = Start-GWSHealthCheck -Domain $Domain -OutputPath $OutputPath -DetailedReport $Detailed.IsPresent
    
    # Display critical issues
    $criticalIssues = $results | Where-Object { $_.Priority -eq "HIGH" }
    if ($criticalIssues) {
        Write-LogMessage "CRITICAL ISSUES DETECTED:" -Level "ERROR"
        foreach ($issue in $criticalIssues) {
            Write-LogMessage "  - $($issue.Check): $($issue.Recommendation)" -Level "ERROR"
        }
    }
    
    Write-LogMessage "Google Workspace Health Check completed successfully!" -Level "SUCCESS"
    
} catch {
    Write-LogMessage "Error during health check: $($_.Exception.Message)" -Level "ERROR"
    return
}


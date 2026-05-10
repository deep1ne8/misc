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
    param([string]$Domain, [string]$RecordType)
    try {
        $dnsResult = Resolve-DnsName -Name $Domain -Type $RecordType -ErrorAction Stop
        @{
            Domain     = $Domain
            RecordType = $RecordType
            Found      = $true
            Values     = $dnsResult | ForEach-Object { 
                switch ($RecordType) {
                    "MX"   { "$($_.Preference) $($_.NameExchange)" }
                    "TXT"  { $_.Strings }
                    "CNAME"{ $_.NameHost }
                    "A"    { $_.IPAddress }
                }
            }
        }
    }
    catch {
        @{
            Domain     = $Domain
            RecordType = $RecordType
            Found      = $false
            Error      = $_.Exception.Message
        }
    }
}

# === MX Routing ===
function Test-EmailRouting {
    param([string]$Domain)
    Write-LogMessage "Testing email routing configuration for $Domain"
    $results = @()
    
    $mxTest = Test-DNSRecord -Domain $Domain -RecordType "MX"
    $mxValuesCleaned = if ($mxTest.Found) {
        $mxTest.Values | ForEach-Object { ($_ -replace '^\d+\s+', '').TrimEnd('.').ToLower() }
    }
    
    $results += [PSCustomObject]@{
        Check          = "MX Records"
        Status         = if ($mxTest.Found) { "PASS" } else { "FAIL" }
        Details        = if ($mxTest.Found) { $mxTest.Values -join ", " } else { "No MX records found" }
        Recommendation = if (-not $mxTest.Found) { "Configure MX record pointing to smtp.google.com" } else { "" }
        Priority       = if (-not $mxTest.Found) { "HIGH" } else { "NONE" }
    }
    
    $hasGoogleMX = $mxValuesCleaned -contains "smtp.google.com"
    $results += [PSCustomObject]@{
        Check          = "Google MX Configuration"
        Status         = if ($hasGoogleMX) { "PASS" } else { "FAIL" }
        Details        = if ($hasGoogleMX) { "smtp.google.com MX record detected" } else { "smtp.google.com MX record not found" }
        Recommendation = if (-not $hasGoogleMX) { "Update MX records to include smtp.google.com" } else { "" }
        Priority       = if (-not $hasGoogleMX) { "HIGH" } else { "NONE" }
    }
    return $results
}

# === SPF ===
function Test-SPFRecord {
    param([string]$Domain)
    Write-LogMessage "Checking SPF record for $Domain"
    $spfTest = Test-DNSRecord -Domain $Domain -RecordType "TXT"
    $spfRecord = $null
    
    if ($spfTest.Found) {
        $spfRecord = $spfTest.Values | ForEach-Object { ($_ -replace '"','') -join "" } | Where-Object { $_ -match "^v=spf1" } | Select-Object -First 1
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

# === DKIM ===
function Test-DKIMRecord {
    param([string]$Domain)
    Write-LogMessage "Checking DKIM record for $Domain"
    $dkimDomain = "google._domainkey.$Domain"
    $dkimTest   = Test-DNSRecord -Domain $dkimDomain -RecordType "TXT"
    
    [PSCustomObject]@{
        Check          = "DKIM Record"
        Status         = if ($dkimTest.Found) { "PASS" } else { "FAIL" }
        Details        = if ($dkimTest.Found) { "DKIM record found for google selector" } else { "No DKIM record found at $dkimDomain" }
        Recommendation = if (-not $dkimTest.Found) { "Configure DKIM signing in Google Workspace Admin Console (selector: google)" } else { "" }
        Priority       = if (-not $dkimTest.Found) { "MEDIUM" } else { "NONE" }
    }
}

# === DMARC ===
function Test-DMARCRecord {
    param([string]$Domain)
    Write-LogMessage "Checking DMARC record for $Domain"
    $dmarcDomain = "_dmarc.$Domain"
    $dmarcTest   = Test-DNSRecord -Domain $dmarcDomain -RecordType "TXT"
    $dmarcRecord = $null
    
    if ($dmarcTest.Found) {
        $dmarcRecord = $dmarcTest.Values | ForEach-Object { ($_ -replace '"','') -join "" } | Where-Object { $_ -match "^v=DMARC1" } | Select-Object -First 1
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

# === Site Verification ===
function Test-GoogleSiteVerification {
    param([string]$Domain)
    Write-LogMessage "Checking Google Site Verification for $Domain"
    $txtTest = Test-DNSRecord -Domain $Domain -RecordType "TXT"
    $googleVerification = $false
    if ($txtTest.Found) {
        $googleVerification = $txtTest.Values | ForEach-Object { $_ -replace '"','' } | Where-Object { $_ -like "google-site-verification=*" } | Select-Object -First 1
    }
    
    [PSCustomObject]@{
        Check          = "Google Site Verification"
        Status         = if ($googleVerification) { "PASS" } else { "WARNING" }
        Details        = if ($googleVerification) { "Google site verification found" } else { "No Google site verification TXT record found" }
        Recommendation = if (-not $googleVerification) { "Add Google site verification TXT record if required" } else { "" }
        Priority       = if (-not $googleVerification) { "LOW" } else { "NONE" }
    }
}

# === Subdomains & CAA (unchanged) ===
function Test-SubdomainDelegation {
    param([string]$Domain)
    $results = @()
    foreach ($subdomain in @("mail","calendar","drive","docs","sites")) {
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
    } catch {
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

# === Report Generation (unchanged style) ===
function Generate-HTMLReport {
    param([array]$Results, [string]$Domain, [string]$OutputFile)
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $html = "<html><body><h2>Google Workspace Health Report - $Domain</h2><p>Generated $timestamp</p><table border=1 cellspacing=0 cellpadding=5>"
    $html += "<tr><th>Check</th><th>Status</th><th>Details</th><th>Recommendation</th></tr>"
    foreach ($r in $Results) {
        $html += "<tr><td>$($r.Check)</td><td>$($r.Status)</td><td>$($r.Details)</td><td>$($r.Recommendation)</td></tr>"
    }
    $html += "</table></body></html>"
    $html | Out-File -FilePath $OutputFile -Encoding UTF8
    Write-LogMessage "HTML report generated: $OutputFile" -Level "SUCCESS"
}

# === Health Check Runner ===
function Start-GWSHealthCheck {
    param([string]$Domain, [string]$OutputPath, [bool]$DetailedReport)
    $results = @()
    $results += Test-EmailRouting -Domain $Domain
    $results += Test-SPFRecord -Domain $Domain
    $results += Test-DKIMRecord -Domain $Domain
    $results += Test-DMARCRecord -Domain $Domain
    $results += Test-GoogleSiteVerification -Domain $Domain
    if ($DetailedReport) {
        $results += Test-SubdomainDelegation -Domain $Domain
        $results += Test-SecurityHeaders -Domain $Domain
    }
    Generate-HTMLReport -Results $results -Domain $Domain -OutputFile $ReportFile
    return $results
}

# === Main ===
Write-Host "GWSHealth Check is starting ...."
Write-Host ""
Start-Sleep 5

$Domain = Read-Host -Prompt "Please enter domain name...."

try {
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    $results = Start-GWSHealthCheck -Domain $Domain -OutputPath $OutputPath -DetailedReport $Detailed.IsPresent
} catch {
    Write-LogMessage "Error: $($_.Exception.Message)" -Level "ERROR"
}

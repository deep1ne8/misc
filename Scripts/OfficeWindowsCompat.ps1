#Requires -Version 5.1
<#
.SYNOPSIS
    Office and Windows Compatibility Verification Tool
.DESCRIPTION
    Checks installed Office and Windows versions for compatibility and support status.
    Optimized for speed and accuracy with comprehensive error handling.
.NOTES
    Version: 2.0
    Supports: Office 2010-2024, Microsoft 365, Windows 10/11
#>

[CmdletBinding()]
param()

#region Functions
function Get-WindowsVersionInfo {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $build = [int]$os.BuildNumber
        
        # Determine Windows version from build number
        $winName = switch ($build) {
            {$_ -ge 26100} { "Windows 11 24H2+" }
            {$_ -ge 22631} { "Windows 11 23H2" }
            {$_ -ge 22621} { "Windows 11 22H2" }
            {$_ -ge 22000} { "Windows 11" }
            {$_ -ge 19045} { "Windows 10 22H2" }
            {$_ -ge 19044} { "Windows 10 21H2" }
            {$_ -ge 19041} { "Windows 10 20H1+" }
            {$_ -ge 18363} { "Windows 10 1909" }
            default { $os.Caption }
        }
        
        return @{
            Name    = $winName
            Version = $os.Version
            Build   = $build
            Arch    = $os.OSArchitecture
        }
    }
    catch {
        Write-Error "Failed to retrieve Windows version: $_"
        return $null
    }
}

function Get-OfficeInstallation {
    $installations = @()
    
    # Check Click-to-Run (C2R) - Modern Office/Microsoft 365
    $c2rPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $c2rPath) {
        try {
            $c2rConfig = Get-ItemProperty -Path $c2rPath -ErrorAction Stop
            if ($c2rConfig.VersionToReport) {
                $version = $c2rConfig.VersionToReport
                $majorVer = $version.Split('.')[0]
                
                $productName = if ($c2rConfig.ProductReleaseIds) {
                    switch -Regex ($c2rConfig.ProductReleaseIds) {
                        "O365" { "Microsoft 365" }
                        "2024" { "Office LTSC 2024" }
                        "2021" { "Office LTSC 2021" }
                        "2019" { "Office 2019" }
                        "2016" { "Office 2016" }
                        default { "Office (Click-to-Run)" }
                    }
                } else {
                    "Office (Click-to-Run)"
                }
                
                $installations += @{
                    Name       = $productName
                    Version    = $version
                    MajorVer   = $majorVer
                    Type       = "Click-to-Run"
                    Platform   = $c2rConfig.Platform
                    UpdatePath = $c2rConfig.UpdateChannel
                }
            }
        }
        catch {
            Write-Verbose "Error reading C2R configuration: $_"
        }
    }
    
    # Check MSI-based installations (Legacy)
    $officeRegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
    )
    
    foreach ($basePath in $officeRegPaths) {
        if (Test-Path $basePath) {
            $versionKeys = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -match '^\d{2}\.\d$' }
            
            foreach ($key in $versionKeys) {
                $version = $key.PSChildName
                $commonPath = Join-Path $key.PSPath "Common\InstallRoot"
                
                if (Test-Path $commonPath) {
                    $installInfo = Get-ItemProperty -Path $commonPath -ErrorAction SilentlyContinue
                    if ($installInfo.Path) {
                        $productName = switch ($version) {
                            "16.0" { "Office 2016 (MSI)" }
                            "15.0" { "Office 2013" }
                            "14.0" { "Office 2010" }
                            "12.0" { "Office 2007" }
                            default { "Office $version" }
                        }
                        
                        # Avoid duplicates
                        if ($installations.MajorVer -notcontains $version.Split('.')[0]) {
                            $installations += @{
                                Name     = $productName
                                Version  = $version
                                MajorVer = $version.Split('.')[0]
                                Type     = "MSI"
                                Platform = if ($basePath -match "WOW6432Node") { "x86" } else { "x64" }
                            }
                        }
                    }
                }
            }
        }
    }
    
    return $installations
}

function Test-Compatibility {
    param(
        [int]$WindowsBuild,
        [string]$OfficeMajorVersion,
        [string]$OfficeType
    )
    
    $result = @{
        Compatible = $false
        Status     = "Unknown"
        Message    = ""
        Notes      = @()
    }
    
    # Windows 11 (Build 22000+)
    if ($WindowsBuild -ge 22000) {
        switch ($OfficeMajorVersion) {
            "16" {
                $result.Compatible = $true
                $result.Status = "Supported"
                $result.Message = "Fully supported on Windows 11"
                if ($OfficeType -eq "MSI") {
                    $result.Notes += "Consider migrating to Click-to-Run for better update management"
                }
            }
            "15" {
                $result.Status = "Not Supported"
                $result.Message = "Office 2013 is not officially supported on Windows 11"
                $result.Notes += "May experience activation, update, or stability issues"
                $result.Notes += "End of Extended Support: April 11, 2023"
            }
            "14" {
                $result.Status = "Not Supported"
                $result.Message = "Office 2010 is incompatible with Windows 11"
                $result.Notes += "End of Extended Support: October 13, 2020"
                $result.Notes += "Security risk - immediate upgrade required"
            }
            default {
                $result.Status = "Unknown"
                $result.Message = "Office version not recognized or too old for Windows 11"
            }
        }
    }
    # Windows 10 (Build 19041-21999)
    elseif ($WindowsBuild -ge 19041) {
        switch ($OfficeMajorVersion) {
            "16" {
                $result.Compatible = $true
                $result.Status = "Supported"
                $result.Message = "Fully supported on Windows 10"
            }
            "15" {
                $result.Compatible = $true
                $result.Status = "Limited Support"
                $result.Message = "Office 2013 basic functionality works on Windows 10"
                $result.Notes += "End of Extended Support: April 11, 2023"
                $result.Notes += "No security updates - upgrade recommended"
            }
            "14" {
                $result.Status = "Not Supported"
                $result.Message = "Office 2010 is out of support"
                $result.Notes += "End of Extended Support: October 13, 2020"
                $result.Notes += "Security vulnerability - upgrade required"
            }
            default {
                $result.Status = "Not Supported"
                $result.Message = "Office version is too old for modern Windows"
            }
        }
    }
    # Older Windows versions
    else {
        $result.Status = "Not Supported"
        $result.Message = "Windows build is outdated - upgrade to Windows 10 22H2 or Windows 11"
        $result.Notes += "Both Windows and Office may be out of security support"
    }
    
    return $result
}
#endregion

#region Main Script
Write-Host "`n╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Office & Windows Compatibility Check Tool v2.0     ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Get Windows Information
Write-Host "[1/2] Detecting Windows version..." -ForegroundColor Yellow
$windows = Get-WindowsVersionInfo

if (-not $windows) {
    Write-Host "`n❌ Failed to detect Windows version. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host "      ✓ Windows: $($windows.Name)" -ForegroundColor Green
Write-Host "      ✓ Build: $($windows.Build) | Architecture: $($windows.Arch)`n" -ForegroundColor Gray

# Get Office Installation
Write-Host "[2/2] Detecting Office installation..." -ForegroundColor Yellow
$officeInstalls = Get-OfficeInstallation

if ($officeInstalls.Count -eq 0) {
    Write-Host "`n❌ No Office installation detected on this system.`n" -ForegroundColor Red
    Write-Host "If Office is installed, try running as Administrator.`n" -ForegroundColor Yellow
    exit 1
}

# Display Office Information
foreach ($office in $officeInstalls) {
    Write-Host "      ✓ Office: $($office.Name)" -ForegroundColor Green
    Write-Host "      ✓ Version: $($office.Version) | Type: $($office.Type) | Platform: $($office.Platform)" -ForegroundColor Gray
}

Write-Host "`n" + ("─" * 60) -ForegroundColor DarkGray

# Compatibility Analysis
Write-Host "`n📊 COMPATIBILITY ANALYSIS" -ForegroundColor Cyan
Write-Host ("─" * 60) -ForegroundColor DarkGray

foreach ($office in $officeInstalls) {
    $compatibility = Test-Compatibility -WindowsBuild $windows.Build `
                                        -OfficeMajorVersion $office.MajorVer `
                                        -OfficeType $office.Type
    
    Write-Host "`n📦 $($office.Name)" -ForegroundColor White
    
    if ($compatibility.Compatible) {
        Write-Host "   Status: ✅ $($compatibility.Status)" -ForegroundColor Green
    } elseif ($compatibility.Status -eq "Limited Support") {
        Write-Host "   Status: ⚠️  $($compatibility.Status)" -ForegroundColor Yellow
    } else {
        Write-Host "   Status: ❌ $($compatibility.Status)" -ForegroundColor Red
    }
    
    Write-Host "   Details: $($compatibility.Message)" -ForegroundColor Gray
    
    if ($compatibility.Notes.Count -gt 0) {
        Write-Host "   Notes:" -ForegroundColor Yellow
        foreach ($note in $compatibility.Notes) {
            Write-Host "      • $note" -ForegroundColor Yellow
        }
    }
}

# Recommendations
Write-Host "`n" + ("─" * 60) -ForegroundColor DarkGray
Write-Host "💡 BEST PRACTICES & RECOMMENDATIONS" -ForegroundColor Cyan
Write-Host ("─" * 60) -ForegroundColor DarkGray

$recommendations = @(
    "Keep both Windows and Office within their support lifecycle",
    "Use Microsoft 365 or Office LTSC 2021/2024 with Windows 10 22H2 or Windows 11",
    "Click-to-Run installations provide better security and update management",
    "Avoid running out-of-support Office versions (2010/2013) on Windows 11",
    "Check support status: https://learn.microsoft.com/lifecycle"
)

foreach ($rec in $recommendations) {
    Write-Host "   • $rec" -ForegroundColor White
}

Write-Host "`n" + ("═" * 60) -ForegroundColor DarkGray
Write-Host "Scan completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray
Write-Host ("═" * 60) + "`n" -ForegroundColor DarkGray

#endregion

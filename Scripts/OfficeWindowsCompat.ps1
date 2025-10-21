#Requires -Version 5.1
<#
.SYNOPSIS
    Office, Teams, and Windows Compatibility Verification Tool
.DESCRIPTION
    Checks installed Office, Microsoft Teams, and Windows versions for compatibility and support status.
    Optimized for speed and accuracy with comprehensive error handling.
    Detects both Classic Teams and New Teams installations.
.NOTES
    Version: 2.1
    Supports: Office 2010-2024, Microsoft 365, Classic Teams, New Teams, Windows 10/11
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

function Get-TeamsInstallation {
    $installations = @()
    
    # Check New Teams (Teams 2.x) - Modern WebView2-based client
    $newTeamsPaths = @(
        "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe",
        "$env:ProgramFiles\WindowsApps\MSTeams_*\ms-teams.exe",
        "HKLM:\SOFTWARE\Microsoft\Teams",
        "HKCU:\SOFTWARE\Microsoft\Office\Teams"
    )
    
    # New Teams detection via registry (most reliable)
    $newTeamsReg = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\MSTeams"
    if (Test-Path $newTeamsReg) {
        try {
            $teamsInfo = Get-ItemProperty -Path $newTeamsReg -ErrorAction Stop
            $installations += @{
                Name        = "Microsoft Teams (New)"
                Version     = $teamsInfo.DisplayVersion
                Type        = "New Teams"
                MajorVer    = "2"
                InstallPath = $teamsInfo.InstallLocation
                Scope       = "Machine-wide"
            }
        }
        catch {
            Write-Verbose "Error reading New Teams registry: $_"
        }
    }
    
    # Check for New Teams user installation
    $newTeamsUserPath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe"
    if ((Test-Path $newTeamsUserPath) -and ($installations.Type -notcontains "New Teams")) {
        try {
            $fileVersion = (Get-Item $newTeamsUserPath -ErrorAction Stop).VersionInfo.FileVersion
            $installations += @{
                Name        = "Microsoft Teams (New)"
                Version     = $fileVersion
                Type        = "New Teams"
                MajorVer    = "2"
                InstallPath = $newTeamsUserPath
                Scope       = "Per-user"
            }
        }
        catch {
            Write-Verbose "Error reading New Teams executable: $_"
        }
    }
    
    # Check Classic Teams (Teams 1.x) - Electron-based legacy client
    $classicTeamsPaths = @(
        "$env:LOCALAPPDATA\Microsoft\Teams\current\Teams.exe",
        "$env:ProgramFiles\Microsoft\Teams\current\Teams.exe",
        "${env:ProgramFiles(x86)}\Microsoft\Teams\current\Teams.exe"
    )
    
    foreach ($path in $classicTeamsPaths) {
        if (Test-Path $path) {
            try {
                $fileInfo = Get-Item $path -ErrorAction Stop
                $version = $fileInfo.VersionInfo.FileVersion
                
                # Avoid duplicates
                if ($installations.Type -notcontains "Classic Teams") {
                    $scope = if ($path -match "ProgramFiles") { "Machine-wide" } else { "Per-user" }
                    
                    $installations += @{
                        Name        = "Microsoft Teams (Classic)"
                        Version     = $version
                        Type        = "Classic Teams"
                        MajorVer    = "1"
                        InstallPath = $path
                        Scope       = $scope
                    }
                    break
                }
            }
            catch {
                Write-Verbose "Error reading Classic Teams at $path : $_"
            }
        }
    }
    
    # Check Teams Machine-Wide Installer registry
    $teamsMWI = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Teams Machine-Wide Installer"
    if ((Test-Path $teamsMWI) -and ($installations.Count -eq 0)) {
        try {
            $mwiInfo = Get-ItemProperty -Path $teamsMWI -ErrorAction Stop
            $installations += @{
                Name        = "Microsoft Teams (Classic)"
                Version     = $mwiInfo.DisplayVersion
                Type        = "Classic Teams"
                MajorVer    = "1"
                InstallPath = "Machine-Wide Installer"
                Scope       = "Machine-wide"
            }
        }
        catch {
            Write-Verbose "Error reading Teams MWI registry: $_"
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

function Test-TeamsCompatibility {
    param(
        [int]$WindowsBuild,
        [string]$TeamsMajorVersion,
        [string]$TeamsType
    )
    
    $result = @{
        Compatible = $false
        Status     = "Unknown"
        Message    = ""
        Notes      = @()
    }
    
    # Windows 11 (Build 22000+)
    if ($WindowsBuild -ge 22000) {
        if ($TeamsMajorVersion -eq "2" -or $TeamsType -eq "New Teams") {
            $result.Compatible = $true
            $result.Status = "Supported"
            $result.Message = "New Teams is optimized for Windows 11"
            $result.Notes += "Recommended version for Windows 11"
        }
        elseif ($TeamsMajorVersion -eq "1" -or $TeamsType -eq "Classic Teams") {
            $result.Compatible = $true
            $result.Status = "Limited Support"
            $result.Message = "Classic Teams works but is deprecated"
            $result.Notes += "Classic Teams retired as of March 31, 2024"
            $result.Notes += "Upgrade to New Teams for continued support and performance"
        }
    }
    # Windows 10 (Build 19041+)
    elseif ($WindowsBuild -ge 19041) {
        if ($TeamsMajorVersion -eq "2" -or $TeamsType -eq "New Teams") {
            $result.Compatible = $true
            $result.Status = "Supported"
            $result.Message = "New Teams is fully supported on Windows 10"
            $result.Notes += "Requires Windows 10 version 10.0.19041 or higher"
        }
        elseif ($TeamsMajorVersion -eq "1" -or $TeamsType -eq "Classic Teams") {
            $result.Compatible = $true
            $result.Status = "Limited Support"
            $result.Message = "Classic Teams works but support has ended"
            $result.Notes += "Classic Teams retired as of March 31, 2024"
            $result.Notes += "Migrate to New Teams immediately"
        }
    }
    # Windows 10 older builds (18362-19040)
    elseif ($WindowsBuild -ge 18362) {
        if ($TeamsMajorVersion -eq "2" -or $TeamsType -eq "New Teams") {
            $result.Status = "Not Supported"
            $result.Message = "New Teams requires Windows 10 build 19041 or higher"
            $result.Notes += "Update Windows 10 to version 2004 (19041) or later"
        }
        elseif ($TeamsMajorVersion -eq "1" -or $TeamsType -eq "Classic Teams") {
            $result.Status = "Not Supported"
            $result.Message = "Classic Teams is retired and unsupported"
            $result.Notes += "Update Windows 10 to build 19041+ and install New Teams"
        }
    }
    # Older Windows versions
    else {
        $result.Status = "Not Supported"
        $result.Message = "Windows build is too old for modern Teams"
        $result.Notes += "Upgrade to Windows 10 version 2004+ or Windows 11"
    }
    
    return $result
}
#endregion

#region Main Script
Write-Host "`n╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║  Office, Teams & Windows Compatibility Check v2.1   ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

# Get Windows Information
Write-Host "[1/3] Detecting Windows version..." -ForegroundColor Yellow
$windows = Get-WindowsVersionInfo

if (-not $windows) {
    Write-Host "`n❌ Failed to detect Windows version. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host "      ✓ Windows: $($windows.Name)" -ForegroundColor Green
Write-Host "      ✓ Build: $($windows.Build) | Architecture: $($windows.Arch)`n" -ForegroundColor Gray

# Get Office Installation
Write-Host "[2/3] Detecting Office installation..." -ForegroundColor Yellow
$officeInstalls = Get-OfficeInstallation

if ($officeInstalls.Count -eq 0) {
    Write-Host "      ⚠️  No Office installation detected" -ForegroundColor Yellow
} else {
    # Display Office Information
    foreach ($office in $officeInstalls) {
        Write-Host "      ✓ Office: $($office.Name)" -ForegroundColor Green
        Write-Host "      ✓ Version: $($office.Version) | Type: $($office.Type) | Platform: $($office.Platform)" -ForegroundColor Gray
    }
}

# Get Teams Installation
Write-Host "`n[3/3] Detecting Microsoft Teams installation..." -ForegroundColor Yellow
$teamsInstalls = Get-TeamsInstallation

if ($teamsInstalls.Count -eq 0) {
    Write-Host "      ⚠️  No Microsoft Teams installation detected" -ForegroundColor Yellow
} else {
    # Display Teams Information
    foreach ($teams in $teamsInstalls) {
        Write-Host "      ✓ Teams: $($teams.Name)" -ForegroundColor Green
        Write-Host "      ✓ Version: $($teams.Version) | Scope: $($teams.Scope)" -ForegroundColor Gray
    }
}

Write-Host "`n" + ("─" * 60) -ForegroundColor DarkGray

# Compatibility Analysis
Write-Host "`n📊 COMPATIBILITY ANALYSIS" -ForegroundColor Cyan
Write-Host ("─" * 60) -ForegroundColor DarkGray

# Office Compatibility
if ($officeInstalls.Count -gt 0) {
    Write-Host "`n🏢 MICROSOFT OFFICE" -ForegroundColor White
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
}

# Teams Compatibility
if ($teamsInstalls.Count -gt 0) {
    Write-Host "`n💬 MICROSOFT TEAMS" -ForegroundColor White
    Write-Host ("─" * 60) -ForegroundColor DarkGray
    
    foreach ($teams in $teamsInstalls) {
        $compatibility = Test-TeamsCompatibility -WindowsBuild $windows.Build `
                                                  -TeamsMajorVersion $teams.MajorVer `
                                                  -TeamsType $teams.Type
        
        Write-Host "`n📦 $($teams.Name)" -ForegroundColor White
        
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
}

# Recommendations
Write-Host "`n" + ("─" * 60) -ForegroundColor DarkGray
Write-Host "💡 BEST PRACTICES & RECOMMENDATIONS" -ForegroundColor Cyan
Write-Host ("─" * 60) -ForegroundColor DarkGray

$recommendations = @(
    "Keep Windows, Office, and Teams within their support lifecycle",
    "Use Microsoft 365 or Office LTSC 2021/2024 with Windows 10 22H2 or Windows 11",
    "Migrate from Classic Teams to New Teams (Classic retired March 31, 2024)",
    "New Teams requires Windows 10 build 19041 (version 2004) or later",
    "Click-to-Run Office installations provide better security and update management",
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

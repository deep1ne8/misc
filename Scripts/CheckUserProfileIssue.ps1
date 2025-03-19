# Windows 11 Profile Troubleshooting Script
# This script helps diagnose user profile issues, particularly with Windows 11 23H2
# Run with administrator privileges for complete results

# Create output directory and log file
$outputDir = "$env:USERPROFILE\Desktop\ProfileDiagnostics"
$logFile = "$outputDir\ProfileDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Initialize
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    "$timestamp - $message" | Out-File -Append -FilePath $logFile
    Write-Host $message
}

function Initialize-Diagnostics {
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir | Out-Null
    }
    
    Write-Log "=== Windows 11 Profile Troubleshooting Script ==="
    Write-Log "Script started at $(Get-Date)"
    Write-Log "Computer Name: $env:COMPUTERNAME"
    Write-Log "Current User: $env:USERNAME"
    
    # OS Version Info
    $osInfo = Get-CimInstance Win32_OperatingSystem
    Write-Log "OS: $($osInfo.Caption)"
    Write-Log "Version: $($osInfo.Version)"
    Write-Log "Build: $($osInfo.BuildNumber)"
    
    # Check for 23H2 specifically
    $currentBuild = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "CurrentBuild").CurrentBuild
    $ubr = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "UBR").UBR
    $displayVersion = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion" -ErrorAction SilentlyContinue).DisplayVersion
    
    Write-Log "Current Build: $currentBuild.$ubr"
    if ($displayVersion) {
        Write-Log "Display Version: $displayVersion"
        if ($displayVersion -eq "23H2") {
            Write-Log "CONFIRMED: System is running Windows 11 23H2"
        }
    }
}

# System & Profile Checks
function Test-UserProfileService {
    Write-Log "Checking User Profile Service status..."
    $service = Get-Service -Name "ProfSvc" -ErrorAction SilentlyContinue
    if ($service) {
        Write-Log "User Profile Service Status: $($service.Status)"
        if ($service.Status -ne "Running") {
            Write-Log "WARNING: User Profile Service is not running!" 
            
            # Try to start the service
            try {
                Start-Service -Name "ProfSvc" -ErrorAction Stop
                Write-Log "Attempted to start User Profile Service."
                $service = Get-Service -Name "ProfSvc"
                Write-Log "New status: $($service.Status)"
            }
            catch {
                Write-Log "Failed to start User Profile Service: $_"
            }
        }
    }
    else {
        Write-Log "ERROR: Could not retrieve User Profile Service information!"
    }
}

function Get-ProfileList {
    Write-Log "Gathering user profile information..."
    
    # Get profiles from registry
    $profileList = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -ErrorAction SilentlyContinue | 
                   Where-Object { $_.PSChildName -match 'S-1-5-21' } |
                   Select-Object @{Name="SID";Expression={$_.PSChildName}}, 
                                @{Name="ProfilePath";Expression={$_.ProfileImagePath}} 

    Write-Log "Found $($profileList.Count) user profiles on this system."
    
    # Export profile list
    $profileList | Export-Csv -Path "$outputDir\ProfileList.csv" -NoTypeInformation
    
    return $profileList
}

function Test-ProfileRegistry {
    Write-Log "Checking profile registry for corruption..."
    
    try {
        $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        $profileKeys = Get-ChildItem -Path $profileListPath -ErrorAction Stop
        
        foreach ($key in $profileKeys) {
            $sid = Split-Path -Path $key.PSPath -Leaf
            
            if ($sid -notmatch 'S-1-5-21') { continue }
            
            $profilePath = (Get-ItemProperty -Path $key.PSPath -Name "ProfileImagePath" -ErrorAction SilentlyContinue).ProfileImagePath
            
            if ($sid.EndsWith(".bak")) {
                Write-Log "WARNING: Found backup profile registry key: $sid"
            }
            
            if ($profilePath -and -not (Test-Path $profilePath)) {
                Write-Log "WARNING: Profile registry points to non-existent path: $profilePath for SID $sid"
            }
        }
    }
    catch {
        Write-Log "ERROR checking profile registry: $_"
    }
}

function Get-ProfileEvents {
    Write-Log "Analyzing Event Logs for profile-related issues..."
    
    $startTime = (Get-Date).AddDays(-7)
    
    $profileServiceEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'System'
        ProviderName = 'Microsoft-Windows-User Profiles Service'
        Level = 2,3 
        StartTime = $startTime
    } -ErrorAction SilentlyContinue

    if ($profileServiceEvents) {
        Write-Log "Found $($profileServiceEvents.Count) User Profile Service errors/warnings in the last 7 days."
        $profileServiceEvents | Select-Object TimeCreated, Id, Message | 
                               Export-Csv -Path "$outputDir\ProfileServiceEvents.csv" -NoTypeInformation

        $criticalEvents = $profileServiceEvents | Where-Object { $_.Id -in @(1511, 1500, 1508, 1515, 1530, 1534) }
        if ($criticalEvents) {
            Write-Log "WARNING: Found critical profile-related events:"
            foreach ($evt in $criticalEvents | Select-Object -First 5) {
                $messageSnippet = if ($evt.Message) { $evt.Message.Substring(0, [Math]::Min(100, $evt.Message.Length)) } else { "No message available" }
                Write-Log "  - Event ID $($evt.Id) at $($evt.TimeCreated): $messageSnippet..."
            }
        }
    }
    else {
        Write-Log "No User Profile Service errors found in the Event Log."
    }
}

Initialize-Diagnostics
Test-UserProfileService
Get-ProfileList
Test-ProfileRegistry
Get-ProfileEvents

Write-Log "Script completed."
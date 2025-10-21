#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Migrate Classic Teams to New Teams for all users on the system
.DESCRIPTION
    Comprehensive migration script that:
    - Detects existing Teams installations (Classic and New)
    - Stops running Teams processes
    - Uninstalls Classic Teams for all users
    - Cleans residual data and cache
    - Installs New Teams (MSIX) machine-wide
    - Verifies successful installation
.NOTES
    Version: 2.0
    Requires: PowerShell 5.1+, Administrator privileges
    Tested on: Windows 10 (19041+), Windows 11
#>

[CmdletBinding()]
param(
    [switch]$SkipBackup,
    [switch]$Force,
    [string]$LogPath = "$env:TEMP\TeamsMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Console output with color
    $color = switch ($Level) {
        'Info'    { 'White' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
    }
    Write-Host $logMessage -ForegroundColor $color
    
    # File output
    Add-Content -Path $LogPath -Value $logMessage -ErrorAction SilentlyContinue
}

function Test-Prerequisites {
    Write-Log "Checking prerequisites..." -Level Info
    
    # Check Windows version
    $build = [int](Get-CimInstance Win32_OperatingSystem).BuildNumber
    if ($build -lt 19041) {
        Write-Log "Windows 10 build 19041 (version 2004) or later required. Current build: $build" -Level Error
        return $false
    }
    Write-Log "Windows build: $build ✓" -Level Success
    
    # Check internet connectivity
    try {
        $null = Test-NetConnection -ComputerName "aka.ms" -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue -ErrorAction Stop
        Write-Log "Internet connectivity ✓" -Level Success
    }
    catch {
        Write-Log "No internet connection detected. Cannot download New Teams." -Level Error
        return $false
    }
    
    return $true
}

function Get-AllUserProfiles {
    $profiles = @()
    
    # Get all user profile paths from registry
    $profileList = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21' } |
        Select-Object -ExpandProperty ProfileImagePath
    
    foreach ($profilePath in $profileList) {
        if (Test-Path $profilePath) {
            $profiles += $profilePath
        }
    }
    
    Write-Log "Found $($profiles.Count) user profiles" -Level Info
    return $profiles
}

function Stop-TeamsProcesses {
    Write-Log "Stopping all Teams processes..." -Level Info
    
    $processes = @('Teams', 'ms-teams', 'ms-teamsupdate')
    $stopped = 0
    
    foreach ($processName in $processes) {
        $procs = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($procs) {
            $procs | Stop-Process -Force -ErrorAction SilentlyContinue
            $stopped += $procs.Count
            Write-Log "Stopped $($procs.Count) '$processName' process(es)" -Level Info
        }
    }
    
    if ($stopped -eq 0) {
        Write-Log "No Teams processes were running" -Level Info
    } else {
        Start-Sleep -Seconds 2
        Write-Log "All Teams processes stopped" -Level Success
    }
}

function Get-ClassicTeamsInstallation {
    $installations = @()
    
    # Method 1: Check per-user installations for all profiles
    $userProfiles = Get-AllUserProfiles
    foreach ($profile in $userProfiles) {
        $teamsPath = Join-Path $profile "AppData\Local\Microsoft\Teams"
        $updateExe = Join-Path $teamsPath "Update.exe"
        
        if (Test-Path $updateExe) {
            $installations += @{
                Type = "PerUser"
                Path = $teamsPath
                Uninstaller = $updateExe
                Profile = $profile
            }
        }
    }
    
    # Method 2: Check machine-wide installation
    $machineWidePaths = @(
        "$env:ProgramFiles\Microsoft\Teams",
        "${env:ProgramFiles(x86)}\Microsoft\Teams"
    )
    
    foreach ($path in $machineWidePaths) {
        if (Test-Path "$path\Update.exe") {
            $installations += @{
                Type = "MachineWide"
                Path = $path
                Uninstaller = "$path\Update.exe"
                Profile = "N/A"
            }
        }
    }
    
    # Method 3: Check Teams Machine-Wide Installer via Registry
    $uninstallKeys = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    foreach ($key in $uninstallKeys) {
        $apps = Get-ItemProperty $key -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*Teams Machine-Wide Installer*" }
        
        foreach ($app in $apps) {
            if ($app.UninstallString) {
                $installations += @{
                    Type = "MWI"
                    Path = $app.InstallLocation
                    Uninstaller = $app.UninstallString
                    DisplayName = $app.DisplayName
                    Profile = "Machine-Wide"
                }
            }
        }
    }
    
    return $installations
}

function Uninstall-ClassicTeams {
    Write-Log "Detecting Classic Teams installations..." -Level Info
    
    $installations = Get-ClassicTeamsInstallation
    
    if ($installations.Count -eq 0) {
        Write-Log "No Classic Teams installations found" -Level Info
        return $true
    }
    
    Write-Log "Found $($installations.Count) Classic Teams installation(s)" -Level Warning
    
    foreach ($install in $installations) {
        Write-Log "Uninstalling: $($install.Type) - $($install.Path)" -Level Info
        
        try {
            if ($install.Type -eq "MWI" -and $install.Uninstaller) {
                # Parse uninstall string
                if ($install.Uninstaller -match 'MsiExec.exe\s+(/X|/I)({[^}]+})') {
                    $productCode = $matches[2]
                    Write-Log "Uninstalling MWI with product code: $productCode" -Level Info
                    
                    $msiArgs = @(
                        "/x"
                        $productCode
                        "/qn"
                        "/norestart"
                        "/L*v"
                        "`"$env:TEMP\TeamsMWI_Uninstall.log`""
                    )
                    
                    $process = Start-Process "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -NoNewWindow
                    
                    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                        Write-Log "MWI uninstalled successfully (Exit code: $($process.ExitCode))" -Level Success
                    } else {
                        Write-Log "MWI uninstall completed with exit code: $($process.ExitCode)" -Level Warning
                    }
                }
            }
            elseif ($install.Uninstaller -and (Test-Path $install.Uninstaller)) {
                Write-Log "Running uninstaller: $($install.Uninstaller)" -Level Info
                
                # Use Update.exe to uninstall
                $uninstallArgs = @("--uninstall", "-s")
                $process = Start-Process -FilePath $install.Uninstaller -ArgumentList $uninstallArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
                
                if ($process.ExitCode -eq 0) {
                    Write-Log "Uninstalled successfully from: $($install.Path)" -Level Success
                } else {
                    Write-Log "Uninstall completed with exit code: $($process.ExitCode)" -Level Warning
                }
            }
            
            Start-Sleep -Milliseconds 500
        }
        catch {
            Write-Log "Error uninstalling from $($install.Path): $($_.Exception.Message)" -Level Error
        }
    }
    
    return $true
}

function Remove-TeamsCache {
    param([switch]$AllUsers)
    
    Write-Log "Cleaning Teams cache and residual data..." -Level Info
    
    $cachePaths = @(
        'AppData\Local\Microsoft\Teams',
        'AppData\Roaming\Microsoft\Teams',
        'AppData\Local\Microsoft\TeamsMeetingAddin',
        'AppData\Local\SquirrelTemp',
        'AppData\Local\Microsoft\TeamsPresenceAddin'
    )
    
    $profiles = if ($AllUsers) {
        Get-AllUserProfiles
    } else {
        @($env:USERPROFILE)
    }
    
    $removedCount = 0
    foreach ($profile in $profiles) {
        foreach ($cachePath in $cachePaths) {
            $fullPath = Join-Path $profile $cachePath
            
            if (Test-Path $fullPath) {
                try {
                    Remove-Item -Path $fullPath -Recurse -Force -ErrorAction Stop
                    $removedCount++
                    Write-Log "Removed: $fullPath" -Level Info
                }
                catch {
                    Write-Log "Could not remove: $fullPath - $($_.Exception.Message)" -Level Warning
                }
            }
        }
    }
    
    if ($removedCount -gt 0) {
        Write-Log "Cleaned $removedCount cache location(s)" -Level Success
    } else {
        Write-Log "No cache data found to clean" -Level Info
    }
}

function Get-NewTeamsInstallation {
    # Check for New Teams MSIX package
    $msixPackage = Get-AppxPackage -Name "MSTeams" -AllUsers -ErrorAction SilentlyContinue
    
    if ($msixPackage) {
        return @{
            Installed = $true
            Version = $msixPackage.Version
            InstallLocation = $msixPackage.InstallLocation
            PackageFullName = $msixPackage.PackageFullName
        }
    }
    
    # Check via registry
    $regPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\MSTeams"
    if (Test-Path $regPath) {
        $regInfo = Get-ItemProperty -Path $regPath
        return @{
            Installed = $true
            Version = $regInfo.DisplayVersion
            InstallLocation = $regInfo.InstallLocation
        }
    }
    
    return @{ Installed = $false }
}

function Install-NewTeams {
    Write-Log "Checking for existing New Teams installation..." -Level Info
    
    $existingInstall = Get-NewTeamsInstallation
    if ($existingInstall.Installed) {
        Write-Log "New Teams is already installed (Version: $($existingInstall.Version))" -Level Warning
        
        if (-not $Force) {
            $response = Read-Host "New Teams already exists. Reinstall? (Y/N)"
            if ($response -ne 'Y') {
                Write-Log "Installation skipped by user" -Level Info
                return $false
            }
        }
    }
    
    Write-Log "Downloading New Teams installer..." -Level Info
    
    # Use the official bootstrapper URL
    $installerUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    $installerPath = "$env:TEMP\teamsbootstrapper.exe"
    
    try {
        # Download with progress
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing -ErrorAction Stop
        $ProgressPreference = 'Continue'
        
        if (-not (Test-Path $installerPath)) {
            throw "Installer download failed - file not found"
        }
        
        $fileSize = (Get-Item $installerPath).Length / 1MB
        Write-Log "Downloaded installer: $([math]::Round($fileSize, 2)) MB" -Level Success
    }
    catch {
        Write-Log "Failed to download installer: $($_.Exception.Message)" -Level Error
        return $false
    }
    
    Write-Log "Installing New Teams (this may take several minutes)..." -Level Info
    
    try {
        # Install with machine-wide provisioning
        $installArgs = @(
            "-p"  # Provision for all users
            "-o", "`"$env:TEMP\TeamsInstall.log`""
        )
        
        $process = Start-Process -FilePath $installerPath -ArgumentList $installArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
        
        if ($process.ExitCode -eq 0) {
            Write-Log "New Teams installed successfully" -Level Success
            
            # Clean up installer
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
            
            return $true
        } else {
            Write-Log "Installation completed with exit code: $($process.ExitCode)" -Level Warning
            Write-Log "Check log at: $env:TEMP\TeamsInstall.log" -Level Info
            return $false
        }
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Test-NewTeamsInstallation {
    Write-Log "Verifying New Teams installation..." -Level Info
    
    Start-Sleep -Seconds 3
    
    $newTeams = Get-NewTeamsInstallation
    
    if ($newTeams.Installed) {
        Write-Log "✅ Verification successful" -Level Success
        Write-Log "   Version: $($newTeams.Version)" -Level Info
        Write-Log "   Location: $($newTeams.InstallLocation)" -Level Info
        
        # Check if executable exists
        $exePath = "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe"
        if (Test-Path $exePath) {
            Write-Log "   Executable: $exePath" -Level Info
        }
        
        return $true
    } else {
        Write-Log "⚠️ New Teams installation could not be verified" -Level Warning
        Write-Log "Try checking: Get-AppxPackage -Name MSTeams" -Level Info
        return $false
    }
}
#endregion

#region Main Script
Write-Host "`n╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║     Classic Teams → New Teams Migration Tool        ║" -ForegroundColor Cyan
Write-Host "║                   Version 2.0                        ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

Write-Log "Migration started" -Level Info
Write-Log "Log file: $LogPath" -Level Info

# Step 1: Prerequisites
Write-Host "`n[1/6] Checking prerequisites..." -ForegroundColor Yellow
if (-not (Test-Prerequisites)) {
    Write-Log "Prerequisites check failed. Exiting." -Level Error
    exit 1
}

# Step 2: Stop Teams
Write-Host "`n[2/6] Stopping Teams processes..." -ForegroundColor Yellow
Stop-TeamsProcesses

# Step 3: Uninstall Classic Teams
Write-Host "`n[3/6] Uninstalling Classic Teams..." -ForegroundColor Yellow
if (-not (Uninstall-ClassicTeams)) {
    Write-Log "Failed to uninstall Classic Teams" -Level Error
    if (-not $Force) {
        exit 1
    }
}

# Step 4: Clean cache
Write-Host "`n[4/6] Cleaning Teams cache..." -ForegroundColor Yellow
Remove-TeamsCache -AllUsers

# Step 5: Install New Teams
Write-Host "`n[5/6] Installing New Teams..." -ForegroundColor Yellow
$installSuccess = Install-NewTeams

if (-not $installSuccess) {
    Write-Log "Installation failed or was skipped" -Level Warning
    exit 1
}

# Step 6: Verify installation
Write-Host "`n[6/6] Verifying installation..." -ForegroundColor Yellow
$verifySuccess = Test-NewTeamsInstallation

Write-Host "`n" + ("═" * 60) -ForegroundColor DarkGray
Write-Host "🎯 MIGRATION SUMMARY" -ForegroundColor Cyan
Write-Host ("═" * 60) -ForegroundColor DarkGray

if ($verifySuccess) {
    Write-Host "`n✅ Migration completed successfully!`n" -ForegroundColor Green
    
    Write-Host "Next Steps:" -ForegroundColor White
    Write-Host "   1. Users should sign in to New Teams from the Start menu" -ForegroundColor Gray
    Write-Host "   2. New Teams updates automatically via Windows Update" -ForegroundColor Gray
    Write-Host "   3. Verify with: Get-AppxPackage -Name MSTeams" -ForegroundColor Gray
    Write-Host "   4. Review log: $LogPath`n" -ForegroundColor Gray
} else {
    Write-Host "`n⚠️ Migration completed with warnings`n" -ForegroundColor Yellow
    Write-Host "Review log for details: $LogPath`n" -ForegroundColor Gray
}

Write-Host ("═" * 60) -ForegroundColor DarkGray
Write-Log "Migration process completed" -Level Info
#endregion

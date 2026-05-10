#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Migrate Classic Teams to New Teams for all users on the system
.DESCRIPTION
    Enterprise-grade migration script featuring:
    - Intelligent detection of all Teams installations
    - Safe uninstallation with rollback capability
    - Comprehensive cache cleanup for all users
    - New Teams (MSIX) installation with verification
    - Detailed logging and error handling
    - Backup and restore functionality
.PARAMETER Force
    Skip confirmation prompts for unattended deployment
.PARAMETER SkipBackup
    Skip backing up Teams settings and data
.PARAMETER LogPath
    Custom path for log file (default: %TEMP%\TeamsMigration_TIMESTAMP.log)
.PARAMETER InstallTimeout
    Maximum time in seconds to wait for installation (default: 600)
.EXAMPLE
    .\Migrate-ClassicTeams-to-NewTeams.ps1
    Interactive migration with prompts
.EXAMPLE
    .\Migrate-ClassicTeams-to-NewTeams.ps1 -Force
    Unattended migration for enterprise deployment
.NOTES
    Version: 2.1
    Requires: PowerShell 5.1+, Administrator privileges
    Tested on: Windows 10 (19041+), Windows 11
#>

[CmdletBinding()]
param(
    [switch]$SkipBackup,
    [switch]$Force,
    [string]$LogPath = "$env:TEMP\TeamsMigration_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    [int]$InstallTimeout = 600
)

# Script-level variables
$script:BackupPath = "$env:TEMP\TeamsBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$script:RollbackData = @()

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success', 'Verbose')]
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
        'Verbose' { 'Gray' }
    }
    
    if ($Level -eq 'Verbose' -and -not $VerbosePreference) {
        # Skip verbose messages unless -Verbose is specified
    } else {
        Write-Host $logMessage -ForegroundColor $color
    }
    
    # File output
    Add-Content -Path $LogPath -Value $logMessage -ErrorAction SilentlyContinue
}

function Test-Prerequisites {
    Write-Log "Checking prerequisites..." -Level Info
    
    # Check Windows version
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $build = [int]$os.BuildNumber
        
        if ($build -lt 19041) {
            Write-Log "Windows 10 build 19041 (version 2004) or later required. Current build: $build" -Level Error
            return $false
        }
        Write-Log "Windows build: $build ✓" -Level Success
        
        # Check available disk space (require at least 2GB free)
        $systemDrive = $env:SystemDrive
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop
        $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
        
        if ($freeSpaceGB -lt 2) {
            Write-Log "Insufficient disk space: ${freeSpaceGB}GB free. At least 2GB required." -Level Error
            return $false
        }
        Write-Log "Free disk space: ${freeSpaceGB}GB ✓" -Level Success
    }
    catch {
        Write-Log "Error checking system requirements: $($_.Exception.Message)" -Level Error
        return $false
    }
    
    # Check internet connectivity (faster method than Test-NetConnection)
    Write-Log "Testing internet connectivity..." -Level Verbose
    try {
        $testUrls = @(
            "https://go.microsoft.com",
            "https://aka.ms"
        )
        
        $connected = $false
        foreach ($url in $testUrls) {
            try {
                $null = Invoke-WebRequest -Uri $url -Method Head -TimeoutSec 5 -UseBasicParsing -ErrorAction Stop
                $connected = $true
                Write-Log "Internet connectivity verified via $url ✓" -Level Success
                break
            }
            catch {
                Write-Log "Could not reach $url, trying next..." -Level Verbose
            }
        }
        
        if (-not $connected) {
            Write-Log "No internet connection detected. Cannot download New Teams." -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Internet connectivity check failed: $($_.Exception.Message)" -Level Error
        return $false
    }
    
    return $true
}

function Get-AllUserProfiles {
    $profiles = @()
    
    # Get all user profile paths from registry
    $profileListPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
    
    try {
        $profileKeys = Get-ChildItem -Path $profileListPath -ErrorAction Stop |
            Where-Object { $_.PSChildName -match '^S-1-5-21-' }
        
        foreach ($key in $profileKeys) {
            $profilePath = (Get-ItemProperty -Path $key.PSPath -Name ProfileImagePath -ErrorAction SilentlyContinue).ProfileImagePath
            if ($profilePath -and (Test-Path $profilePath)) {
                $profiles += $profilePath
                Write-Log "Found profile: $profilePath" -Level Verbose
            }
        }
    }
    catch {
        Write-Log "Error enumerating user profiles: $($_.Exception.Message)" -Level Warning
    }
    
    Write-Log "Found $($profiles.Count) user profiles" -Level Info
    return $profiles
}

function Stop-TeamsProcesses {
    param([int]$MaxRetries = 3)
    
    Write-Log "Stopping all Teams processes..." -Level Info
    
    $processes = @('Teams', 'ms-teams', 'ms-teamsupdate', 'Update')
    $stopped = 0
    $attempt = 0
    
    while ($attempt -lt $MaxRetries) {
        $attempt++
        $runningProcesses = @()
        
        foreach ($processName in $processes) {
            $procs = Get-Process -Name $processName -ErrorAction SilentlyContinue
            if ($procs) {
                $runningProcesses += $procs
            }
        }
        
        if ($runningProcesses.Count -eq 0) {
            if ($attempt -eq 1) {
                Write-Log "No Teams processes were running" -Level Info
            } else {
                Write-Log "All Teams processes stopped successfully" -Level Success
            }
            return $true
        }
        
        Write-Log "Attempt $attempt/$MaxRetries - Stopping $($runningProcesses.Count) process(es)" -Level Info
        
        foreach ($proc in $runningProcesses) {
            try {
                $proc | Stop-Process -Force -ErrorAction Stop
                $stopped++
                Write-Log "Stopped: $($proc.Name) (PID: $($proc.Id))" -Level Verbose
            }
            catch {
                Write-Log "Could not stop $($proc.Name) (PID: $($proc.Id)): $($_.Exception.Message)" -Level Warning
            }
        }
        
        Start-Sleep -Seconds 3
    }
    
    # Final check
    $stillRunning = @()
    foreach ($processName in $processes) {
        $procs = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($procs) {
            $stillRunning += $procs
        }
    }
    
    if ($stillRunning.Count -gt 0) {
        Write-Log "Warning: $($stillRunning.Count) process(es) still running after $MaxRetries attempts" -Level Warning
        return $false
    }
    
    Write-Log "Stopped $stopped process(es) total" -Level Success
    return $true
}

function Get-ClassicTeamsInstallation {
    $installations = @()
    
    # Method 1: Check per-user installations for all profiles
    Write-Log "Scanning user profiles for Classic Teams..." -Level Verbose
    $userProfiles = Get-AllUserProfiles
    
    foreach ($profile in $userProfiles) {
        $teamsPath = Join-Path $profile "AppData\Local\Microsoft\Teams"
        $updateExe = Join-Path $teamsPath "Update.exe"
        $currentExe = Join-Path $teamsPath "current\Teams.exe"
        
        if (Test-Path $updateExe) {
            # Get version if available
            $version = "Unknown"
            if (Test-Path $currentExe) {
                try {
                    $version = (Get-Item $currentExe -ErrorAction Stop).VersionInfo.FileVersion
                }
                catch {
                    Write-Log "Could not read version from $currentExe" -Level Verbose
                }
            }
            
            $installations += @{
                Type = "PerUser"
                Path = $teamsPath
                Uninstaller = $updateExe
                Profile = $profile
                Version = $version
            }
            Write-Log "Found: Per-user Teams (v$version) at $teamsPath" -Level Verbose
        }
    }
    
    # Method 2: Check machine-wide installation
    Write-Log "Checking for machine-wide Classic Teams installations..." -Level Verbose
    $machineWidePaths = @(
        "$env:ProgramFiles\Microsoft\Teams",
        "${env:ProgramFiles(x86)}\Microsoft\Teams"
    )
    
    foreach ($path in $machineWidePaths) {
        $updateExe = "$path\Update.exe"
        if (Test-Path $updateExe) {
            $installations += @{
                Type = "MachineWide"
                Path = $path
                Uninstaller = $updateExe
                Profile = "Machine-Wide"
                Version = "Unknown"
            }
            Write-Log "Found: Machine-wide Teams at $path" -Level Verbose
        }
    }
    
    # Method 3: Check Teams Machine-Wide Installer via Registry
    Write-Log "Checking registry for Teams Machine-Wide Installer..." -Level Verbose
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
    )
    
    foreach ($path in $uninstallPaths) {
        try {
            $keys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            foreach ($key in $keys) {
                $app = Get-ItemProperty -Path $key.PSPath -ErrorAction SilentlyContinue
                if ($app.DisplayName -like "*Teams Machine-Wide Installer*") {
                    # Check if already added
                    $alreadyAdded = $installations | Where-Object { $_.Type -eq "MWI" -and $_.DisplayName -eq $app.DisplayName }
                    if (-not $alreadyAdded) {
                        $installations += @{
                            Type = "MWI"
                            Path = $app.InstallLocation
                            Uninstaller = $app.UninstallString
                            DisplayName = $app.DisplayName
                            Profile = "Machine-Wide"
                            Version = $app.DisplayVersion
                        }
                        Write-Log "Found: $($app.DisplayName) (v$($app.DisplayVersion))" -Level Verbose
                    }
                }
            }
        }
        catch {
            Write-Log "Error scanning $path : $($_.Exception.Message)" -Level Verbose
        }
    }
    
    return $installations
}

function Backup-TeamsData {
    if ($SkipBackup) {
        Write-Log "Backup skipped by user request" -Level Info
        return $true
    }
    
    Write-Log "Backing up Teams data..." -Level Info
    
    try {
        # Create backup directory
        if (-not (Test-Path $script:BackupPath)) {
            New-Item -Path $script:BackupPath -ItemType Directory -Force | Out-Null
            Write-Log "Created backup directory: $script:BackupPath" -Level Verbose
        }
        
        $userProfiles = Get-AllUserProfiles
        $backupCount = 0
        
        foreach ($profile in $userProfiles) {
            $teamsDataPath = Join-Path $profile "AppData\Roaming\Microsoft\Teams"
            
            if (Test-Path $teamsDataPath) {
                $profileName = Split-Path $profile -Leaf
                $profileBackupPath = Join-Path $script:BackupPath $profileName
                
                try {
                    Copy-Item -Path $teamsDataPath -Destination $profileBackupPath -Recurse -Force -ErrorAction Stop
                    $backupCount++
                    Write-Log "Backed up Teams data for: $profileName" -Level Verbose
                }
                catch {
                    Write-Log "Could not backup $profileName : $($_.Exception.Message)" -Level Warning
                }
            }
        }
        
        if ($backupCount -gt 0) {
            Write-Log "Backed up Teams data for $backupCount user(s) to: $script:BackupPath" -Level Success
            return $true
        } else {
            Write-Log "No Teams data found to backup" -Level Info
            return $true
        }
    }
    catch {
        Write-Log "Backup failed: $($_.Exception.Message)" -Level Warning
        return $false
    }
}

function Uninstall-ClassicTeams {
    Write-Log "Detecting Classic Teams installations..." -Level Info
    
    $installations = Get-ClassicTeamsInstallation
    
    if ($installations.Count -eq 0) {
        Write-Log "No Classic Teams installations found" -Level Info
        return $true
    }
    
    Write-Log "Found $($installations.Count) Classic Teams installation(s)" -Level Warning
    
    # Store for potential rollback
    $script:RollbackData = $installations
    
    $successCount = 0
    $failCount = 0
    
    foreach ($install in $installations) {
        Write-Log "Uninstalling: $($install.Type) - $($install.Path) (v$($install.Version))" -Level Info
        
        try {
            if ($install.Type -eq "MWI" -and $install.Uninstaller) {
                # Parse uninstall string - improved regex
                if ($install.Uninstaller -match 'MsiExec\.exe\s*(/[IX]|--)\s*(\{[0-9A-Fa-f\-]+\})') {
                    $productCode = $matches[2]
                    Write-Log "Uninstalling MWI with product code: $productCode" -Level Info
                    
                    $msiArgs = @(
                        "/x"
                        $productCode
                        "/qn"
                        "/norestart"
                        "/L*v"
                        "`"$env:TEMP\TeamsMWI_Uninstall_$((Get-Date).Ticks).log`""
                    )
                    
                    $process = Start-Process "msiexec.exe" -ArgumentList $msiArgs -Wait -PassThru -NoNewWindow
                    
                    if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                        Write-Log "MWI uninstalled successfully (Exit code: $($process.ExitCode))" -Level Success
                        $successCount++
                    } elseif ($process.ExitCode -eq 1605) {
                        Write-Log "MWI not found or already uninstalled (Exit code: 1605)" -Level Info
                        $successCount++
                    } else {
                        Write-Log "MWI uninstall completed with exit code: $($process.ExitCode)" -Level Warning
                        $failCount++
                    }
                } else {
                    Write-Log "Could not parse MWI uninstall string: $($install.Uninstaller)" -Level Warning
                    $failCount++
                }
            }
            elseif ($install.Uninstaller -and (Test-Path $install.Uninstaller)) {
                Write-Log "Running uninstaller: $($install.Uninstaller)" -Level Verbose
                
                # Use Update.exe to uninstall
                $uninstallArgs = @("--uninstall", "-s")
                $process = Start-Process -FilePath $install.Uninstaller -ArgumentList $uninstallArgs -Wait -PassThru -NoNewWindow -ErrorAction Stop
                
                if ($process.ExitCode -eq 0) {
                    Write-Log "Uninstalled successfully from: $($install.Path)" -Level Success
                    $successCount++
                } else {
                    Write-Log "Uninstall completed with exit code: $($process.ExitCode)" -Level Warning
                    $failCount++
                }
            } else {
                Write-Log "Uninstaller not found: $($install.Uninstaller)" -Level Warning
                $failCount++
            }
            
            Start-Sleep -Milliseconds 500
        }
        catch {
            Write-Log "Error uninstalling from $($install.Path): $($_.Exception.Message)" -Level Error
            $failCount++
        }
    }
    
    Write-Log "Uninstall summary: $successCount succeeded, $failCount failed" -Level Info
    return ($successCount -gt 0 -or $failCount -eq 0)
}

function Remove-TeamsCache {
    param([switch]$AllUsers)
    
    Write-Log "Cleaning Teams cache and residual data..." -Level Info
    
    $cachePaths = @(
        'AppData\Local\Microsoft\Teams',
        'AppData\Roaming\Microsoft\Teams',
        'AppData\Local\Microsoft\TeamsMeetingAddin',
        'AppData\Local\SquirrelTemp',
        'AppData\Local\Microsoft\TeamsPresenceAddin',
        'AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache'
    )
    
    $profiles = if ($AllUsers) {
        Get-AllUserProfiles
    } else {
        @($env:USERPROFILE)
    }
    
    $removedCount = 0
    $skippedCount = 0
    
    foreach ($profile in $profiles) {
        foreach ($cachePath in $cachePaths) {
            $fullPath = Join-Path $profile $cachePath
            
            if (Test-Path $fullPath) {
                try {
                    Remove-Item -Path $fullPath -Recurse -Force -ErrorAction Stop
                    $removedCount++
                    Write-Log "Removed: $fullPath" -Level Verbose
                }
                catch {
                    Write-Log "Could not remove: $fullPath - $($_.Exception.Message)" -Level Verbose
                    $skippedCount++
                }
            }
        }
    }
    
    if ($removedCount -gt 0) {
        Write-Log "Cleaned $removedCount cache location(s) ($skippedCount skipped)" -Level Success
    } else {
        Write-Log "No cache data found to clean" -Level Info
    }
}

function Get-NewTeamsInstallation {
    # Check for New Teams MSIX package
    try {
        $msixPackage = Get-AppxPackage -Name "MSTeams" -AllUsers -ErrorAction Stop
        
        if ($msixPackage) {
            return @{
                Installed = $true
                Version = $msixPackage.Version
                InstallLocation = $msixPackage.InstallLocation
                PackageFullName = $msixPackage.PackageFullName
            }
        }
    }
    catch {
        Write-Log "Could not query MSIX packages: $($_.Exception.Message)" -Level Verbose
    }
    
    # Check via registry
    $regPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\MSTeams"
    if (Test-Path $regPath) {
        try {
            $regInfo = Get-ItemProperty -Path $regPath -ErrorAction Stop
            return @{
                Installed = $true
                Version = $regInfo.DisplayVersion
                InstallLocation = $regInfo.InstallLocation
            }
        }
        catch {
            Write-Log "Could not read registry: $($_.Exception.Message)" -Level Verbose
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
                return $true  # Return true since Teams is already there
            }
        } else {
            Write-Log "Existing installation will be updated/repaired" -Level Info
        }
    }
    
    Write-Log "Downloading New Teams installer..." -Level Info
    
    # Primary and fallback URLs
    $installerUrls = @(
        "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409",
        "https://statics.teams.cdn.office.net/production-windows-x64/enterprise/webview2/lkg/MSTeams-x64.msix"
    )
    
    $installerPath = "$env:TEMP\teamsbootstrapper_$((Get-Date).Ticks).exe"
    $downloadSuccess = $false
    
    foreach ($url in $installerUrls) {
        try {
            Write-Log "Attempting download from: $url" -Level Verbose
            
            # Download with timeout
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $url -OutFile $installerPath -UseBasicParsing -TimeoutSec 120 -ErrorAction Stop
            $ProgressPreference = 'Continue'
            
            if (Test-Path $installerPath) {
                $fileSize = (Get-Item $installerPath).Length / 1MB
                if ($fileSize -gt 1) {  # Sanity check - should be at least 1MB
                    Write-Log "Downloaded installer: $([math]::Round($fileSize, 2)) MB" -Level Success
                    $downloadSuccess = $true
                    break
                } else {
                    Write-Log "Downloaded file too small ($([math]::Round($fileSize, 2)) MB), trying next URL..." -Level Warning
                    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
                }
            }
        }
        catch {
            Write-Log "Download failed from $url : $($_.Exception.Message)" -Level Warning
        }
    }
    
    if (-not $downloadSuccess) {
        Write-Log "Failed to download installer from all sources" -Level Error
        return $false
    }
    
    Write-Log "Installing New Teams (this may take several minutes)..." -Level Info
    
    try {
        # Install with machine-wide provisioning
        $installArgs = @(
            "-p"  # Provision for all users
            "-o", "`"$env:TEMP\TeamsInstall_$((Get-Date).Ticks).log`""
        )
        
        # Create job to handle timeout
        $job = Start-Job -ScriptBlock {
            param($FilePath, $Args)
            $process = Start-Process -FilePath $FilePath -ArgumentList $Args -Wait -PassThru -NoNewWindow
            return $process.ExitCode
        } -ArgumentList $installerPath, $installArgs
        
        $completed = Wait-Job $job -Timeout $InstallTimeout
        
        if ($completed) {
            $exitCode = Receive-Job $job
            Remove-Job $job -Force
            
            if ($exitCode -eq 0) {
                Write-Log "New Teams installed successfully" -Level Success
                
                # Clean up installer
                Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
                
                return $true
            } else {
                Write-Log "Installation completed with exit code: $exitCode" -Level Warning
                Write-Log "Check log at: $env:TEMP\TeamsInstall*.log" -Level Info
                return $false
            }
        } else {
            # Timeout occurred
            Write-Log "Installation timed out after $InstallTimeout seconds" -Level Error
            Stop-Job $job -Force
            Remove-Job $job -Force
            
            # Try to kill installer process
            Get-Process | Where-Object { $_.Path -eq $installerPath } | Stop-Process -Force -ErrorAction SilentlyContinue
            
            return $false
        }
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level Error
        return $false
    }
    finally {
        # Cleanup installer file
        if (Test-Path $installerPath) {
            Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Test-NewTeamsInstallation {
    Write-Log "Verifying New Teams installation..." -Level Info
    
    Start-Sleep -Seconds 5
    
    $newTeams = Get-NewTeamsInstallation
    
    if ($newTeams.Installed) {
        Write-Log "✅ Verification successful" -Level Success
        Write-Log "   Version: $($newTeams.Version)" -Level Info
        if ($newTeams.InstallLocation) {
            Write-Log "   Location: $($newTeams.InstallLocation)" -Level Info
        }
        
        # Check if executable exists
        $exePaths = @(
            "$env:LOCALAPPDATA\Microsoft\WindowsApps\ms-teams.exe",
            "$env:ProgramFiles\WindowsApps\MSTeams_*\ms-teams.exe"
        )
        
        foreach ($path in $exePaths) {
            $resolved = Resolve-Path $path -ErrorAction SilentlyContinue
            if ($resolved) {
                Write-Log "   Executable: $($resolved.Path)" -Level Verbose
                break
            }
        }
        
        return $true
    } else {
        Write-Log "⚠️ New Teams installation could not be verified" -Level Warning
        Write-Log "Manual check: Get-AppxPackage -Name MSTeams" -Level Info
        return $false
    }
}

function Invoke-Rollback {
    Write-Log "Attempting rollback..." -Level Warning
    
    if (-not $script:BackupPath -or -not (Test-Path $script:BackupPath)) {
        Write-Log "No backup available for rollback" -Level Error
        return $false
    }
    
    Write-Log "Restoring Teams data from: $script:BackupPath" -Level Info
    
    try {
        $backupFolders = Get-ChildItem -Path $script:BackupPath -Directory
        $restoreCount = 0
        
        foreach ($folder in $backupFolders) {
            $profileName = $folder.Name
            $userProfile = Get-AllUserProfiles | Where-Object { $_ -like "*\$profileName" } | Select-Object -First 1
            
            if ($userProfile) {
                $restorePath = Join-Path $userProfile "AppData\Roaming\Microsoft\Teams"
                
                try {
                    if (Test-Path $restorePath) {
                        Remove-Item -Path $restorePath -Recurse -Force -ErrorAction Stop
                    }
                    
                    Copy-Item -Path $folder.FullName -Destination $restorePath -Recurse -Force -ErrorAction Stop
                    $restoreCount++
                    Write-Log "Restored Teams data for: $profileName" -Level Success
                }
                catch {
                    Write-Log "Could not restore $profileName : $($_.Exception.Message)" -Level Warning
                }
            }
        }
        
        Write-Log "Rollback completed: Restored $restoreCount user profile(s)" -Level Success
        return $true
    }
    catch {
        Write-Log "Rollback failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}
#endregion

#region Main Script
Write-Host "`n╔══════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║     Classic Teams → New Teams Migration Tool        ║" -ForegroundColor Cyan
Write-Host "║                   Version 2.1                        ║" -ForegroundColor Cyan
Write-Host "╚══════════════════════════════════════════════════════╝`n" -ForegroundColor Cyan

Write-Log "Migration started" -Level Info
Write-Log "Log file: $LogPath" -Level Info
Write-Log "Timeout: $InstallTimeout seconds" -Level Verbose

# Step 1: Prerequisites
Write-Host "`n[1/7] Checking prerequisites..." -ForegroundColor Yellow
if (-not (Test-Prerequisites)) {
    Write-Log "Prerequisites check failed. Exiting." -Level Error
    exit 1
}

# Step 2: Backup
Write-Host "`n[2/7] Backing up Teams data..." -ForegroundColor Yellow
if (-not (Backup-TeamsData)) {
    Write-Log "Backup failed, but continuing..." -Level Warning
}

# Step 3: Stop Teams
Write-Host "`n[3/7] Stopping Teams processes..." -ForegroundColor Yellow
if (-not (Stop-TeamsProcesses -MaxRetries 3)) {
    Write-Log "Warning: Some Teams processes may still be running" -Level Warning
    if (-not $Force) {
        $response = Read-Host "Continue anyway? (Y/N)"
        if ($response -ne 'Y') {
            Write-Log "Migration cancelled by user" -Level Info
            exit 0
        }
    }
}

# Step 4: Uninstall Classic Teams
Write-Host "`n[4/7] Uninstalling Classic Teams..." -ForegroundColor Yellow
if (-not (Uninstall-ClassicTeams)) {
    Write-Log "Failed to uninstall Classic Teams" -Level Error
    if (-not $Force) {
        Write-Log "Migration aborted" -Level Error
        exit 1
    }
}

# Step 5: Clean cache
Write-Host "`n[5/7] Cleaning Teams cache..." -ForegroundColor Yellow
Remove-TeamsCache -AllUsers

# Step 6: Install New Teams
Write-Host "`n[6/7] Installing New Teams..." -ForegroundColor Yellow
$installSuccess = Install-NewTeams

if (-not $installSuccess) {
    Write-Log "Installation failed" -Level Error
    
    if (-not $SkipBackup -and (Test-Path $script:BackupPath)) {
        Write-Log "Backup available at: $script:BackupPath" -Level Info
        if (-not $Force) {
            $response = Read-Host "Attempt rollback? (Y/N)"
            if ($response -eq 'Y') {
                Invoke-Rollback
            }
        }
    }
    
    exit 1
}

# Step 7: Verify installation
Write-Host "`n[7/7] Verifying installation..." -ForegroundColor Yellow
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
    Write-Host "   4. Review log: $LogPath" -ForegroundColor Gray
    
    if (-not $SkipBackup -and (Test-Path $script:BackupPath)) {
        Write-Host "   5. Backup saved to: $script:BackupPath`n" -ForegroundColor Gray
    } else {
        Write-Host ""
    }
    
    exit 0
} else {
    Write-Host "`n⚠️ Migration completed with warnings`n" -ForegroundColor Yellow
    Write-Host "Review log for details: $LogPath`n" -ForegroundColor Gray
    exit 1
}

Write-Host ("═" * 60) -ForegroundColor DarkGray
Write-Log "Migration process completed" -Level Info
#endregion

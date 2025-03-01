<#
.SYNOPSIS
    Complete Windows 11 24H2 Update Script with Automatic ISO Download and TPM/CPU Bypass
.DESCRIPTION
    This script automates the entire Windows 11 24H2 update process by:
    1. Downloading the latest Windows 11 24H2 ISO
    2. Applying all necessary TPM and CPU compatibility bypasses
    3. Mounting the ISO and running the update with appropriate options
    4. Preserving all files and programs during the update
.NOTES
    Author: Earl Daniels
    Date: February 28, 2025
    Requires: Administrator privileges, PowerShell 5.1+
#>

# Script configuration
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue" # Speeds up downloads
$IsoPath = "$env:USERPROFILE\Downloads\Windows11_24H2.iso"
$LogPath = "$env:SystemDrive\Scripts\logs"
$LogFile = "$LogPath\Win11_24H2_Update.log"

# Console colors
$colors = @{
    "Success" = "Green"
    "Info" = "Cyan"
    "Warning" = "Yellow"
    "Error" = "Red"
    "Step" = "Magenta"
    "Default" = "White"
}

# Functions for logging and console output
function Write-ColorOutput {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [string]$Type = "Default",
        
        [Parameter(Mandatory=$false)]
        [switch]$NoNewLine
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = $colors[$Type]
    
    # Log to file
    if (-not (Test-Path $LogPath)) {
        New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
    }
    "$timestamp - $Message" | Out-File -FilePath $LogFile -Append
    
    # Write to console with color
    if ($NoNewLine) {
        Write-Host $Message -ForegroundColor $color -NoNewline
    } else {
        Write-Host $Message -ForegroundColor $color
    }
}


function Write-Step {
    param([string]$StepNumber, [string]$StepDescription)
    
    $separator = "=" * 80
    Write-ColorOutput $separator "Info"
    Write-ColorOutput "STEP $StepNumber : $StepDescription" "Step"
    Write-ColorOutput $separator "Info"
}

function Exit-WithError {
    param([string]$ErrorMessage)
    
    Write-ColorOutput $ErrorMessage "Error"
    Write-ColorOutput "Script execution stopped due to error. Check log at: $LogFile" "Error"
    return
}

# Function to check if Windows 11 24H2 update is needed
function Test-NeedsUpdate {
    $osInfo = Get-ComputerInfo
    
    # Display current Windows version information
    Write-ColorOutput "Current System: $($osInfo.WindowsProductName) $($osInfo.WindowsVersion)" "Info"
    Write-ColorOutput "Build Number: $($osInfo.OsBuildNumber)" "Info"
    
    # No need to update if already on 24H2
    if ($osInfo.OsBuildNumber -ge 26100) {
        Write-ColorOutput "Your system is already running Windows 11 version 24H2 or newer." "Success"
        return $false
    }
    
    # Check if running Windows 11
    if ($osInfo.OsBuildNumber -lt 22000) {
        Write-ColorOutput "This script is designed to update Windows 11 to 24H2. You appear to be running Windows 10." "Warning"
        $confirm = Read-Host "Would you like to continue anyway? This might not work as expected. (y/n)"
        if ($confirm -ne "y") {
            return $false
        }
    }
    
    return $true
}

# Function to create the TPM/CPU bypass script
function Install-CompatibilityBypass {
    Write-ColorOutput "Creating TPM/CPU compatibility bypass..." "Info"

    # Batch script content
    $batchScript = @'
@echo off & title Windows 11 Compatibility Bypass
setlocal EnableDelayedExpansion

:: Set console colors
color 0B
echo [94m===========================================================================
echo                  WINDOWS 11 COMPATIBILITY BYPASS UTILITY
echo ===========================================================================[0m

:: Create log directory
if not exist "%SystemDrive%\Scripts\logs" mkdir "%SystemDrive%\Scripts\logs" >nul 2>nul
set LOGFILE=%SystemDrive%\Scripts\logs\win11bypass.log

:: Log start
echo %date% %time% - Compatibility bypass started >> %LOGFILE%

:: Apply registry bypasses for TPM and CPU compatibility
echo [96mApplying registry bypasses for TPM and CPU compatibility checks...[0m
echo %date% %time% - Setting registry bypass keys >> %LOGFILE%

reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\TargetReleaseVersion" /f /v TargetReleaseVersionInfo /d "24H2" /t reg_sz >nul 2>>%LOGFILE%
reg add "HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" /f /v DisableWUfBSafeguards /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "HKLM\SYSTEM\Setup\MoSetup" /f /v AllowUpgradesWithUnsupportedTPMorCPU /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "HKLM\SYSTEM\Setup\LabConfig" /f /v BypassTPMCheck /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "HKLM\SYSTEM\Setup\LabConfig" /f /v BypassSecureBootCheck /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "HKLM\SYSTEM\Setup\LabConfig" /f /v BypassRAMCheck /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "HKLM\SYSTEM\Setup\LabConfig" /f /v BypassStorageCheck /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "HKLM\SYSTEM\Setup\LabConfig" /f /v BypassCPUCheck /d 1 /t reg_dword >nul 2>>%LOGFILE%

:: Modify setuphost behaviors via IFEO
echo [96mSetting up SetupHost.exe interception for compatibility bypass...[0m
echo %date% %time% - Configuring SetupHost.exe interception >> %LOGFILE%

set IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options
reg add "%IFEO%\SetupHost.exe" /f /v UseFilter /d 1 /t reg_dword >nul 2>>%LOGFILE%
reg add "%IFEO%\SetupHost.exe\0" /f /v FilterFullPath /d "%SystemDrive%\$WINDOWS.~BT\Sources\SetupHost.exe" >nul 2>>%LOGFILE%
reg add "%IFEO%\SetupHost.exe\0" /f /v Debugger /d "%SystemDrive%\Scripts\setuphost_hook.cmd" >nul 2>>%LOGFILE%

:: Create the setuphost hook script
echo [96mCreating SetupHost hook script...[0m
if not exist "%SystemDrive%\Scripts" mkdir "%SystemDrive%\Scripts" >nul 2>nul

echo @echo off > "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo setlocal >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo set SOURCES=%%SystemDrive%%\$WINDOWS.~BT\Sources >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo set LOGFILE=%%SystemDrive%%\Scripts\logs\setuphost_hook.log >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo echo %%date%% %%time%% - SetupHost hook activated with args: %%* ^>^> %%LOGFILE%% >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo if not exist %%SOURCES%%\WindowsUpdateBox.exe mklink /h %%SOURCES%%\WindowsUpdateBox.exe %%SOURCES%%\SetupHost.exe ^>^>%%LOGFILE%% 2^>^&1 >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo if not exist %%SOURCES%%\appraiserres.dll echo. ^> %%SOURCES%%\appraiserres.dll >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo set OPT=/Compat IgnoreWarning /MigrateDrivers All /Telemetry Disable >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo %%SOURCES%%\WindowsUpdateBox.exe %%OPT%% %%* >> "%SystemDrive%\Scripts\setuphost_hook.cmd"
echo exit /b %%errorlevel%% >> "%SystemDrive%\Scripts\setuphost_hook.cmd"

:: Cleanup legacy items
echo [96mRemoving any conflicting legacy items...[0m
echo %date% %time% - Removing legacy bypass items >> %LOGFILE%
wmic /namespace:"\\root\subscription" path __EventFilter where Name="Skip TPM Check on Dynamic Update" delete >nul 2>>%LOGFILE%
reg delete "%IFEO%\vdsldr.exe" /f >nul 2>>%LOGFILE%

echo [92m===========================================================================
echo                  COMPATIBILITY BYPASS SUCCESSFULLY INSTALLED
echo ===========================================================================[0m
echo %date% %time% - Compatibility bypass successfully installed >> %LOGFILE%
exit /b 0
'@

    # Save batch script to file
    try {
        if (-not (Test-Path "$env:SystemDrive\Scripts")) {
            New-Item -Path "$env:SystemDrive\Scripts" -ItemType Directory -Force | Out-Null
        }
        Set-Content -Path "$env:SystemDrive\Scripts\win11bypass.cmd" -Value $batchScript -Force
        Write-ColorOutput "TPM/CPU compatibility bypass script created successfully." "Success"
    }
    catch {
        Exit-WithError "Failed to create compatibility bypass script: $_"
    }

    # Execute the batch script
    try {
        Write-ColorOutput "Executing compatibility bypass script..." "Info"
        $process = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $env:SystemDrive\Scripts\win11bypass.cmd" -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -ne 0) {
            Exit-WithError "Compatibility bypass script failed with exit code: $($process.ExitCode)"
        }
        Write-ColorOutput "Compatibility bypass successfully applied." "Success"
    }
    catch {
        Exit-WithError "Failed to execute compatibility bypass script: $_"
    }
}

# Fallback function for UUP dump method - previously missing
function Get-Windows11ISOFallback {
    param([string]$OutputPath)
    
    Write-ColorOutput "Attempting fallback download method using UUP dump..." "Info"
    
    # Microsoft's Edge browser user agent to prevent throttling
    $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0"
    
    # Create temporary directory for download scripts
    $tempDir = "$env:TEMP\Win11_24H2_Download_Fallback"
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    
    try {
        # Step 1: Download UUP dump script package
        Write-ColorOutput "Downloading UUP dump script package..." "Info"
        
        # UUP dump links for Windows 11 24H2 (Latest Retail build)
        $uupScriptUrl = "https://uupdump.net/get.php?id=d347bc84-5a90-4dfa-8012-cf5a82543c21&pack=en-us&edition=professional&download=1"
        
        Invoke-WebRequest -Uri $uupScriptUrl -OutFile "$tempDir\uup_download_windows.zip" -UserAgent $userAgent
        
        # Step 2: Extract the UUP dump script package
        Write-ColorOutput "Extracting UUP dump script package..." "Info"
        Expand-Archive -Path "$tempDir\uup_download_windows.zip" -DestinationPath $tempDir -Force
        
        # Step 3: Run the UUP dump script to download and convert to ISO
        Write-ColorOutput "Starting Windows 11 24H2 download and ISO creation..." "Info"
        Write-ColorOutput "This process will take some time depending on your internet speed..." "Warning"
        
        $convertScript = Get-ChildItem -Path $tempDir -Filter "*.cmd" | Where-Object { $_.Name -like "*convert*" } | Select-Object -First 1
        if ($null -eq $convertScript) {
            Exit-WithError "Could not find UUP conversion script in the downloaded package"
        }
        
        # Execute the UUP dump script
        $process = Start-Process -FilePath $convertScript.FullName -WorkingDirectory $tempDir -Wait -PassThru -NoNewWindow
        if ($process.ExitCode -ne 0) {
            Exit-WithError "UUP dump script failed with exit code: $($process.ExitCode)"
        }
        
        # Step 4: Find and move the created ISO
        $createdISO = Get-ChildItem -Path $tempDir -Filter "*.iso" | Select-Object -First 1
        if ($null -eq $createdISO) {
            Exit-WithError "Could not find created ISO in the download directory"
        }
        
        # Move ISO to final destination
        Move-Item -Path $createdISO.FullName -Destination $OutputPath -Force
        Write-ColorOutput "Windows 11 24H2 ISO successfully downloaded to: $OutputPath" "Success"
        
        # Cleanup temp directory
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        
        return $true
    }
    catch {
        Write-ColorOutput "UUP dump fallback method failed with error: $_" "Error"
        return $false
    }
}

# Main flow control function - properly structured
function Start-Win11Upgrade {
    Write-ColorOutput "=========================================================" "Info"
    Write-ColorOutput "      WINDOWS 11 24H2 AUTOMATED UPGRADE UTILITY" "Step"
    Write-ColorOutput "                  WITH TPM/CPU BYPASS" "Step"
    Write-ColorOutput "=========================================================" "Info"
    
    # Check if update is needed
    if (-not (Test-NeedsUpdate)) {
        Write-ColorOutput "No update needed. Exiting." "Info"
        return
    }
    
    # Apply compatibility bypasses first
    Write-Step "2" "Installing Windows 11 Compatibility Bypasses"
    Install-CompatibilityBypass
    
    # Define ISO path
    $IsoPath = "$env:USERPROFILE\Downloads\Windows11_24H2.iso"
    
    # Check first if ISO exists without downloading
    $existingIso = $null
    $possibleIsoPatterns = @(
        "$env:USERPROFILE\Downloads\*Win*11*.iso",
        "$env:USERPROFILE\Downloads\*Windows*11*.iso"
    )
    
    foreach ($pattern in $possibleIsoPatterns) {
        $matchingFiles = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        if ($matchingFiles -and $matchingFiles.Count -gt 0) {
            $existingIso = $matchingFiles[0]
            $IsoPath = $existingIso.FullName
            Write-ColorOutput "Found existing Windows 11 ISO: $IsoPath" "Success"
            break
        }
    }
    
    # If no ISO exists or user chooses to download a new one
    if (-not $existingIso) {
        $downloadISO = $true
        Write-ColorOutput "No existing Windows 11 ISO found. Proceeding to download..." "Info"
    } else {
        $fileSize = $existingIso.Length / 1GB
        Write-ColorOutput "Found existing Windows 11 ISO at: $($existingIso.FullName) (Size: $($fileSize.ToString('0.00')) GB)" "Info"
        $choice = Read-Host "Use this existing ISO for the update? (y/n) [Default: y]"
        $downloadISO = -not ([string]::IsNullOrEmpty($choice) -or $choice.ToLower() -eq "y")
    }
    
    # Download ISO if needed
    if ($downloadISO) {
        $success = Get-Windows11ISO -OutputPath $IsoPath
        if (-not $success) {
            Exit-WithError "Failed to obtain Windows 11 24H2 ISO"
        }
    }
    
    # Mount ISO and start update - will now always run if we have a valid ISO
    if (Test-Path $IsoPath) {
        Start-WindowsUpdate -IsoPath $IsoPath
    } else {
        Exit-WithError "Windows 11 ISO not found at path: $IsoPath"
    }
}

# Function to download Windows 11 24H2 ISO using Fido
function Get-Windows11ISO {
    param([string]$OutputPath)
    
    # Set defaults for Windows 11 ISO
    $defaultLang = "en-US"
    $windowsVersion = "11"
    $windowsEdition = "Pro"
    $architecture = "x64"
    
    Write-ColorOutput "Preparing to download Windows 11 24H2 ISO using Fido..." "Info"
    
    # Define potential ISO locations
    $outputDir = Split-Path -Parent $OutputPath
    $possibleIsoPatterns = @(
        "$outputDir\*Win*11*.iso",
        "$outputDir\*Windows*11*.iso",
        "$env:USERPROFILE\Downloads\*Win*11*.iso",
        "$env:USERPROFILE\Downloads\*Windows*11*.iso"
    )
    
    # Check if any Windows 11 ISO already exists in common locations
    $existingIso = $null
    foreach ($pattern in $possibleIsoPatterns) {
        $matchingFiles = Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
        if ($matchingFiles -and $matchingFiles.Count -gt 0) {
            $existingIso = $matchingFiles[0]
            break
        }
    }
    
    # If we found an existing ISO, ask if we should use it
    if ($existingIso) {
        $fileSize = $existingIso.Length / 1GB
        Write-ColorOutput "Found existing Windows 11 ISO at: $($existingIso.FullName) (Size: $($fileSize.ToString('0.00')) GB)" "Info"
        $choice = Read-Host "Use this existing ISO? (y/n) [Default: y]"
        if ([string]::IsNullOrEmpty($choice) -or $choice.ToLower() -eq "y") {
            # If the ISO is not at the expected location, copy it there
            if ($existingIso.FullName -ne $OutputPath) {
                Write-ColorOutput "Copying ISO to target location..." "Info"
                Copy-Item -Path $existingIso.FullName -Destination $OutputPath -Force
            }
            return $true
        }
    } elseif (Test-Path $OutputPath) {
        # If the ISO exists exactly at the specified path
        $fileSize = (Get-Item $OutputPath).Length / 1GB
        Write-ColorOutput "Windows 11 ISO already exists at: $OutputPath (Size: $($fileSize.ToString('0.00')) GB)" "Info"
        $choice = Read-Host "Use existing ISO? (y/n) [Default: y]"
        if ([string]::IsNullOrEmpty($choice) -or $choice.ToLower() -eq "y") {
            return $true
        }
    }
    
    # Create temporary directory for Fido
    $tempDir = "$env:TEMP\Win11_24H2_Download"
    if (Test-Path $tempDir) {
        Remove-Item -Path $tempDir -Recurse -Force
    }
    New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
    
    try {
        # Step 1: Download Fido
        Write-Step "1" "Downloading Fido tool for Windows 11 ISO acquisition"
        Write-ColorOutput "Downloading Fido tool..." "Info"
        
        $fidoUrl = "https://github.com/pbatard/Fido/archive/refs/heads/master.zip"
        $fidoZip = "$tempDir\fido.zip"
        $fidoDir = "$tempDir\Fido"
        
        # Download Fido
        Invoke-WebRequest -Uri $fidoUrl -OutFile $fidoZip
        
        # Extract Fido
        Expand-Archive -Path $fidoZip -DestinationPath $tempDir -Force
        $fidoMasterDir = Get-ChildItem -Path $tempDir -Directory | Where-Object { $_.Name -like "Fido-*" } | Select-Object -First 1
        if ($fidoMasterDir) {
            Move-Item -Path $fidoMasterDir.FullName -Destination $fidoDir -Force
        } else {
            Exit-WithError "Could not find Fido directory after extraction"
        }
        
        # Step 2: Run Fido to download Windows 11 24H2
        Write-ColorOutput "Starting Windows 11 24H2 download with Fido..." "Info"
        Write-ColorOutput "This process will take some time depending on your internet speed..." "Warning"
        
        $fidoScript = "$fidoDir\Fido.ps1"
        if (-not (Test-Path $fidoScript)) {
            Exit-WithError "Could not find Fido.ps1 script in the downloaded package"
        }
        
        # Get the directory for the ISO
        $outputDir = Split-Path -Parent $OutputPath
        
        # Ask user for language preference with default
        Write-ColorOutput "Windows 11 language options:" "Info"
        Write-ColorOutput "1. English US (en-US) [Default]" "Info"
        Write-ColorOutput "2. English UK (en-GB)" "Info"
        Write-ColorOutput "3. German (de-DE)" "Info"
        Write-ColorOutput "4. French (fr-FR)" "Info"
        Write-ColorOutput "5. Spanish (es-ES)" "Info"
        Write-ColorOutput "6. Show all available languages" "Info"
        
        $langChoice = Read-Host "Select language option (1-6) or enter language code directly [Default: 1]"
        
        $language = $defaultLang
        
        switch ($langChoice) {
            "2" { $language = "en-GB" }
            "3" { $language = "de-DE" }
            "4" { $language = "fr-FR" }
            "5" { $language = "es-ES" }
            "6" { 
                # Show all available languages
                Write-ColorOutput "Retrieving available languages..." "Info"
                $process = Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$fidoScript`" -GetLangs" -Wait -PassThru -NoNewWindow
                $customLang = Read-Host "Enter your preferred language code (e.g., en-US, de-DE) [Default: en-US]"
                if (-not [string]::IsNullOrEmpty($customLang)) {
                    $language = $customLang
                }
            }
            default {
                # If input is not empty and doesn't match options 1-5, assume it's a direct language code
                if (-not [string]::IsNullOrEmpty($langChoice) -and $langChoice -ne "1") {
                    $language = $langChoice
                }
            }
        }
        
        # Now download the actual ISO with all parameters explicitly set
        Write-ColorOutput "Downloading Windows 11 with language: $language..." "Info"
        
        # Construct arguments with all explicitly defined parameters
        $fidoArgs = "-ExecutionPolicy Bypass -File `"$fidoScript`" -Win $windowsVersion -Ed $windowsEdition -Lang $language -Arch $architecture -OutDir `"$outputDir`""
        
        # Execute Fido with the prepared arguments
        $process = Start-Process -FilePath "powershell.exe" -ArgumentList $fidoArgs -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -ne 0) {
            Exit-WithError "Fido download failed with exit code: $($process.ExitCode)"
        }
        
        # Step 3: Find the downloaded ISO and rename if needed
        $downloadedISO = Get-ChildItem -Path $outputDir -Filter "*.iso" | Where-Object { $_.Name -like "*windows*11*" -or $_.Name -like "*win*11*" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($null -eq $downloadedISO) {
            Exit-WithError "Could not find downloaded Windows 11 ISO in $outputDir"
        }
        
        # Rename the ISO if it's not already at the target path
        if ($downloadedISO.FullName -ne $OutputPath) {
            Move-Item -Path $downloadedISO.FullName -Destination $OutputPath -Force
        }
        
        Write-ColorOutput "Windows 11 24H2 ISO successfully downloaded to: $OutputPath" "Success"
        
        # Cleanup temp directory but keep the ISO
        Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        
        return $true
    }
    catch {
        Write-ColorOutput "Fido ISO download failed with error: $_" "Error"
        
        # Fallback to manual download
        Write-ColorOutput "Would you like to:" "Warning"
        Write-ColorOutput "1. Try alternative download method (UUP dump) [Default]" "Info"
        Write-ColorOutput "2. Download Windows 11 manually and provide the path" "Info"
        $fallbackChoice = Read-Host "Enter your choice (1 or 2) [Default: 1]"
        
        if ([string]::IsNullOrEmpty($fallbackChoice) -or $fallbackChoice -eq "1") {
            return Get-Windows11ISOFallback -OutputPath $OutputPath
        }
        else {
            Write-ColorOutput "Please download Windows 11 24H2 ISO manually from Microsoft's website:" "Warning"
            Write-ColorOutput "https://www.microsoft.com/software-download/windows11" "Info"
            
            $manualDownloadPath = Read-Host "Enter the full path to your manually downloaded Windows 11 24H2 ISO"
            if (Test-Path $manualDownloadPath) {
                Copy-Item -Path $manualDownloadPath -Destination $OutputPath -Force
                Write-ColorOutput "Using manually provided ISO: $manualDownloadPath" "Success"
                return $true
            }
            else {
                Exit-WithError "Could not find ISO at specified path: $manualDownloadPath"
            }
        }
    }
}

# Function to mount ISO and run silent setup - Fixed function structure
function Start-WindowsUpdate {
    try {
        # Run the upgrade process
        Start-Win11Upgrade
    } catch {
        Write-Error "An error occurred during the Windows 11 upgrade: $_"
    }

    param([string]$IsoPath)
    
    Write-Step "3" "Starting Windows 11 24H2 Silent Update Process"
    
    # Mount the ISO
    try {
        Write-ColorOutput "Mounting Windows 11 ISO..." "Info"
        $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru
        $driveLetter = ($mountResult | Get-Volume).DriveLetter
        Write-ColorOutput "ISO mounted successfully on drive $($driveLetter):" "Success"
    }
    catch {
        Exit-WithError "Failed to mount ISO: $_"
    }
    
    # Create a comprehensive autounattend.xml file for completely silent installation
    $autoUnattendPath = "$env:TEMP\autounattend.xml"
    $autoUnattendXml = @"
<?xml version="1.0" encoding="utf-8"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend">
    <settings pass="windowsPE">
        <component name="Microsoft-Windows-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <UserData>
                <AcceptEula>true</AcceptEula>
            </UserData>
            <ComplianceCheck>
                <DisplayReport>never</DisplayReport>
            </ComplianceCheck>
            <Diagnostics>
                <OptIn>false</OptIn>
            </Diagnostics>
            <DynamicUpdate>
                <Enable>false</Enable>
                <WillShowUI>never</WillShowUI>
            </DynamicUpdate>
            <Display>
                <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
            </Display>
        </component>
        <component name="Microsoft-Windows-International-Core-WinPE" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <SetupUILanguage>
                <UILanguage>en-US</UILanguage>
                <WillShowUI>never</WillShowUI>
            </SetupUILanguage>
        </component>
    </settings>
    <settings pass="oobeSystem">
        <component name="Microsoft-Windows-Shell-Setup" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <OOBE>
                <HideEULAPage>true</HideEULAPage>
                <HideLocalAccountScreen>true</HideLocalAccountScreen>
                <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
                <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
                <NetworkLocation>Home</NetworkLocation>
                <ProtectYourPC>1</ProtectYourPC>
                <SkipMachineOOBE>true</SkipMachineOOBE>
                <SkipUserOOBE>true</SkipUserOOBE>
            </OOBE>
            <UserAccounts>
                <DontDisplayLastUserName>false</DontDisplayLastUserName>
            </UserAccounts>
        </component>
        <component name="Microsoft-Windows-International-Core" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
            <InputLocale>en-US</InputLocale>
            <SystemLocale>en-US</SystemLocale>
            <UILanguage>en-US</UILanguage>
            <UserLocale>en-US</UserLocale>
        </component>
    </settings>
</unattend>
"@

    try {
        Set-Content -Path $autoUnattendPath -Value $autoUnattendXml -Force
        Write-ColorOutput "Created silent setup configuration file" "Success"
    }
    catch {
        Exit-WithError "Failed to create autounattend.xml file: $_"
    }
    
    # Create an EI.cfg file to prevent edition upgrade prompts
    $eiCfgPath = "$env:TEMP\EI.cfg"
    $eiCfgContent = @"
[Channel]
_Default
"@

    try {
        Set-Content -Path $eiCfgPath -Value $eiCfgContent -Force
        
        # Copy EI.cfg to the root of the mounted ISO drive
        Copy-Item -Path $eiCfgPath -Destination "${driveLetter}:\" -Force -ErrorAction SilentlyContinue
        Write-ColorOutput "Created edition configuration file" "Success"
    }
    catch {
        Write-ColorOutput "Warning: Failed to create EI.cfg file: $_" "Warning"
        # Continue anyway as this is not critical
    }
    
    # Create a setup configuration file that enables silent setup
    $setupConfigPath = "$env:TEMP\SetupConfig.ini"
    $setupConfigContent = @"
[SetupConfig]
Priority=5
BitLocker=AlwaysSuspend
Compat=IgnoreWarning
MigrateDrivers=All
DynamicUpdate=Disable
ShowOOBE=None
Telemetry=Disable
"@

    try {
        Set-Content -Path $setupConfigPath -Value $setupConfigContent -Force
        Write-ColorOutput "Created setup configuration file" "Success"
    }
    catch {
        Exit-WithError "Failed to create SetupConfig.ini file: $_"
    }
    
    # Start the Windows Setup
    try {
        $setupPath = "${driveLetter}:\setup.exe"
        if (-not (Test-Path $setupPath)) {
            Exit-WithError "Could not find setup.exe on the mounted ISO"
        }
        
        Write-ColorOutput "Starting silent Windows 11 24H2 Update..." "Info"
        
        # Build setup arguments for completely silent operation
        $setupArgs = "/auto upgrade /quiet /noreboot /migratedrivers all /showoobe none /telemetry disable /dynamicupdate disable /compat ignorewarning"
        
        # Add unattend and config files
        if (Test-Path $autoUnattendPath) {
            $setupArgs += " /unattend:`"$autoUnattendPath`""
        }
        
        if (Test-Path $setupConfigPath) {
            $setupArgs += " /configfile:`"$setupConfigPath`""
        }
        
        # Create a script to monitor the upgrade progress
        $progressMonitorPath = "$env:TEMP\UpgradeProgressMonitor.ps1"
        $progressMonitorContent = @"
# Windows 11 Upgrade Progress Monitor
`$setupProgressPath = "C:\`$WINDOWS.~BT\Sources\SetupProgress.xml"
`$logPath = "C:\`$WINDOWS.~BT\Sources\Panther"

function Write-ProgressUpdate {
    param([string]`$Message, [string]`$Color = "White")
    
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "`$timestamp - `$Message" -ForegroundColor `$Color
}

# Wait for setup to start
Write-ProgressUpdate "Waiting for Windows 11 upgrade to initialize..." "Cyan"
`$retryCount = 0
while (-not (Test-Path "C:\`$WINDOWS.~BT")) {
    Start-Sleep -Seconds 10
    `$retryCount++
    if (`$retryCount -gt 30) {
        Write-ProgressUpdate "Setup initialization timeout. Please check setup manually." "Red"
        exit 1
    }
}

Write-ProgressUpdate "Windows 11 upgrade process has started" "Green"

# Monitor progress
while (`$true) {
    if (Test-Path `$setupProgressPath) {
        try {
            `$progress = [xml](Get-Content `$setupProgressPath -ErrorAction SilentlyContinue)
            `$phase = `$progress.SetupProgress.InstallPhase
            `$percent = `$progress.SetupProgress.PercentComplete
            
            Write-ProgressUpdate "Phase: `$phase - Progress: `$percent%" "Yellow"
        }
        catch {
            Write-ProgressUpdate "Reading progress data..." "Cyan"
        }
    }
    elseif (Test-Path `$logPath) {
        `$latestLog = Get-ChildItem `$logPath -Filter "*.log" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (`$latestLog) {
            `$lastLines = Get-Content `$latestLog.FullName -Tail 5 -ErrorAction SilentlyContinue
            Write-ProgressUpdate "Latest log activity:" "Cyan"
            foreach (`$line in `$lastLines) {
                Write-Host "   `$line"
            }
        }
    }
    
    # Check if the upgrade has completed or failed
    if (-not (Test-Path "C:\`$WINDOWS.~BT")) {
        if (Test-Path "C:\Windows.old") {
            Write-ProgressUpdate "Windows 11 upgrade appears to have completed successfully!" "Green"
        }
        else {
            Write-ProgressUpdate "Windows 11 upgrade process may have failed or been cancelled." "Red"
        }
        break
    }
    
    Start-Sleep -Seconds 60
}
"@
        
        Set-Content -Path $progressMonitorPath -Value $progressMonitorContent -Force
        
        # Start setup process
        $setupProc = Start-Process -FilePath $setupPath -ArgumentList $setupArgs -PassThru -NoNewWindow
        Write-ColorOutput "Windows 11 24H2 Setup launched with Process ID: $($setupProc.Id)" "Success"
        Write-ColorOutput "Setup is now running in silent mode" "Info"
        Write-ColorOutput "The system will restart automatically when ready" "Warning"
        
        # Start the progress monitor in a new window
        Start-Process -FilePath "powershell.exe" -ArgumentList "-ExecutionPolicy Bypass -File `"$progressMonitorPath`"" -WindowStyle Normal
        
        # Add registry key to auto-login after upgrade
        try {
            $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            $username = ($currentUser -split '\\')[1]
            
            # Don't store the actual password, just enable auto-login for the current user after upgrade
            Write-ColorOutput "Setting up auto-login after upgrade for current user..." "Info"
            
            # Create the registry keys for autologon
            $autoLogonKey = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
            Set-ItemProperty -Path $autoLogonKey -Name "AutoAdminLogon" -Value "1" -Force
            Set-ItemProperty -Path $autoLogonKey -Name "DefaultUserName" -Value $username -Force
            # DefaultPassword is intentionally not set for security reasons
            Set-ItemProperty -Path $autoLogonKey -Name "AutoLogonCount" -Value "1" -Force
        }
        catch {
            Write-ColorOutput "Warning: Could not configure auto-login after upgrade: $_" "Warning"
            # Continue anyway as this is not critical
        }
        
        Write-ColorOutput "Progress monitor launched in a separate window" "Success"
        Write-ColorOutput "You can now safely close this window - the update will continue" "Info"
    }
    catch {
        Exit-WithError "Failed to start Windows 11 Setup: $_"
    }
}
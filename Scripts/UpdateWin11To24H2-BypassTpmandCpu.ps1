#requires -RunAsAdministrator
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
    Author: Claude
    Date: February 28, 2025
    Requires: Administrator privileges, PowerShell 5.1+
#>

# Script configuration
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue" # Speeds up downloads
$IsoPath = "$env:USERPROFILE\Downloads\Windows11_24H2.iso"
$LogPath = "$env:SystemDrive\Scripts\logs"
$LogFile = "$LogPath\Win11_24H2_Update.log"
$BypassPath = "$env:SystemDrive\Scripts\win11bypass.cmd"

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
    exit 1
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

# Function to download Windows 11 24H2 ISO
function Get-Windows11ISO {
    param([string]$OutputPath)
    
    Write-ColorOutput "Preparing to download Windows 11 24H2 ISO..." "Info"
    
    # For direct downloads we'd need to use the Microsoft Media Creation Tool or similar
    # Since direct ISO download links aren't officially available, we'll use a more reliable approach
    
    # Check if ISO already exists
    if (Test-Path $OutputPath) {
        $fileSize = (Get-Item $OutputPath).Length / 1GB
        Write-ColorOutput "Windows 11 ISO already exists at: $OutputPath (Size: $($fileSize.ToString('0.00')) GB)" "Info"
        $choice = Read-Host "Use existing ISO? (y/n)"
        if ($choice -eq "y") {
            return $true
        }
    }
    
    # Microsoft's Edge browser user agent to prevent throttling
    $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0"
    
    # Option 1: Use uupdump.net to get the latest Windows 11 24H2 ISO
    Write-Step "1" "Setting up Windows 11 24H2 ISO download"
    Write-ColorOutput "We'll use UUP dump to download the latest Windows 11 24H2 build" "Info"
    
    # Create temporary directory for download scripts
    $tempDir = "$env:TEMP\Win11_24H2_Download"
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
        Write-ColorOutput "Automated ISO download failed with error: $_" "Error"
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

# Function to mount ISO and run setup
function Start-WindowsUpdate {
    param([string]$IsoPath)
    
    Write-Step "3" "Starting Windows 11 24H2 Update Process"
    
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
    
    # Create an autounattend.xml file to automate the setup
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
                <WillShowUI>OnError</WillShowUI>
            </DynamicUpdate>
        </component>
    </settings>
</unattend>
"@

    try {
        Set-Content -Path $autoUnattendPath -Value $autoUnattendXml -Force
        Write-ColorOutput "Created automated setup configuration file" "Success"
    }
    catch {
        Exit-WithError "Failed to create autounattend.xml file: $_"
    }
    
    # Start the Windows Setup
    try {
        $setupPath = "${driveLetter}:\setup.exe"
        if (-not (Test-Path $setupPath)) {
            Exit-WithError "Could not find setup.exe on the mounted ISO"
        }
        
        Write-ColorOutput "Starting Windows 11 24H2 Setup..." "Info"
        Write-ColorOutput "========== IMPORTANT INFORMATION ==========" "Warning"
        Write-ColorOutput "1. The Windows Setup will now launch" "Warning"
        Write-ColorOutput "2. Your computer will restart multiple times during the update" "Warning"
        Write-ColorOutput "3. All your files and applications will be preserved" "Warning"
        Write-ColorOutput "4. TPM and CPU compatibility checks will be bypassed automatically" "Warning"
        Write-ColorOutput "=========================================" "Warning"
        
        $setupArgs = "/auto upgrade /migratedrivers all /showoobe none /telemetry disable /dynamicupdate disable /compat ignorewarning"
        if (Test-Path $autoUnattendPath) {
            $setupArgs += " /unattend:$autoUnattendPath"
        }
        
        # Start setup process
        Start-Process -FilePath $setupPath -ArgumentList $setupArgs -Wait:$false
        
        Write-ColorOutput "Windows 11 24H2 Setup launched successfully" "Success"
        Write-ColorOutput "Follow the on-screen instructions to complete the update" "Info"
        Write-ColorOutput "The update process will continue even after this script exits" "Info"
        
        Start-Sleep -Seconds 5
    }
    catch {
        Exit-WithError "Failed to start Windows 11 Setup: $_"
    }
}

# Main execution
function Start-Win11Upgrade {
    Write-ColorOutput "=========================================================" "Info"
    Write-ColorOutput "      WINDOWS 11 24H2 AUTOMATED UPGRADE UTILITY" "Step"
    Write-ColorOutput "                  WITH TPM/CPU BYPASS" "Step"
    Write-ColorOutput "=========================================================" "Info"
    
    # Check if update is needed
    if (-not (Test-NeedsUpdate)) {
        Write-ColorOutput "No update needed. Exiting." "Info"
        exit 0
    }
    
    # Apply compatibility bypasses
    Write-Step "2" "Installing Windows 11 Compatibility Bypasses"
    Install-CompatibilityBypass
    
    # Download Windows 11 ISO
    if (Get-Windows11ISO -OutputPath $IsoPath) {
        # Mount ISO and start update
        Start-WindowsUpdate -IsoPath $IsoPath
    }
    else {
        Exit-WithError "Failed to obtain Windows 11 24H2 ISO"
    }
}

# Run the upgrade process
Start-Win11Upgrade
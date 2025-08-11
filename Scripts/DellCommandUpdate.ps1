<#
.SYNOPSIS
    Checks and installs Dell drivers, firmware, and BIOS using Dell Command Update CLI.
    Excludes all Dell software except Dell Command Update.

.NOTES
    Author: Earl Daniels
    Date:   2025-08-11
    Tested on: Windows 10/11, Dell laptops/desktops
#>

# --- CONFIG ---
$LogDir = "C:\Logs\Dell"
$LogFile = Join-Path $LogDir ("DellUpdate_" + (Get-Date -Format "yyyy-MM-dd_HH-mm-ss") + ".log")
$DCUPath = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
$DCUInstallerURL = "https://dl.dell.com/FOLDER13309588M/2/Dell-Command-Update-Windows-Universal-Application_C8JXV_WIN64_5.5.0_A00_01.EXE"
$TempInstaller = "$env:TEMP\DellCommandUpdate.exe"

# Create log directory
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory | Out-Null }

# Function to write log
function Write-Log {
    param([string]$Message)
    $time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$time] $Message"
    Write-Output $entry
    Add-Content -Path $LogFile -Value $entry
}

# Function to get installed DCU version
function Get-DCUVersion {
    try {
        $ver = (Get-Item $DCUPath -ErrorAction Stop).VersionInfo.ProductVersion
        return $ver
    }
    catch { return $null }
}

# --- Step 1: Check if DCU-CLI is installed and up-to-date ---
$InstalledVersion = Get-DCUVersion
$LatestVersion = "5.5.0"

if (-not $InstalledVersion) {
    Write-Log "Dell Command Update CLI not found. Installing version $LatestVersion..."
    Invoke-WebRequest -Uri $DCUInstallerURL -OutFile $TempInstaller -UseBasicParsing
    Start-Process -FilePath $TempInstaller -ArgumentList "/S" -Wait
    Remove-Item $TempInstaller -Force
}
elseif ($InstalledVersion -lt $LatestVersion) {
    Write-Log "Older Dell Command Update CLI version ($InstalledVersion) detected. Updating to $LatestVersion..."
    Invoke-WebRequest -Uri $DCUInstallerURL -OutFile $TempInstaller -UseBasicParsing
    Start-Process -FilePath $TempInstaller -ArgumentList "/S" -Wait
    Remove-Item $TempInstaller -Force
}
else {
    Write-Log "Dell Command Update CLI is up-to-date ($InstalledVersion)."
}

# Verify installation
if (-not (Test-Path $DCUPath)) {
    Write-Log "ERROR: Dell Command Update CLI could not be installed or located."
    exit 1
}

# --- Step 2: Scan for drivers, firmware, and BIOS updates ---
try {
    Write-Log "Starting scan for drivers, firmware, and BIOS..."
    Start-Process -FilePath $DCUPath `
        -ArgumentList "/scan -updateType=driver,firmware,BIOS -silent -noreboot" `
        -Wait -NoNewWindow -RedirectStandardOutput $LogFile -RedirectStandardError $LogFile
    Write-Log "Scan completed."
}
catch {
    Write-Log "ERROR: Scan failed. $_"
    exit 1
}

# --- Step 3: Apply updates ---
try {
    Write-Log "Applying updates..."
    Start-Process -FilePath $DCUPath `
        -ArgumentList "/applyUpdates -updateType=driver,firmware,BIOS -silent" `
        -Wait -NoNewWindow -RedirectStandardOutput $LogFile -RedirectStandardError $LogFile
    Write-Log "Updates applied successfully."
}
catch {
    Write-Log "ERROR: Failed to apply updates. $_"
    exit 1
}

# --- Step 4: Handle BIOS reboot if required ---
try {
    Write-Log "Checking if reboot is required..."
    Start-Process -FilePath $DCUPath `
        -ArgumentList "/rebootRequired" `
        -Wait -NoNewWindow -RedirectStandardOutput $LogFile -RedirectStandardError $LogFile

    $RebootPending = Select-String -Path $LogFile -Pattern "Reboot is required"
    if ($RebootPending) {
        Write-Log "BIOS or firmware update requires reboot."
        exit 1
    }
    else {
        Write-Log "No reboot required."
    }
}
catch {
    Write-Log "ERROR: Failed while checking or performing reboot. $_"
}

Write-Log "Script execution complete."

<#
.SYNOPSIS
    Checks and installs Dell drivers, firmware, and BIOS using Dell Command Update CLI.
    Excludes all Dell software except Dell Command Update.

.NOTES
    Author: Earl Daniels
    Date:   2025-08-11
    Tested on: Windows 10/11, Dell laptops/desktops
#>


<#
.SYNOPSIS
    Dell Command Update installer/updater and driver+firmware updater.
.DESCRIPTION
    Checks installed DCU version, compares with latest available version, downloads if newer, 
    installs silently, then runs DCU-CLI to update drivers and BIOS only.
.NOTES
    Requires PowerShell 5.1+ and admin rights.
#>

Function DellCommandUpdate {
$ErrorActionPreference = 'Stop'
$LogFile = "$env:ProgramData\Dell\DCU_Update_Log.txt"

Function Write-Log {
    param([string]$Message)
    $TimeStamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $LogFile -Value "$TimeStamp : $Message"
}

Function Get-InstalledDCUVersion {
    try {
        $dcuPath = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
        if (Test-Path $dcuPath) {
            $versionOutput = & $dcuPath /version 2>&1
            if ($versionOutput -match "(\d+\.\d+\.\d+)") {
                return $Matches[1]
            }
        }
        return $null
    }
    catch {
        Write-Log "Error checking installed DCU version: $_"
        return $null
    }
}

Function Get-LatestDCUVersion {
    # Static link provided
    return "5.5.0"
}

Function Download-DCU {
    param(
        [string]$DownloadUrl,
        [string]$Destination
    )
    $headers = @{
        "Accept"="text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
        "Accept-Encoding"="gzip, deflate, br, zstd"
        "Accept-Language"="en-US,en;q=0.5"
        "Connection"="keep-alive"
        "Cookie"="eSupId=SID=aa6649b4-80fb-4608-a7e1-6f3e1e56b3b5; ..."
        "Host"="dl.dell.com"
        "Priority"="u=0, i"
        "Referer"="https://www.dell.com/support/home/en-us/drivers/driversdetails?driverid=C8JXV"
        "Sec-Fetch-Dest"="document"
        "Sec-Fetch-Mode"="navigate"
        "Sec-Fetch-Site"="same-site"
        "Sec-Fetch-User"="?1"
        "Sec-GPC"="1"
        "Upgrade-Insecure-Requests"="1"
        "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:141.0) Gecko/20100101 Firefox/141.0"
    }
    try {
        Write-Log "Downloading Dell Command Update..."
        Invoke-WebRequest -Uri $DownloadUrl -Headers $headers -OutFile $Destination
        Write-Log "Downloaded to $Destination"
    }
    catch {
        Write-Log "Download failed: $_"
        throw
    }
}

Function Install-DCU {
    param([string]$InstallerPath)
    try {
        Write-Log "Installing Dell Command Update..."
        Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait -NoNewWindow
        Write-Log "Installation completed."
    }
    catch {
        Write-Log "Installation failed: $_"
        throw
    }
}

Function Update-DriversAndFirmware {
    $dcuCLI = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
    if (-not (Test-Path $dcuCLI)) {
        Write-Log "DCU CLI not found after installation."
        throw "DCU CLI missing"
    }
    try {
        Write-Log "Updating drivers and firmware (excluding software)..."
        & $dcuCLI /applyUpdates -updateType=driver,firmware -silent
        Write-Log "Driver and firmware updates applied."
    }
    catch {
        Write-Log "Update process failed: $_"
        throw
    }
}

# === MAIN ===
Write-Log "=== Dell Command Update Script Started ==="

$installedVersion = Get-InstalledDCUVersion
$latestVersion = Get-LatestDCUVersion

Write-Log "Installed DCU Version: $installedVersion"
Write-Log "Latest DCU Version: $latestVersion"

if (-not $installedVersion -or ($installedVersion -ne $latestVersion)) {
    $downloadPath = "$env:TEMP\DellCommandUpdate.exe"
    Download-DCU -DownloadUrl "https://dl.dell.com/FOLDER13309588M/2/Dell-Command-Update-Windows-Universal-Application_C8JXV_WIN64_${latestVersion}_A00_01.EXE" -Destination $downloadPath
    Install-DCU -InstallerPath $downloadPath
}
else {
    Write-Log "DCU is up to date."
}

Update-DriversAndFirmware

Write-Log "=== Update Successful ==="

}


Function DellCommandDriverInstall {
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

<#
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
#>


# --- Step 1: Scan for drivers, firmware, and BIOS updates ---
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

# --- Step 2: Apply updates ---
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

# --- Step 3: Handle BIOS reboot if required ---
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

}

DellCommandUpdate
Start-Sleep 5
DellCommandDriverInstall




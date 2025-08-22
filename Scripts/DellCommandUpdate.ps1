<#
.SYNOPSIS
    Installs/updates Dell Command Update and applies driver/firmware/BIOS updates.
    Excludes Dell software applications, focusing only on essential system updates.

.DESCRIPTION
    - Validates Dell system compatibility
    - Downloads and installs latest Dell Command Update if needed
    - Scans and applies driver, firmware, and BIOS updates
    - Provides comprehensive logging and error handling
    - Handles reboot requirements automatically

.NOTES
    Author: Earl Daniels
    Date: 2025-08-11
    Version: 2.1
    Requires: PowerShell 5.1+, Administrator privileges
    Tested: Windows 10/11, Dell systems
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$Force,
    [switch]$NoReboot,
    [string]$LogPath = "$env:ProgramData\Dell\DCU_Update.log"
)

# Configuration
$Script:Config = @{
    DCUPath     = "C:\Program Files\Dell\CommandUpdate\dcu-cli.exe"
    DCUPathAlt  = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
    DCUVersion  = "5.5.0"
    DCUUrl      = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Assets/Dell-Command-Update-Windows-Universal-Application_C8JXV_WIN64_5.5.0_A00_01.EXE"
    TempPath    = Join-Path $env:TEMP "DellCommandUpdate.exe"
    LogPath     = $LogPath
    UpdateTypes = "driver,firmware,BIOS"
}

# Logging functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARN"    { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }

    Add-Content -Path $Script:Config.LogPath -Value $logEntry -Encoding UTF8
}

function Initialize-Logging {
    $logDir = Split-Path $Script:Config.LogPath -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    Write-Log "=== Dell Command Update Script v2.1 Started ===" -Level "INFO"
}

# System validation
function Test-DellSystem {
    try {
        $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).Manufacturer
        if ($manufacturer -notmatch 'Dell') {
            Write-Log "Non-Dell system detected: $manufacturer" -Level "ERROR"
            return $false
        }
        Write-Log "Dell system confirmed: $manufacturer" -Level "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to detect system manufacturer: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# DCU path/version checks
function Get-DCUExecutablePath {
    foreach ($path in @($Script:Config.DCUPath, $Script:Config.DCUPathAlt)) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

function Get-InstalledDCUVersion {
    try {
        $dcuPath = Get-DCUExecutablePath
        if (-not $dcuPath) { return $null }
        $versionInfo = & $dcuPath /version 2>&1
        if ($versionInfo -match "(\d+\.\d+\.\d+)") {
            return [version]$Matches[1]
        }
        return $null
    }
    catch {
        Write-Log "Error retrieving DCU version: $($_.Exception.Message)" -Level "WARN"
        return $null
    }
}

# Install/update DCU
function Install-DellCommandUpdate {
    try {
        Write-Log "Downloading Dell Command Update v$($Script:Config.DCUVersion)..."
        $retryCount = 0
        $maxRetries = 3
        $delay = 5

        do {
            try {
                Invoke-WebRequest -Uri $Script:Config.DCUUrl -OutFile $Script:Config.TempPath -UseBasicParsing
                break
            }
            catch {
                $retryCount++
                if ($retryCount -eq $maxRetries) {
                    throw "Download failed after $maxRetries attempts: $($_.Exception.Message)"
                }
                Write-Log "Download attempt $retryCount failed, retrying in $delay seconds..." -Level "WARN"
                Start-Sleep -Seconds $delay
                $delay *= 2
            }
        } while ($retryCount -lt $maxRetries)

        Write-Log "Installing Dell Command Update..."
        $installProcess = Start-Process -FilePath $Script:Config.TempPath -ArgumentList "/S" -Wait -PassThru -NoNewWindow
        if ($installProcess.ExitCode -ne 0) {
            throw "Installation failed with exit code: $($installProcess.ExitCode)"
        }

        Remove-Item $Script:Config.TempPath -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 5
        $newPath = Get-DCUExecutablePath
        if (-not $newPath) {
            throw "DCU executable not found after installation"
        }
        Write-Log "Dell Command Update installed successfully" -Level "SUCCESS"
        return $newPath
    }
    catch {
        Write-Log "Installation failed: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Run DCU commands
function Invoke-DCUCommand {
    param(
        [string]$Arguments,
        [string]$Description
    )
    $dcuPath = Get-DCUExecutablePath
    if (-not $dcuPath) { throw "Dell Command Update CLI not found" }

    try {
        Write-Log "Executing: $Description"
        $tempOut = Join-Path $env:TEMP "DCU_Output.log"
        $process = Start-Process -FilePath $dcuPath -ArgumentList $Arguments `
            -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $tempOut `
            -RedirectStandardError $tempOut

        Get-Content $tempOut | ForEach-Object { Write-Log $_ -Level "INFO" }
        Remove-Item $tempOut -Force -ErrorAction SilentlyContinue

        if ($process.ExitCode -eq 0) {
            Write-Log "$Description completed successfully" -Level "SUCCESS"
            return $true
        } else {
            Write-Log "$Description failed with exit code: $($process.ExitCode)" -Level "WARN"
            return $false
        }
    }
    catch {
        Write-Log "$Description error: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

# Check reboot status
function Test-RebootRequired {
    try {
        $tempFile = Join-Path $env:TEMP "temp_reboot_check.txt"
        $dcuPath = Get-DCUExecutablePath
        Start-Process -FilePath $dcuPath -ArgumentList "/rebootRequired" `
            -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $tempFile `
            -RedirectStandardError "nul"

        if (Test-Path $tempFile) {
            $output = Get-Content $tempFile -Raw
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $output -match "Reboot is required"
        }
        return $false
    }
    catch {
        Write-Log "Failed to check reboot status: $($_.Exception.Message)" -Level "WARN"
        return $false
    }
}

# Update process
function Start-UpdateProcess {
    $scanResult = Invoke-DCUCommand -Arguments "/scan -updateType=$($Script:Config.UpdateTypes) -silent -noreboot" -Description "Scanning for updates"
    if (-not $scanResult) {
        Write-Log "Update scan failed or no updates available" -Level "WARN"
        return $false
    }

    $applyArgs = if ($NoReboot) { 
        "/applyUpdates -updateType=$($Script:Config.UpdateTypes) -silent -noreboot"
    } else {
        "/applyUpdates -updateType=$($Script:Config.UpdateTypes) -silent"
    }

    $applyResult = Invoke-DCUCommand -Arguments $applyArgs -Description "Applying updates"
    if (-not $applyResult) {
        Write-Log "Failed to apply some or all updates" -Level "ERROR"
        return $false
    }

    if (-not $NoReboot -and (Test-RebootRequired)) {
        Write-Log "System reboot is required to complete updates" -Level "WARN"
        Write-Log "Please reboot the system manually or run with -NoReboot to skip reboot check" -Level "INFO"
        return $true
    }
    return $true
}

# Main execution
$startTime = Get-Date
try {
    Initialize-Logging
    if (-not (Test-DellSystem)) { return }

    $installedVersion = Get-InstalledDCUVersion
    $requiredVersion = [version]$Script:Config.DCUVersion

    if (-not $installedVersion) {
        Write-Log "Dell Command Update not installed"
        Install-DellCommandUpdate
    }
    elseif ($installedVersion -lt $requiredVersion -or $Force) {
        Write-Log "Updating DCU from v$installedVersion to v$requiredVersion"
        Install-DellCommandUpdate
    }
    else {
        Write-Log "Dell Command Update v$installedVersion is current" -Level "SUCCESS"
    }

    $updateSuccess = Start-UpdateProcess
    $elapsed = (Get-Date) - $startTime
    if ($updateSuccess) {
        Write-Log "=== Dell update process completed successfully in $($elapsed.ToString()) ===" -Level "SUCCESS"
    } else {
        Write-Log "=== Dell update process completed with errors in $($elapsed.ToString()) ===" -Level "ERROR"
    }
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "=== Script execution failed ===" -Level "ERROR"
    exit 1
}
finally {
    if (Test-Path $Script:Config.TempPath) {
        Remove-Item $Script:Config.TempPath -Force -ErrorAction SilentlyContinue
    }
}

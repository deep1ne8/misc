#Requires -Version 5.1
<#
.SYNOPSIS
    Downloads and installs exacqVision Client 25.1.4.0
.DESCRIPTION
    Robust deployment script for exacqVision Client with error handling and logging
#>

[CmdletBinding()]
param()

# Configuration
$AppName = "exacqVision Client"
$Version = "25.1.4.0"
#$DownloadURL = "https://immystrg01016.blob.core.windows.net/software/4c7d8a53-572d-a54a-1594-412e26c2ae34/exacqVisionClient_25.1.4.0_x64.msi?sv=2025-05-05&se=2025-11-15T01%3A47%3A59Z&sr=b&sp=r&sig=d%2BQ79MWYd0qp5HsN4sH%2BWodE5DCet3VC%2FtKu8bCWKkQ%3D"
$DownloadURL = "https://immystrg01016.blob.core.windows.net/software/4c7d8a53-572d-a54a-1594-412e26c2ae34/exacqVisionClient_25.1.4.0_x64.msi?sv=2025-05-05&se=2025-11-17T19%3A18%3A13Z&sr=b&sp=r&sig=U2uq6gwSuA38pvcC%2BZL40XPI7tRbEIaMs6VhqQNJUKE%3D"
$TempPath = "$env:TEMP\exacqVisionClient_Install"
$InstallerPath = "$TempPath\exacqVisionClient.msi"
$LogPath = "$TempPath\Install.log"

# Functions
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Write-Host $LogMessage
    Add-Content -Path $LogPath -Value $LogMessage -ErrorAction SilentlyContinue
}

function Test-IsElevated {
    return ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Main Script
try {
    # Verify elevation
    if (-not (Test-IsElevated)) {
        throw "This script requires administrative privileges. Please run as Administrator."
    }

    # Create temp directory
    if (-not (Test-Path $TempPath)) {
        New-Item -Path $TempPath -ItemType Directory -Force | Out-Null
        Write-Log "Created temporary directory: $TempPath"
    }

    # Download installer
    Write-Log "Downloading $AppName $Version..."
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $DownloadURL -OutFile $InstallerPath -UseBasicParsing -ErrorAction Stop
    $ProgressPreference = 'Continue'
    
    # Verify download
    if (-not (Test-Path $InstallerPath)) {
        throw "Download failed. Installer not found at $InstallerPath"
    }
    
    $FileSize = (Get-Item $InstallerPath).Length / 1MB
    Write-Log "Download complete. File size: $([math]::Round($FileSize, 2)) MB"

    # Install MSI
    Write-Log "Installing $AppName..."
    $MSIArgs = @(
        "/i"
        "`"$InstallerPath`""
        "/qn"
        "/norestart"
        "/L*v"
        "`"$LogPath`""
    )
    
    $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList $MSIArgs -Wait -PassThru -NoNewWindow
    
    if ($Process.ExitCode -eq 0) {
        Write-Log "Installation completed successfully" -Level "SUCCESS"
    } elseif ($Process.ExitCode -eq 3010) {
        Write-Log "Installation completed. Reboot required (Exit Code: 3010)" -Level "WARNING"
    } else {
        throw "Installation failed with exit code: $($Process.ExitCode). Check log: $LogPath"
    }

} catch {
    Write-Log "ERROR: $($_.Exception.Message)" -Level "ERROR"
    exit 1
} finally {
    # Cleanup
    if (Test-Path $InstallerPath) {
        Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
        Write-Log "Cleaned up installer file"
    }
}

exit 0

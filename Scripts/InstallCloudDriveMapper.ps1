#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Removes all Cloud Drive Mapper versions and performs clean installation with optimal drive letter configuration.
.DESCRIPTION
    Comprehensive script that uninstalls all Cloud Drive Mapper instances, removes registry remnants,
    and installs the latest version with best practice drive letter assignment (Z -> Y -> X, etc.).
.PARAMETER InstallerPath
    Path to the CloudDriveMapper.msi installer file.
.PARAMETER LicenseKey
    License key for Cloud Drive Mapper activation.
.EXAMPLE
    .\Deploy-CloudDriveMapper.ps1 -InstallerPath "\\server\share\CloudDriveMapper.msi" -LicenseKey "YOUR-LICENSE-KEY"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$InstallerPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LicenseKey
)

$ErrorActionPreference = 'Stop'
$LogPath = "$env:ProgramData\CDM_Deployment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $LogMessage = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogPath -Value $LogMessage
    switch ($Level) {
        'ERROR' { Write-Error $Message }
        'WARN'  { Write-Warning $Message }
        default { Write-Host $Message -ForegroundColor Cyan }
    }
}

function Get-InstalledCDM {
    Write-Log "Scanning for installed Cloud Drive Mapper instances..."
    $UninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    
    Get-ItemProperty $UninstallKeys -ErrorAction SilentlyContinue | 
        Where-Object { $_.DisplayName -like "*Cloud Drive Mapper*" -or $_.DisplayName -like "*CloudDriveMapper*" } |
        Select-Object DisplayName, DisplayVersion, UninstallString, PSChildName
}

function Remove-CDMInstances {
    $Instances = Get-InstalledCDM
    
    if (-not $Instances) {
        Write-Log "No existing Cloud Drive Mapper installations found."
        return
    }

    foreach ($Instance in $Instances) {
        Write-Log "Removing: $($Instance.DisplayName) $($Instance.DisplayVersion)"
        
        if ($Instance.PSChildName -match '^{[A-F0-9-]+}$') {
            $ProductCode = $Instance.PSChildName
            $Arguments = "/x `"$ProductCode`" /qn /norestart"
            
            try {
                $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
                if ($Process.ExitCode -eq 0 -or $Process.ExitCode -eq 3010) {
                    Write-Log "Successfully uninstalled: $($Instance.DisplayName)"
                } else {
                    Write-Log "Uninstall completed with exit code: $($Process.ExitCode)" -Level 'WARN'
                }
            } catch {
                Write-Log "Failed to uninstall $($Instance.DisplayName): $_" -Level 'ERROR'
            }
        }
    }
    
    Start-Sleep -Seconds 3
}

function Remove-CDMRegistry {
    Write-Log "Cleaning registry remnants..."
    
    $RegistryPaths = @(
        'HKLM:\SOFTWARE\CloudDriveMapper',
        'HKLM:\SOFTWARE\WOW6432Node\CloudDriveMapper',
        'HKCU:\SOFTWARE\CloudDriveMapper',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    )

    foreach ($Path in $RegistryPaths) {
        if (Test-Path $Path) {
            try {
                if ($Path -like "*\Run") {
                    $RunKeys = Get-ItemProperty $Path -ErrorAction SilentlyContinue
                    $RunKeys.PSObject.Properties | Where-Object { $_.Name -like "*CloudDrive*" } | ForEach-Object {
                        Remove-ItemProperty -Path $Path -Name $_.Name -Force -ErrorAction SilentlyContinue
                        Write-Log "Removed Run key: $($_.Name)"
                    }
                } else {
                    Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                    Write-Log "Removed registry path: $Path"
                }
            } catch {
                Write-Log "Could not remove $Path : $_" -Level 'WARN'
            }
        }
    }
}

function Remove-CDMFiles {
    Write-Log "Removing residual files..."
    
    $Paths = @(
        "$env:ProgramFiles\CloudDriveMapper",
        "$env:ProgramFiles\Cloud Drive Mapper",
        "${env:ProgramFiles(x86)}\CloudDriveMapper",
        "${env:ProgramFiles(x86)}\Cloud Drive Mapper",
        "$env:ProgramData\CloudDriveMapper",
        "$env:LOCALAPPDATA\CloudDriveMapper",
        "$env:APPDATA\CloudDriveMapper"
    )

    foreach ($Path in $Paths) {
        if (Test-Path $Path) {
            try {
                Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
                Write-Log "Removed directory: $Path"
            } catch {
                Write-Log "Could not remove $Path : $_" -Level 'WARN'
            }
        }
    }
}

function Get-OptimalDriveLetter {
    Write-Log "Determining optimal drive letter..."
    
    # Best practice: Z, Y, X, W, V in descending order
    $PreferredLetters = @('Z', 'Y', 'X', 'W', 'V')
    $UsedDrives = (Get-PSDrive -PSProvider FileSystem).Name
    
    foreach ($Letter in $PreferredLetters) {
        if ($Letter -notin $UsedDrives) {
            Write-Log "Selected drive letter: $Letter"
            return $Letter
        }
    }
    
    # Fallback: Find any available letter from T-Z
    for ([int]$i = 90; $i -ge 84; $i--) {
        $Letter = [char]$i
        if ($Letter -notin $UsedDrives) {
            Write-Log "Fallback drive letter selected: $Letter"
            return $Letter
        }
    }
    
    Write-Log "No suitable drive letters available!" -Level 'ERROR'
    throw "Unable to find available drive letter for Cloud Drive Mapper."
}

function Install-CDM {
    param([string]$DriveLetter)
    
    Write-Log "Installing Cloud Drive Mapper with drive letter: $DriveLetter"
    
    $Arguments = @(
        "/i `"$InstallerPath`""
        "/qn"
        "/norestart"
        "LICENSEKEY=$LicenseKey"
        "DRIVELETTER=$DriveLetter"
        "DESKTOP_SHORTCUT=0"
        "STARTMENU_SHORTCUT=1"
        "STARTUP_SHORTCUT=1"
        "LAUNCHCDM=1"
        "INSTALLCDMLEGACY=1"
        "/l*v `"$env:ProgramData\CDM_Install.log`""
    )
    
    try {
        $Process = Start-Process -FilePath "msiexec.exe" -ArgumentList ($Arguments -join ' ') -Wait -PassThru -NoNewWindow
        
        if ($Process.ExitCode -eq 0) {
            Write-Log "Cloud Drive Mapper installed successfully."
            return $true
        } elseif ($Process.ExitCode -eq 3010) {
            Write-Log "Installation successful. Reboot required." -Level 'WARN'
            return $true
        } else {
            Write-Log "Installation failed with exit code: $($Process.ExitCode). Check log: $env:ProgramData\CDM_Install.log" -Level 'ERROR'
            return $false
        }
    } catch {
        Write-Log "Installation exception: $_" -Level 'ERROR'
        return $false
    }
}

function Verify-Installation {
    Write-Log "Verifying installation..."
    Start-Sleep -Seconds 5
    
    $Installed = Get-InstalledCDM
    if ($Installed) {
        Write-Log "Verification successful: $($Installed.DisplayName) $($Installed.DisplayVersion)"
        return $true
    } else {
        Write-Log "Verification failed: Cloud Drive Mapper not detected in registry." -Level 'ERROR'
        return $false
    }
}

# ============================================================================
# MAIN EXECUTION
# ============================================================================

try {
    Write-Log "========== Cloud Drive Mapper Deployment Started =========="
    Write-Log "Installer: $InstallerPath"
    Write-Log "Log file: $LogPath"
    
    # Phase 1: Complete removal
    Write-Log "Phase 1: Removing existing installations..."
    Remove-CDMInstances
    Remove-CDMRegistry
    Remove-CDMFiles
    
    # Phase 2: Drive letter determination
    Write-Log "Phase 2: Drive letter configuration..."
    $DriveLetter = Get-OptimalDriveLetter
    
    # Phase 3: Installation
    Write-Log "Phase 3: Installing Cloud Drive Mapper..."
    $InstallSuccess = Install-CDM -DriveLetter $DriveLetter
    
    if (-not $InstallSuccess) {
        throw "Installation failed. Check logs for details."
    }
    
    # Phase 4: Verification
    Write-Log "Phase 4: Post-installation verification..."
    $VerifySuccess = Verify-Installation
    
    if ($VerifySuccess) {
        Write-Log "========== Deployment Completed Successfully ==========" -Level 'INFO'
        Write-Log "Log saved to: $LogPath"
        exit 0
    } else {
        throw "Installation verification failed."
    }
    
} catch {
    Write-Log "========== Deployment Failed ==========" -Level 'ERROR'
    Write-Log "Error: $_" -Level 'ERROR'
    Write-Log "Log saved to: $LogPath"
    exit 1
}

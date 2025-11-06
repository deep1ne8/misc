#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Reconfigures Cloud Drive Mapper by finding the installed version, cleaning up, and reinstalling with optimal configuration.
.DESCRIPTION
    Locates the existing Cloud Drive Mapper installation, extracts the MSI from Windows Installer cache,
    performs complete cleanup, and reinstalls with best practice drive letter assignment (Z -> Y -> X, etc.).
.PARAMETER LicenseKey
    License key for Cloud Drive Mapper activation.
.EXAMPLE
    .\Reconfigure-CloudDriveMapper.ps1 -LicenseKey "5d9a0dba906345e24b3382926c2557fc7abb44841d28d055bd2d"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LicenseKey
)

$ErrorActionPreference = 'Continue'
$LogPath = "$env:ProgramData\CDM_Deployment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if (-not(Test-Path $LogPath)){
Write-Host "Creating Log Directory..."
New-Item -Path "$env:ProgramData" -ItemType "File" -Name "CDM_Deployment_$(Get-Date -Format 'yyyyMMdd_HHmmss').log" -Verbose -Force
Start-Sleep 3
Write-Host "Log File Created...."
}

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

function Get-CDMInstallerPath {
    Write-Log "Locating existing Cloud Drive Mapper installer..."
    
    $CDMInstallation = Get-InstalledCDM | Select-Object -First 1
    
    if (-not $CDMInstallation) {
        Write-Log "No Cloud Drive Mapper installation found on this machine!" -Level 'ERROR'
        throw "Cloud Drive Mapper must be installed before reconfiguration. Please install it first."
    }
    
    Write-Log "Found installation: $($CDMInstallation.DisplayName) $($CDMInstallation.DisplayVersion)"
    
    # Extract product code from PSChildName (GUID)
    $ProductCode = $CDMInstallation.PSChildName
    
    if ($ProductCode -notmatch '^{[A-F0-9-]+}$') {
        Write-Log "Invalid product code format: $ProductCode" -Level 'ERROR'
        throw "Unable to extract valid product code from installation."
    }
    
    Write-Log "Product Code: $ProductCode"
    
    # Query Windows Installer for the local package path
    try {
        $InstallerType = [Type]::GetTypeFromProgID("WindowsInstaller.Installer")
        $Installer = [Activator]::CreateInstance($InstallerType)
        $LocalPackage = $Installer.GetType().InvokeMember("ProductInfo", "GetProperty", $null, $Installer, @($ProductCode, "LocalPackage"))
        
        if ($LocalPackage -and (Test-Path $LocalPackage)) {
            Write-Log "Found cached MSI: $LocalPackage"
            
            # Copy to a working location to ensure we can use it
            $BackupPath = "$env:TEMP\CloudDriveMapper_Backup_$(Get-Date -Format 'yyyyMMddHHmmss').msi"
            Copy-Item -Path $LocalPackage -Destination $BackupPath -Force
            Write-Log "Backed up installer to: $BackupPath"
            
            return $BackupPath
        }
    } catch {
        Write-Log "COM method failed: $_" -Level 'WARN'
    }
    
    # Fallback: Search Windows Installer cache directory
    Write-Log "Attempting fallback search in Windows Installer cache..."
    $InstallerCache = "$env:SystemRoot\Installer"
    
    if (Test-Path $InstallerCache) {
        # Get all MSI files and search for Cloud Drive Mapper
        $MSIFiles = Get-ChildItem -Path $InstallerCache -Filter "*.msi" -File -ErrorAction SilentlyContinue
        
        foreach ($MSI in $MSIFiles) {
            try {
                $InstallerType = [Type]::GetTypeFromProgID("WindowsInstaller.Installer")
                $Installer = [Activator]::CreateInstance($InstallerType)
                $Database = $Installer.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $Installer, @($MSI.FullName, 0))
                
                $View = $Database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $Database, @("SELECT Value FROM Property WHERE Property='ProductName'"))
                $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
                $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
                
                if ($Record) {
                    $ProductName = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, @(1))
                    
                    if ($ProductName -like "*Cloud Drive Mapper*" -or $ProductName -like "*CloudDriveMapper*") {
                        Write-Log "Found matching MSI: $($MSI.FullName) - Product: $ProductName"
                        
                        $BackupPath = "$env:TEMP\CloudDriveMapper_Backup_$(Get-Date -Format 'yyyyMMddHHmmss').msi"
                        Copy-Item -Path $MSI.FullName -Destination $BackupPath -Force
                        Write-Log "Backed up installer to: $BackupPath"
                        
                        return $BackupPath
                    }
                }
            } catch {
                # Continue searching
            }
        }
    }
    
    Write-Log "Unable to locate Cloud Drive Mapper installer in cache!" -Level 'ERROR'
    throw "Could not find the original installer. Please provide the MSI file manually."
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
    param(
        [string]$InstallerPath,
        [string]$DriveLetter
    )
    
    Write-Log "Installing Cloud Drive Mapper with drive letter: $DriveLetter"
    Write-Log "Using installer: $InstallerPath"
    
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
    Write-Log "========== Cloud Drive Mapper Reconfiguration Started =========="
    Write-Log "Log file: $LogPath"
    
    # Phase 0: Locate existing installer
    Write-Log "Phase 0: Locating existing Cloud Drive Mapper installer..."
    $InstallerPath = Get-CDMInstallerPath
    
    if (-not $InstallerPath -or -not (Test-Path $InstallerPath)) {
        throw "Failed to locate valid installer path."
    }
    
    # Phase 1: Complete removal
    Write-Log "Phase 1: Removing existing installation..."
    Remove-CDMInstances
    Remove-CDMRegistry
    Remove-CDMFiles
    
    Write-Log "Waiting for cleanup to complete..."
    Start-Sleep -Seconds 5
    
    # Phase 2: Drive letter determination
    Write-Log "Phase 2: Drive letter configuration..."
    $DriveLetter = Get-OptimalDriveLetter
    
    # Phase 3: Installation
    Write-Log "Phase 3: Reinstalling Cloud Drive Mapper..."
    $InstallSuccess = Install-CDM -InstallerPath $InstallerPath -DriveLetter $DriveLetter
    
    if (-not $InstallSuccess) {
        throw "Installation failed. Check logs for details."
    }
    
    # Phase 4: Verification
    Write-Log "Phase 4: Post-installation verification..."
    $VerifySuccess = Verify-Installation
    
    if ($VerifySuccess) {
        Write-Log "========== Reconfiguration Completed Successfully ==========" -Level 'INFO'
        Write-Log "Drive Letter: $DriveLetter"
        Write-Log "Log saved to: $LogPath"
        
        # Cleanup temp installer backup
        if (Test-Path $InstallerPath) {
            Remove-Item -Path $InstallerPath -Force -ErrorAction SilentlyContinue
            Write-Log "Cleaned up temporary installer backup."
        }
        
        exit 0
    } else {
        throw "Installation verification failed."
    }
    
} catch {
    Write-Log "========== Reconfiguration Failed ==========" -Level 'ERROR'
    Write-Log "Error: $_" -Level 'ERROR'
    Write-Log "Log saved to: $LogPath"
    exit 1
}

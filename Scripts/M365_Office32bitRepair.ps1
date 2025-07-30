#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Automated Microsoft Apps for Business 32-bit Deployment Script
.DESCRIPTION
    Detects and removes 64-bit Microsoft Apps for Business, then installs 32-bit version
    with Access, Current Channel, English language
.NOTES
    Version: 1.0
    Requires: Administrator privileges
    Author: PowerShell Automation Script
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$env:TEMP\OfficeDeployment.log",
    [string]$WorkingDir = "$env:TEMP\OfficeDeployment",
    [switch]$Force
)

# Initialize logging
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARN"){"Yellow"}else{"Green"})
    Add-Content -Path $LogPath -Value $logMessage -Force
}

# Create working directory
try {
    if (Test-Path $WorkingDir) { Remove-Item $WorkingDir -Recurse -Force }
    New-Item -ItemType Directory -Path $WorkingDir -Force | Out-Null
    Write-Log "Working directory created: $WorkingDir"
} catch {
    Write-Log "Failed to create working directory: $($_.Exception.Message)" "ERROR"
    exit 1
}

# Detect existing Office installations
function Get-OfficeInstallations {
    Write-Log "Detecting existing Office installations..."
    $installations = @()
    
    # Check Click-to-Run installations
    $c2rPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $c2rPath) {
        $config = Get-ItemProperty -Path $c2rPath -ErrorAction SilentlyContinue
        if ($config) {
            $installations += [PSCustomObject]@{
                Type = "Click-to-Run"
                Version = $config.VersionToReport
                Platform = $config.Platform
                ProductIds = $config.ProductReleaseIds
                ClientFolder = $config.ClientFolder
            }
        }
    }
    
    # Check MSI installations
    $msiProducts = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -match "Microsoft.*Office|Microsoft 365|Microsoft Apps"
    }
    
    foreach ($product in $msiProducts) {
        $installations += [PSCustomObject]@{
            Type = "MSI"
            Name = $product.Name
            Version = $product.Version
            IdentifyingNumber = $product.IdentifyingNumber
        }
    }
    
    return $installations
}

# Remove existing Office installations
function Remove-OfficeInstallations {
    param([array]$Installations)
    
    Write-Log "Removing existing Office installations..."
    
    foreach ($install in $Installations) {
        try {
            if ($install.Type -eq "Click-to-Run") {
                Write-Log "Removing Click-to-Run installation..."
                
                # Download Office Deployment Tool if not exists
                $odtPath = "$WorkingDir\setup.exe"
                if (!(Test-Path $odtPath)) {
                    Write-Log "Downloading Office Deployment Tool..."
                    $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20170.exe"
                    Invoke-WebRequest -Uri $odtUrl -OutFile "$WorkingDir\odt.exe" -UseBasicParsing
                    Start-Process -FilePath "$WorkingDir\odt.exe" -ArgumentList "/quiet", "/extract:$WorkingDir" -Wait
                }
                
                # Create removal configuration
                $removeConfig = @"
<Configuration>
  <Remove All="TRUE" />
</Configuration>
"@
                $removeConfig | Out-File "$WorkingDir\remove.xml" -Encoding UTF8
                
                # Execute removal
                Start-Process -FilePath $odtPath -ArgumentList "/configure", "$WorkingDir\remove.xml" -Wait -NoNewWindow
                Write-Log "Click-to-Run removal completed"
                
            } elseif ($install.Type -eq "MSI") {
                Write-Log "Removing MSI installation: $($install.Name)"
                Start-Process -FilePath "msiexec.exe" -ArgumentList "/x", $install.IdentifyingNumber, "/quiet", "/norestart" -Wait
                Write-Log "MSI removal completed: $($install.Name)"
            }
        } catch {
            Write-Log "Failed to remove installation: $($_.Exception.Message)" "ERROR"
        }
    }
    
    # Clean registry entries
    Write-Log "Cleaning registry entries..."
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*",
        "HKCU:\SOFTWARE\Microsoft\Office"
    )
    
    foreach ($regPath in $regPaths) {
        try {
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-Log "Registry cleanup warning: $($_.Exception.Message)" "WARN"
        }
    }
}

# Install 32-bit Microsoft Apps for Business
function Install-Office32Bit {
    Write-Log "Installing 32-bit Microsoft Apps for Business..."
    
    try {
        # Download Office Deployment Tool if not exists
        $odtPath = "$WorkingDir\setup.exe"
        if (!(Test-Path $odtPath)) {
            Write-Log "Downloading Office Deployment Tool..."
            $odtUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20170.exe"
            Invoke-WebRequest -Uri $odtUrl -OutFile "$WorkingDir\odt.exe" -UseBasicParsing
            Start-Process -FilePath "$WorkingDir\odt.exe" -ArgumentList "/quiet", "/extract:$WorkingDir" -Wait
        }
        
        # Create installation configuration
        $installConfig = @"
<Configuration>
  <Add OfficeClientEdition="32" Channel="Current" MigrateArch="TRUE">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us" />
      <ExcludeApp ID="Groove" />
      <ExcludeApp ID="Lync" />
      <ExcludeApp ID="OneDrive" />
      <ExcludeApp ID="Teams" />
    </Product>
    <Product ID="AccessRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <Setup Name="Company" Value="Organization" />
  </AppSettings>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
        
        $installConfig | Out-File "$WorkingDir\install.xml" -Encoding UTF8
        Write-Log "Installation configuration created"
        
        # Execute installation
        Write-Log "Starting Office installation... This may take several minutes."
        $process = Start-Process -FilePath $odtPath -ArgumentList "/configure", "$WorkingDir\install.xml" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log "Office installation completed successfully"
        } else {
            Write-Log "Office installation failed with exit code: $($process.ExitCode)" "ERROR"
            return $false
        }
        
        return $true
    } catch {
        Write-Log "Installation failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Verify installation
function Test-OfficeInstallation {
    Write-Log "Verifying Office installation..."
    
    $officeApps = @(
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"
    )
    
    $installed = $true
    foreach ($app in $officeApps) {
        if (Test-Path $app) {
            $appName = Split-Path $app -Leaf
            Write-Log "✓ $appName found"
        } else {
            $appName = Split-Path $app -Leaf
            Write-Log "✗ $appName missing" "WARN"
            $installed = $false
        }
    }
    
    # Check architecture
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $regPath) {
        $config = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($config.Platform -eq "x86") {
            Write-Log "✓ Architecture: 32-bit confirmed"
        } else {
            Write-Log "✗ Architecture: $($config.Platform) detected" "WARN"
            $installed = $false
        }
    }
    
    return $installed
}

# Cleanup function
function Remove-WorkingDirectory {
    try {
        if (Test-Path $WorkingDir) {
            Remove-Item $WorkingDir -Recurse -Force
            Write-Log "Working directory cleaned up"
        }
    } catch {
        Write-Log "Cleanup warning: $($_.Exception.Message)" "WARN"
    }
}

# Main execution
try {
    Write-Log "=== Microsoft Apps for Business 32-bit Deployment Started ==="
    Write-Log "Script running with PowerShell version: $($PSVersionTable.PSVersion)"
    
    # Check admin privileges
    if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
        Write-Log "Script requires administrator privileges" "ERROR"
        exit 1
    }
    
    # Stop Office processes
    Write-Log "Stopping Office processes..."
    $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "msaccess", "lync", "communicator")
    foreach ($process in $officeProcesses) {
        Get-Process -Name $process -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    
    # Detect existing installations
    $existingInstalls = Get-OfficeInstallations
    
    if ($existingInstalls.Count -gt 0) {
        Write-Log "Found $($existingInstalls.Count) existing Office installation(s)"
        foreach ($install in $existingInstalls) {
            Write-Log "- Type: $($install.Type), Platform: $($install.Platform)"
        }
        
        # Check if 32-bit is already installed
        $has32Bit = $existingInstalls | Where-Object { $_.Platform -eq "x86" }
        if ($has32Bit -and !$Force) {
            Write-Log "32-bit Office already detected. Use -Force to reinstall." "WARN"
            exit 0
        }
        
        # Remove existing installations
        Remove-OfficeInstallations -Installations $existingInstalls
        
        # Wait for removal to complete
        Write-Log "Waiting for removal to complete..."
        Start-Sleep -Seconds 30
    } else {
        Write-Log "No existing Office installations detected"
    }
    
    # Install 32-bit Office
    $installSuccess = Install-Office32Bit
    
    if ($installSuccess) {
        # Wait for installation to complete
        Write-Log "Waiting for installation to stabilize..."
        Start-Sleep -Seconds 60
        
        # Verify installation
        if (Test-OfficeInstallation) {
            Write-Log "=== Deployment completed successfully ===" "INFO"
            Write-Log "Microsoft Apps for Business 32-bit with Access installed"
        } else {
            Write-Log "=== Deployment completed with warnings ===" "WARN"
            Write-Log "Some components may not have installed correctly"
        }
    } else {
        Write-Log "=== Deployment failed ===" "ERROR"
        exit 1
    }
    
} catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    exit 1
} finally {
    Remove-WorkingDirectory
    Write-Log "Log file saved to: $LogPath"
}

exit 0
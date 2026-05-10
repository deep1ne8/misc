
<#
.SYNOPSIS
    Automated Microsoft Apps for Business 64-bit Deployment Script
.DESCRIPTION
    Detects and removes 32-bit Microsoft Apps for Business, then installs 64-bit version
    with Access, Current Channel, English language
.NOTES
    Version: 1.0
    Requires: Administrator privileges
    Author: PowerShell Automation Script
#>

[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "$env:TEMP\OfficeDeployment64bit.log",
    [ValidateNotNullOrEmpty()]
    [string]$WorkingDir = "$env:TEMP\OfficeDeployment64bit",
    [switch]$Force
)


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
            }#Requires -RunAsAdministrator
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
            InstallLocation = $product.InstallLocation
        }
    }
    
    return $installations
}



# Initialize logging
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARN"){"Yellow"}else{"Green"})
    Add-Content -Path $LogPath -Value $logMessage -Force
}

# Alternative: Get ODT from Microsoft directly
function Get-LatestODT {
    param([string]$DestinationPath)
    
    Write-Log "Attempting to get latest ODT..."
    try {
        # Try direct Microsoft download page scraping
        $downloadPage = Invoke-WebRequest -Uri "https://www.microsoft.com/en-us/download/details.aspx?id=49117" -UseBasicParsing
        $downloadLink = ($downloadPage.Links | Where-Object { $_.href -match "officedeploymenttool.*\.exe" }).href | Select-Object -First 1
        
        if ($downloadLink) {
            Write-Log "Found ODT download link: $downloadLink"
            Invoke-WebRequest -Uri $downloadLink -OutFile "$DestinationPath\odt.exe" -UseBasicParsing -TimeoutSec 60
            return $true
        }
    } catch {
        Write-Log "Failed to get latest ODT: $($_.Exception.Message)" "WARN"
    }
    
    return $false
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
    
    # Check Windows Store/UWP Office installations
    Write-Log "Checking Windows Store Office installations..."
    try {
        $storeApps = Get-AppxPackage -AllUsers | Where-Object {
            $_.Name -match "Microsoft\.Office|Microsoft\.WindowsOffice|Microsoft\.MicrosoftOfficeHub"
        }
        
        foreach ($app in $storeApps) {
            $installations += [PSCustomObject]@{
                Type = "Windows Store"
                Name = $app.Name
                Version = $app.Version
                PackageFullName = $app.PackageFullName
                InstallLocation = $app.InstallLocation
                Architecture = $app.Architecture
            }
            Write-Log "Found Windows Store app: $($app.Name) v$($app.Version) ($($app.Architecture))"
        }
    } catch {
        Write-Log "Error checking Windows Store apps: $($_.Exception.Message)" "WARN"
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
                    $odtUrls = @(
                        "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe",
                        "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20170.exe"
                    )
                    
                    $downloaded = $false
                    foreach ($url in $odtUrls) {
                        try {
                            Write-Log "Attempting download from: $url"
                            Invoke-WebRequest -Uri $url -OutFile "$WorkingDir\odt.exe" -UseBasicParsing -TimeoutSec 30
                            if (Test-Path "$WorkingDir\odt.exe") {
                                $downloaded = $true
                                break
                            }
                        } catch {
                            Write-Log "Download failed from $url : $($_.Exception.Message)" "WARN"
                        }
                    }
                    
                    if (!$downloaded) {
                        Write-Log "All ODT download attempts failed. Please download manually from https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool" "ERROR"
                        throw "Unable to download Office Deployment Tool"
                    }
                    
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
                
            } elseif ($install.Type -eq "Windows Store") {
                Write-Log "Removing Windows Store installation: $($install.Name)"
                try {
                    Remove-AppxPackage -Package $install.PackageFullName -AllUsers -ErrorAction SilentlyContinue
                    Write-Log "Windows Store app removal completed: $($install.Name)"
                } catch {
                    Write-Log "Failed to remove Windows Store app: $($install.Name) - $($_.Exception.Message)" "WARN"
                    # Try alternative removal method
                    try {
                        Get-AppxPackage -AllUsers -Name $install.Name | Remove-AppxPackage -ErrorAction SilentlyContinue
                    } catch {
                        Write-Log "Alternative removal also failed for: $($install.Name)" "WARN"
                    }
                }
                
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
        "HKCU:\SOFTWARE\Microsoft\Office",
        "HKCU:\SOFTWARE\Classes\AppX*"
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
    
    # Remove Windows Store app registrations
    Write-Log "Cleaning Windows Store app registrations..."
    try {
        $storeRegPaths = @(
            "HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages\Microsoft.Office*",
            "HKCU:\SOFTWARE\Classes\Local Settings\Software\Microsoft\Windows\CurrentVersion\AppModel\Repository\Packages\Microsoft.WindowsOffice*"
        )
        
        foreach ($regPath in $storeRegPaths) {
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    } catch {
        Write-Log "Windows Store registry cleanup warning: $($_.Exception.Message)" "WARN"
    }
}

# Install 64-bit Microsoft Apps for Business
function Install-Office64Bit {
    Write-Log "Installing 64-bit Microsoft Apps for Business..."
    
    try {
        # Download Office Deployment Tool if not exists
        $odtPath = "$WorkingDir\setup.exe"
        if (!(Test-Path $odtPath)) {
            Write-Log "Downloading Office Deployment Tool..."
            $odtUrls = @(
                "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe",
                "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20170.exe"
            )
            
            $downloaded = $false
            foreach ($url in $odtUrls) {
                try {
                    Write-Log "Attempting download from: $url"
                    Invoke-WebRequest -Uri $url -OutFile "$WorkingDir\odt.exe" -UseBasicParsing -TimeoutSec 30
                    if (Test-Path "$WorkingDir\odt.exe") {
                        $downloaded = $true
                        break
                    }
                } catch {
                    Write-Log "Download failed from $url : $($_.Exception.Message)" "WARN"
                }
            }
            
            if (!$downloaded) {
                Write-Log "ODT download failed. Attempting alternative method..." "WARN"
                
                # Alternative: Use cached ODT if available
                $cachedODT = "$env:ProgramFiles\Microsoft Office 15\ClientX64\OfficeClickToRun.exe"
                if (Test-Path $cachedODT) {
                    Write-Log "Using cached Office installation tool"
                    Copy-Item $cachedODT "$WorkingDir\setup.exe"
                } else {
                    Write-Log "Please download ODT manually from: https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool" "ERROR"
                    Write-Log "Place setup.exe in: $WorkingDir" "ERROR"
                    
                    # Wait for manual intervention
                    do {
                        Write-Log "Waiting for setup.exe to be placed in working directory..." "WARN"
                        Start-Sleep -Seconds 10
                    } while (!(Test-Path "$WorkingDir\setup.exe"))
                }
            } else {
                Start-Process -FilePath "$WorkingDir\odt.exe" -ArgumentList "/quiet", "/extract:$WorkingDir" -Wait
            }
        }
        
        # Create installation configuration for 64-bit
        $installConfig = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current" MigrateArch="TRUE">
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
        Write-Log "64-bit installation configuration created"
        
        # Execute installation
        Write-Log "Starting 64-bit Office installation... This may take several minutes."
        $process = Start-Process -FilePath $odtPath -ArgumentList "/configure", "$WorkingDir\install.xml" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Log "64-bit Office installation completed successfully"
        } else {
            Write-Log "64-bit Office installation failed with exit code: $($process.ExitCode)" "ERROR"
            return $false
        }
        
        return $true
    } catch {
        Write-Log "Installation failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Manage Office Icons and Shortcuts (64-bit version)
function Set-OfficeIcons {
    Write-Log "Configuring Office icons and shortcuts..."
    
    try {
        # Define Office applications with their 64-bit paths
        $officeApps = @{
            "Word" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE,0"
                "DisplayName" = "Microsoft Word"
            }
            "Excel" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE,0"
                "DisplayName" = "Microsoft Excel"
            }
            "PowerPoint" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE,0"
                "DisplayName" = "Microsoft PowerPoint"
            }
            "Outlook" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE,0"
                "DisplayName" = "Microsoft Outlook"
            }
            "Access" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\MSACCESS.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\MSACCESS.EXE,0"
                "DisplayName" = "Microsoft Access"
            }
            "Publisher" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\MSPUB.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\MSPUB.EXE,0"
                "DisplayName" = "Microsoft Publisher"
            }
            "OneNote" = @{
                "Path" = "$env:ProgramFiles\Microsoft Office\root\Office16\ONENOTE.EXE"
                "IconPath" = "$env:ProgramFiles\Microsoft Office\root\Office16\ONENOTE.EXE,0"
                "DisplayName" = "Microsoft OneNote"
            }
        }
        
        # Create desktop shortcuts
        $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
        $shell = New-Object -ComObject WScript.Shell
        
        foreach ($appName in $officeApps.Keys) {
            $app = $officeApps[$appName]
            if (Test-Path $app.Path) {
                try {
                    $shortcutPath = "$desktopPath\$($app.DisplayName).lnk"
                    $shortcut = $shell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $app.Path
                    $shortcut.IconLocation = $app.IconPath
                    $shortcut.WorkingDirectory = Split-Path $app.Path
                    $shortcut.Description = $app.DisplayName
                    $shortcut.Save()
                    Write-Log "✓ Desktop shortcut created: $($app.DisplayName)"
                } catch {
                    Write-Log "Failed to create desktop shortcut for $appName : $($_.Exception.Message)" "WARN"
                }
            }
        }
        
        # Create Start Menu shortcuts
        $startMenuPath = [Environment]::GetFolderPath("CommonStartMenu") + "\Programs\Microsoft Office"
        if (!(Test-Path $startMenuPath)) {
            New-Item -ItemType Directory -Path $startMenuPath -Force | Out-Null
        }
        
        foreach ($appName in $officeApps.Keys) {
            $app = $officeApps[$appName]
            if (Test-Path $app.Path) {
                try {
                    $shortcutPath = "$startMenuPath\$($app.DisplayName).lnk"
                    $shortcut = $shell.CreateShortcut($shortcutPath)
                    $shortcut.TargetPath = $app.Path
                    $shortcut.IconLocation = $app.IconPath
                    $shortcut.WorkingDirectory = Split-Path $app.Path
                    $shortcut.Description = $app.DisplayName
                    $shortcut.Save()
                    Write-Log "✓ Start Menu shortcut created: $($app.DisplayName)"
                } catch {
                    Write-Log "Failed to create Start Menu shortcut for $appName : $($_.Exception.Message)" "WARN"
                }
            }
        }
        
        # Pin to taskbar (requires additional registry entries)
        Write-Log "Configuring taskbar pinning preferences..."
        try {
            $taskbarRegPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
            if (!(Test-Path $taskbarRegPath)) {
                New-Item -Path $taskbarRegPath -Force | Out-Null
            }
            
            # Set taskbar pin preferences for key Office apps
            $pinApps = @("WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE", "OUTLOOK.EXE")
            foreach ($pinApp in $pinApps) {
                $appPath = "$env:ProgramFiles\Microsoft Office\root\Office16\$pinApp"
                if (Test-Path $appPath) {
                    # Create file association for better integration
                    $regPath = "HKLM:\SOFTWARE\Classes\Applications\$pinApp"
                    if (!(Test-Path $regPath)) {
                        New-Item -Path $regPath -Force | Out-Null
                        New-ItemProperty -Path $regPath -Name "FriendlyAppName" -Value $officeApps.Keys.Where({$officeApps[$_].Path -like "*$pinApp*"})[0] -PropertyType String -Force | Out-Null
                    }
                }
            }
        } catch {
            Write-Log "Taskbar configuration warning: $($_.Exception.Message)" "WARN"
        }
        
        # Cleanup COM object
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        
    } catch {
        Write-Log "Icon configuration failed: $($_.Exception.Message)" "ERROR"
    }
}

# Verify installation
function Test-OfficeInstallation {
    Write-Log "Verifying 64-bit Office installation..."
    
    # Check 64-bit paths first
    $officeApps64 = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\WINWORD.EXE",
        "$env:ProgramFiles\Microsoft Office\root\Office16\EXCEL.EXE",
        "$env:ProgramFiles\Microsoft Office\root\Office16\POWERPNT.EXE",
        "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "$env:ProgramFiles\Microsoft Office\root\Office16\MSACCESS.EXE"
    )
    
    $installed = $true
    foreach ($app in $officeApps64) {
        if (Test-Path $app) {
            $appName = Split-Path $app -Leaf
            Write-Log "✓ $appName found (64-bit)"
        } else {
            $appName = Split-Path $app -Leaf
            Write-Log "✗ $appName missing (64-bit)" "WARN"
            $installed = $false
        }
    }
    
    # Check optional apps
    $optionalApps = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\MSPUB.EXE",
        "$env:ProgramFiles\Microsoft Office\root\Office16\ONENOTE.EXE"
    )
    
    foreach ($app in $optionalApps) {
        if (Test-Path $app) {
            $appName = Split-Path $app -Leaf
            Write-Log "✓ $appName found (optional, 64-bit)"
        }
    }
    
    # Check architecture
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $regPath) {
        $config = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($config.Platform -eq "x64") {
            Write-Log "✓ Architecture: 64-bit confirmed"
        } else {
            Write-Log "✗ Architecture: $($config.Platform) detected (expected x64)" "WARN"
            $installed = $false
        }
    }
    
    # Verify no 32-bit remnants
    $officeApps32 = @(
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    )
    
    $has32BitRemnants = $false
    foreach ($app in $officeApps32) {
        if (Test-Path $app) {
            Write-Log "⚠ 32-bit remnant detected: $(Split-Path $app -Leaf)" "WARN"
            $has32BitRemnants = $true
        }
    }
    
    if (!$has32BitRemnants) {
        Write-Log "✓ No 32-bit remnants detected"
    }
    
    # Verify no Windows Store versions remain
    try {
        $storeApps = Get-AppxPackage -AllUsers | Where-Object {
            $_.Name -match "Microsoft\.Office|Microsoft\.WindowsOffice"
        }
        
        if ($storeApps.Count -eq 0) {
            Write-Log "✓ No conflicting Windows Store versions detected"
        } else {
            Write-Log "⚠ $($storeApps.Count) Windows Store Office app(s) still present" "WARN"
            foreach ($app in $storeApps) {
                Write-Log "  - $($app.Name) v$($app.Version)" "WARN"
            }
        }
    } catch {
        Write-Log "Could not verify Windows Store apps: $($_.Exception.Message)" "WARN"
    }
    
    # Check shortcuts
    $desktopPath = [Environment]::GetFolderPath("CommonDesktopDirectory")
    $shortcuts = @("Microsoft Word.lnk", "Microsoft Excel.lnk", "Microsoft PowerPoint.lnk", "Microsoft Outlook.lnk")
    
    $shortcutsFound = 0
    foreach ($shortcut in $shortcuts) {
        if (Test-Path "$desktopPath\$shortcut") {
            $shortcutsFound++
        }
    }
    
    Write-Log "✓ $shortcutsFound/$($shortcuts.Count) desktop shortcuts created"
    
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
    Write-Log "=== Microsoft Apps for Business 64-bit Deployment Started ==="
    Write-Log "Script running with PowerShell version: $($PSVersionTable.PSVersion)"
    
    # Check admin privileges
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-NOT $isAdmin) {
        Write-Log "Script requires administrator privileges. Attempting to restart as admin..." "ERROR"
        
        # Attempt to restart as admin
        $scriptPath = $MyInvocation.MyCommand.Path
        $arguments = $MyInvocation.BoundParameters.GetEnumerator() | ForEach-Object { "-$($_.Key) $($_.Value)" }
        
        try {
            Start-Process PowerShell -Verb RunAs -ArgumentList "-ExecutionPolicy Bypass -File `"$scriptPath`" $arguments"
            exit 0
        } catch {
            Write-Log "Failed to restart as administrator. Please run PowerShell as Administrator." "ERROR"
            exit 1
        }
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
        
        # Check if 64-bit is already installed
        $has64Bit = $existingInstalls | Where-Object { $_.Platform -eq "x64" }
        if ($has64Bit -and !$Force) {
            Write-Log "64-bit Office already detected. Use -Force to reinstall." "WARN"
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
    
    # Install 64-bit Office
    $installSuccess = Install-Office64Bit
    
    if ($installSuccess) {
        # Wait for installation to complete
        Write-Log "Waiting for installation to stabilize..."
        Start-Sleep -Seconds 60
        
        # Configure icons and shortcuts
        Set-OfficeIcons
        
        # Verify installation
        if (Test-OfficeInstallation) {
            Write-Log "=== Deployment completed successfully ===" "INFO"
            Write-Log "Microsoft Apps for Business 64-bit with Access installed"
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
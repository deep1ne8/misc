#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Automated Microsoft Apps for Business 32-bit Deployment Script
.DESCRIPTION
    Detects and removes 64-bit Microsoft Apps for Business, then installs 32-bit version
    with Access, Current Channel, English language
.PARAMETER Force
    Forces reinstallation even if 32-bit Office is already installed
.PARAMETER Interactive
    Prompts user for Force confirmation when 32-bit Office is detected
.PARAMETER SkipRemoval
    Skip removal of existing installations (for troubleshooting)
.NOTES
    Version: 1.1
    Requires: Administrator privileges
    Author: PowerShell Automation Script
.EXAMPLE
    .\OfficeDeployment.ps1
    Standard deployment with interactive prompts
.EXAMPLE
    .\OfficeDeployment.ps1 -Force
    Force reinstallation regardless of existing installations
.EXAMPLE
    .\OfficeDeployment.ps1 -Interactive
    Prompt user when 32-bit Office is detected
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$env:TEMP\OfficeDeployment.log",
    [string]$WorkingDir = "$env:TEMP\OfficeDeployment",
    [switch]$Force,
    [switch]$Interactive,
    [switch]$SkipRemoval
)

# Initialize logging
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    $color = switch($Level) {
        "ERROR" { "Red" }
        "WARN" { "Yellow" }
        "SUCCESS" { "Green" }
        "INFO" { "Cyan" }
        default { "White" }
    }
    Write-Host $logMessage -ForegroundColor $color
    Add-Content -Path $LogPath -Value $logMessage -Force
}

# Enhanced user prompt for Force option
function Get-ForceDecision {
    param(
        [string]$ExistingVersion,
        [string]$ExistingPlatform
    )
    
    Write-Host "`n" -NoNewline
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    Write-Host "                    EXISTING OFFICE DETECTED                    " -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    Write-Host "Current Installation Details:" -ForegroundColor Cyan
    Write-Host "  • Platform: $ExistingPlatform" -ForegroundColor White
    Write-Host "  • Version: $ExistingVersion" -ForegroundColor White
    
    if ($ExistingPlatform -eq "x86") {
        Write-Host "`n32-bit Office is already installed!" -ForegroundColor Green
        Write-Host "Do you want to proceed with reinstallation?" -ForegroundColor Yellow
    } else {
        Write-Host "`n64-bit Office detected - recommending 32-bit installation" -ForegroundColor Yellow
        Write-Host "Do you want to proceed with replacement?" -ForegroundColor Yellow
    }
    
    Write-Host "`nOptions:" -ForegroundColor Cyan
    Write-Host "  [Y] Yes - Continue with installation (Force)" -ForegroundColor Green
    Write-Host "  [N] No  - Exit without changes" -ForegroundColor Red
    Write-Host "  [S] Skip - Continue without removing existing Office" -ForegroundColor Yellow
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    
    do {
        $response = Read-Host "Please select an option [Y/N/S]"
        $response = $response.ToUpper()
        
        switch ($response) {
            "Y" { 
                Write-Log "User selected: Force installation" "INFO"
                return "FORCE" 
            }
            "N" { 
                Write-Log "User selected: Exit without changes" "INFO"
                return "EXIT" 
            }
            "S" { 
                Write-Log "User selected: Skip removal" "INFO"
                return "SKIP" 
            }
            default { 
                Write-Host "Invalid selection. Please enter Y, N, or S." -ForegroundColor Red 
            }
        }
    } while ($true)
}

# Display script banner with Force status
function Show-ScriptBanner {
    $forceStatus = if (!$Force) { "ENABLED" } else { "DISABLED" }
    $interactiveStatus = if (!$Interactive) { "ENABLED" } else { "DISABLED" }
    
    Write-Host "`n" -NoNewline
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Blue
    Write-Host "          MICROSOFT OFFICE 32-BIT DEPLOYMENT SCRIPT            " -ForegroundColor Blue
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Blue
    Write-Host "Configuration:" -ForegroundColor Cyan
    Write-Host "  • Force Mode: $forceStatus" -ForegroundColor $(if($Force){"Green"}else{"Yellow"})
    Write-Host "  • Interactive Mode: $interactiveStatus" -ForegroundColor $(if($Interactive){"Green"}else{"Yellow"})
    Write-Host "  • Skip Removal: $(if($SkipRemoval){"ENABLED"}else{"DISABLED"})" -ForegroundColor $(if($SkipRemoval){"Yellow"}else{"White"})
    Write-Host "  • Log Path: $LogPath" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Blue
    Write-Host ""
}

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
            Platform = "Unknown"
        }
    }
    
    return $installations
}

# Enhanced removal function with Force awareness
function Remove-OfficeInstallations {
    param([array]$Installations)
    
    if ($SkipRemoval) {
        Write-Log "Skipping removal as requested" "WARN"
        return
    }
    
    Write-Log "Removing existing Office installations..."
    
    foreach ($install in $Installations) {
        try {
            if ($install.Type -eq "Click-to-Run") {
                Write-Log "Removing Click-to-Run installation..." "INFO"
                
                # Download ODT if needed
                $odtPath = "$WorkingDir\setup.exe"
                if (!(Test-Path $odtPath)) {
                    Write-Log "Downloading Office Deployment Tool..."
                    if (!(Get-ODTTool -DestinationPath $WorkingDir)) {
                        throw "Failed to obtain Office Deployment Tool"
                    }
                }
                
                # Create removal configuration
                $removeConfig = @"
<Configuration>
  <Remove All="TRUE" />
</Configuration>
"@
                $removeConfig | Out-File "$WorkingDir\remove.xml" -Encoding UTF8
                
                # Execute removal
                Write-Log "Executing Office removal..." "INFO"
                $removeProcess = Start-Process -FilePath $odtPath -ArgumentList "/configure", "$WorkingDir\remove.xml" -Wait -PassThru -NoNewWindow
                
                if ($removeProcess.ExitCode -eq 0) {
                    Write-Log "Click-to-Run removal completed successfully" "SUCCESS"
                } else {
                    Write-Log "Click-to-Run removal completed with warnings (Exit Code: $($removeProcess.ExitCode))" "WARN"
                }
                
            } elseif ($install.Type -eq "MSI") {
                Write-Log "Removing MSI installation: $($install.Name)" "INFO"
                $msiProcess = Start-Process -FilePath "msiexec.exe" -ArgumentList "/x", $install.IdentifyingNumber, "/quiet", "/norestart" -Wait -PassThru
                
                if ($msiProcess.ExitCode -eq 0) {
                    Write-Log "MSI removal completed successfully: $($install.Name)" "SUCCESS"
                } else {
                    Write-Log "MSI removal completed with warnings: $($install.Name) (Exit Code: $($msiProcess.ExitCode))" "WARN"
                }
            }
        } catch {
            Write-Log "Failed to remove installation: $($_.Exception.Message)" "ERROR"
        }
    }
    
    # Enhanced registry cleanup
    Write-Log "Performing registry cleanup..." "INFO"
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
        "HKCU:\SOFTWARE\Microsoft\Office"
    )
    
    foreach ($regPath in $regPaths) {
        try {
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Cleaned registry path: $regPath" "INFO"
            }
        } catch {
            Write-Log "Registry cleanup warning for $regPath : $($_.Exception.Message)" "WARN"
        }
    }
}

# Enhanced ODT download function
function Get-ODTTool {
    param([string]$DestinationPath)
    
    Write-Log "Obtaining Office Deployment Tool..." "INFO"
    
    # Multiple ODT download sources
    $odtSources = @(
        @{
            Name = "Microsoft Direct Download"
            Url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17830-20162.exe"
        },
        @{
            Name = "Alternative Microsoft Link"
            Url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16626-20170.exe"
        }
    )
    
    foreach ($source in $odtSources) {
        try {
            Write-Log "Attempting download from: $($source.Name)" "INFO"
            Invoke-WebRequest -Uri $source.Url -OutFile "$DestinationPath\odt.exe" -UseBasicParsing -TimeoutSec 60
            
            if (Test-Path "$DestinationPath\odt.exe") {
                Write-Log "ODT downloaded successfully from: $($source.Name)" "SUCCESS"
                
                # Extract ODT
                Write-Log "Extracting Office Deployment Tool..." "INFO"
                $extractProcess = Start-Process -FilePath "$DestinationPath\odt.exe" -ArgumentList "/quiet", "/extract:$DestinationPath" -Wait -PassThru
                
                if ($extractProcess.ExitCode -eq 0 -and (Test-Path "$DestinationPath\setup.exe")) {
                    Write-Log "ODT extraction completed successfully" "SUCCESS"
                    return $true
                } else {
                    Write-Log "ODT extraction failed" "ERROR"
                }
            }
        } catch {
            Write-Log "Download failed from $($source.Name): $($_.Exception.Message)" "WARN"
        }
    }
    
    Write-Log "All ODT download attempts failed" "ERROR"
    Write-Log "Please download ODT manually from: https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool" "ERROR"
    return $false
}

# Install 32-bit Microsoft Apps for Business
function Install-Office32Bit {
    Write-Log "Installing 32-bit Microsoft Apps for Business..." "INFO"
    
    try {
        # Ensure ODT is available
        $odtPath = "$WorkingDir\setup.exe"
        if (!(Test-Path $odtPath)) {
            if (!(Get-ODTTool -DestinationPath $WorkingDir)) {
                throw "Office Deployment Tool not available"
            }
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
        Write-Log "Installation configuration created" "SUCCESS"
        
        # Execute installation
        Write-Log "Starting Office installation... This may take several minutes." "INFO"
        Write-Host "Please wait while Office is being installed..." -ForegroundColor Yellow
        
        $installProcess = Start-Process -FilePath $odtPath -ArgumentList "/configure", "$WorkingDir\install.xml" -Wait -PassThru -NoNewWindow
        
        if ($installProcess.ExitCode -eq 0) {
            Write-Log "Office installation completed successfully" "SUCCESS"
            return $true
        } else {
            Write-Log "Office installation failed with exit code: $($installProcess.ExitCode)" "ERROR"
            return $false
        }
        
    } catch {
        Write-Log "Installation failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Verify installation
function Test-OfficeInstallation {
    Write-Log "Verifying Office installation..." "INFO"
    
    $officeApps = @(
        @{ Path = "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\WINWORD.EXE"; Name = "Word" },
        @{ Path = "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\EXCEL.EXE"; Name = "Excel" },
        @{ Path = "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"; Name = "PowerPoint" },
        @{ Path = "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"; Name = "Outlook" },
        @{ Path = "$env:ProgramFiles (x86)\Microsoft Office\root\Office16\MSACCESS.EXE"; Name = "Access" }
    )
    
    $installed = $true
    $installedApps = @()
    $missingApps = @()
    
    foreach ($app in $officeApps) {
        if (Test-Path $app.Path) {
            Write-Log "✓ $($app.Name) installed successfully" "SUCCESS"
            $installedApps += $app.Name
        } else {
            Write-Log "✗ $($app.Name) not found" "WARN"
            $missingApps += $app.Name
            $installed = $false
        }
    }
    
    # Check architecture
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (Test-Path $regPath) {
        $config = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($config.Platform -eq "x86") {
            Write-Log "✓ Architecture: 32-bit confirmed" "SUCCESS"
        } else {
            Write-Log "✗ Architecture: $($config.Platform) detected (Expected: x86)" "WARN"
            $installed = $false
        }
    }
    
    # Summary
    Write-Host "`nInstallation Summary:" -ForegroundColor Cyan
    Write-Host "Installed Apps: $($installedApps -join ', ')" -ForegroundColor Green
    if ($missingApps.Count -gt 0) {
        Write-Host "Missing Apps: $($missingApps -join ', ')" -ForegroundColor Red
    }
    
    return $installed
}

# Cleanup function
function Remove-WorkingDirectory {
    try {
        if (Test-Path $WorkingDir) {
            Remove-Item $WorkingDir -Recurse -Force
            Write-Log "Working directory cleaned up" "INFO"
        }
    } catch {
        Write-Log "Cleanup warning: $($_.Exception.Message)" "WARN"
    }
}

# Main execution
try {
    Show-ScriptBanner
    Write-Log "=== Microsoft Apps for Business 32-bit Deployment Started ===" "INFO"
    Write-Log "Script running with PowerShell version: $($PSVersionTable.PSVersion)" "INFO"
    
    # Check admin privileges
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-NOT $isAdmin) {
        Write-Log "Script requires administrator privileges. Attempting to restart as admin..." "ERROR"
        
        # Prepare arguments for restart
        $scriptPath = $MyInvocation.MyCommand.Path
        $argList = @()
        if ($Force) { $argList += "-Force" }
        if ($Interactive) { $argList += "-Interactive" }
        if ($SkipRemoval) { $argList += "-SkipRemoval" }
        if ($LogPath -ne "$env:TEMP\OfficeDeployment.log") { $argList += "-LogPath `"$LogPath`"" }
        if ($WorkingDir -ne "$env:TEMP\OfficeDeployment") { $argList += "-WorkingDir `"$WorkingDir`"" }
        
        try {
            $argumentString = "-ExecutionPolicy Bypass -File `"$scriptPath`" " + ($argList -join " ")
            Start-Process PowerShell -Verb RunAs -ArgumentList $argumentString
            exit 0
        } catch {
            Write-Log "Failed to restart as administrator. Please run PowerShell as Administrator." "ERROR"
            exit 1
        }
    }
    
    # Stop Office processes
    Write-Log "Stopping Office processes..." "INFO"
    $officeProcesses = @("winword", "excel", "powerpnt", "outlook", "msaccess", "lync", "communicator", "teams")
    $stoppedProcesses = @()
    
    foreach ($process in $officeProcesses) {
        $runningProcess = Get-Process -Name $process -ErrorAction SilentlyContinue
        if ($runningProcess) {
            Stop-Process -Name $process -Force -ErrorAction SilentlyContinue
            $stoppedProcesses += $process
        }
    }
    
    if ($stoppedProcesses.Count -gt 0) {
        Write-Log "Stopped processes: $($stoppedProcesses -join ', ')" "INFO"
    }
    
    # Detect existing installations
    $existingInstalls = Get-OfficeInstallations
    
    if ($existingInstalls.Count -gt 0) {
        Write-Log "Found $($existingInstalls.Count) existing Office installation(s)" "INFO"
        
        $primaryInstall = $existingInstalls | Where-Object { $_.Type -eq "Click-to-Run" } | Select-Object -First 1
        if (-not $primaryInstall) {
            $primaryInstall = $existingInstalls | Select-Object -First 1
        }
        
        foreach ($install in $existingInstalls) {
            Write-Log "- Type: $($install.Type), Platform: $($install.Platform), Version: $($install.Version)" "INFO"
        }
        
        # Check if 32-bit is already installed and handle Force logic
        $has32Bit = $existingInstalls | Where-Object { $_.Platform -eq "x86" }
        
        if ($has32Bit -and -not $Force) {
            if ($Interactive) {
                $decision = Get-ForceDecision -ExistingVersion $primaryInstall.Version -ExistingPlatform $primaryInstall.Platform
                
                switch ($decision) {
                    "FORCE" { 
                        $Force = $true
                        Write-Log "Force mode activated by user selection" "INFO"
                    }
                    "SKIP" { 
                        $SkipRemoval = $true
                        Write-Log "Skip removal activated by user selection" "INFO"
                    }
                    "EXIT" { 
                        Write-Log "Deployment cancelled by user" "INFO"
                        exit 0 
                    }
                }
            } else {
                Write-Log "32-bit Office already detected." "WARN"
                Write-Log "Use -Force parameter to force reinstallation" "WARN"
                Write-Log "Use -Interactive parameter for guided options" "WARN"
                exit 0
            }
        }
        
        # Remove existing installations if not skipping
        if (-not $SkipRemoval) {
            Remove-OfficeInstallations -Installations $existingInstalls
            Write-Log "Waiting for removal to complete..." "INFO"
            Start-Sleep -Seconds 30
        }
    } else {
        Write-Log "No existing Office installations detected" "INFO"
    }
    
    # Install 32-bit Office
    $installSuccess = Install-Office32Bit
    
    if ($installSuccess) {
        Write-Log "Waiting for installation to stabilize..." "INFO"
        Start-Sleep -Seconds 60
        
        # Verify installation
        if (Test-OfficeInstallation) {
            Write-Log "=== DEPLOYMENT COMPLETED SUCCESSFULLY ===" "SUCCESS"
            Write-Log "Microsoft Apps for Business 32-bit with Access installed and verified" "SUCCESS"
            Write-Host "`n🎉 Deployment completed successfully!" -ForegroundColor Green
        } else {
            Write-Log "=== DEPLOYMENT COMPLETED WITH WARNINGS ===" "WARN"
            Write-Log "Some components may not have installed correctly" "WARN"
            Write-Host "`n⚠️  Deployment completed with warnings. Check log for details." -ForegroundColor Yellow
        }
    } else {
        Write-Log "=== DEPLOYMENT FAILED ===" "ERROR"
        Write-Host "`n❌ Deployment failed. Check log for details." -ForegroundColor Red
        exit 1
    }
    
} catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    Write-Host "`n❌ Critical error occurred. Check log for details." -ForegroundColor Red
    exit 1
} finally {
    Remove-WorkingDirectory
    Write-Log "Log file saved to: $LogPath" "INFO"
    Write-Host "`nLog file: $LogPath" -ForegroundColor Cyan
}

exit 0

<#
.SYNOPSIS
    ShareX 18.0.1 - Complete Intune Platform Script Deployment
.DESCRIPTION
    All-in-one script for deploying ShareX via Intune Platform Scripts (Proactive Remediations/Scripts).
    Automatically detects if installation is needed and installs if missing or outdated.
    
    Features:
    - Auto-detection of installed version
    - Silent installation with retry logic
    - PowerShell 5.1 and 7 compatible
    - Comprehensive logging for troubleshooting
    - SYSTEM context compatible
    - Proper exit codes for Intune reporting
    
.PARAMETER Mode
    Operation mode:
    - 'Auto' (default): Detect and install if needed
    - 'DetectOnly': Only check installation status
    - 'ForceInstall': Install regardless of current state
    
.EXAMPLE
    .\Deploy-ShareX-Complete.ps1
    Default behavior: Checks if installed, installs if needed
    
.EXAMPLE
    .\Deploy-ShareX-Complete.ps1 -Mode DetectOnly
    Only checks if ShareX 18.0.1 is installed
    
.EXAMPLE
    .\Deploy-ShareX-Complete.ps1 -Mode ForceInstall
    Forces installation even if already present
    
.NOTES
    Author: IT Department
    Version: 1.0
    Compatible: PowerShell 5.1+, Windows 10/11
    Deployment: Intune Platform Scripts, RMM tools, Manual execution
    
    Exit Codes:
        0 = Success (installed or already present)
        1 = Download failure
        2 = Installation failure
        3 = Validation failure
        10 = Not installed (DetectOnly mode)
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('Auto','DetectOnly','ForceInstall')]
    [string]$Mode = 'Auto'
)

#region Configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Application Configuration
$Config = @{
    AppName         = "ShareX"
    AppVersion      = "18.0.1"
    Publisher       = "ShareX Team"
    DownloadUrl     = "https://github.com/ShareX/ShareX/releases/download/v18.0.1/ShareX-18.0.1-setup.exe"
    
    # Paths
    WorkRoot        = "$env:ProgramData\ShareXDeploy"
    LogDir          = "$env:ProgramData\ShareXDeploy\Logs"
    
    # Installation Settings
    RetryAttempts   = 3
    RetryDelay      = 5
    MinFileSize     = 1048576  # 1 MB minimum expected download size
    
    # Expected Installation Paths
    ExePaths        = @(
        "$env:ProgramFiles\ShareX\ShareX.exe"
        "${env:ProgramFiles(x86)}\ShareX\ShareX.exe"
    )
    
    # Registry Paths
    RegPaths        = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall"
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )
}

# Generate dynamic paths
$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$Config.InstallerPath = Join-Path $Config.WorkRoot "ShareX-$($Config.AppVersion)-setup.exe"
$Config.LogPath = Join-Path $Config.LogDir "deployment-$timestamp.log"
$Config.LastRunPath = Join-Path $Config.WorkRoot "last-run.txt"

#endregion Configuration

#region Functions

function Initialize-Environment {
    <#
    .SYNOPSIS
        Prepares the deployment environment
    #>
    try {
        # Create directories
        @($Config.WorkRoot, $Config.LogDir) | ForEach-Object {
            if (!(Test-Path $_)) {
                $null = New-Item -Path $_ -ItemType Directory -Force
            }
        }
        
        # Configure TLS for secure downloads (PS 5.1 and 7 compatible)
        $currentProtocol = [Net.ServicePointManager]::SecurityProtocol
        if ($currentProtocol -notmatch 'Tls12') {
            try {
                # Try to set TLS 1.2
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            }
            catch {
                # Fallback for older .NET versions
                [Net.ServicePointManager]::SecurityProtocol = 3072
            }
        }
        
        return $true
    }
    catch {
        Write-Output "ERROR: Failed to initialize environment: $_"
        return $false
    }
}

function Write-Log {
    <#
    .SYNOPSIS
        Writes log entries to both file and console
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [Parameter()]
        [ValidateSet('Info','Success','Warning','Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to console for Intune capture
    switch ($Level) {
        'Success' { Write-Output "SUCCESS: $Message" }
        'Warning' { Write-Output "WARNING: $Message" }
        'Error'   { Write-Output "ERROR: $Message" }
        default   { Write-Output $Message }
    }
    
    # Write to log file
    try {
        Add-Content -Path $Config.LogPath -Value $logEntry -ErrorAction SilentlyContinue
    }
    catch {
        # Silently continue if logging fails
    }
}

function Test-ShareXInstalled {
    <#
    .SYNOPSIS
        Checks if ShareX is installed and returns version info
    #>
    try {
        Write-Log "Checking for existing ShareX installation..."
        
        # Check registry
        foreach ($regPath in $Config.RegPaths) {
            if (Test-Path $regPath) {
                $apps = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
                
                foreach ($app in $apps) {
                    $appInfo = Get-ItemProperty -Path $app.PSPath -ErrorAction SilentlyContinue
                    
                    if ($appInfo.DisplayName -eq $Config.AppName) {
                        $installedVersion = $appInfo.DisplayVersion
                        Write-Log "Found $($Config.AppName) version $installedVersion in registry"
                        
                        # Verify executable exists
                        $exeFound = $false
                        foreach ($exePath in $Config.ExePaths) {
                            if (Test-Path $exePath) {
                                Write-Log "Verified executable at: $exePath"
                                $exeFound = $true
                                break
                            }
                        }
                        
                        if ($exeFound) {
                            return @{
                                Installed = $true
                                Version = $installedVersion
                                IsTargetVersion = ($installedVersion -eq $Config.AppVersion)
                            }
                        }
                        else {
                            Write-Log "Registry entry found but executable missing" -Level Warning
                        }
                    }
                }
            }
        }
        
        Write-Log "ShareX is not installed"
        return @{
            Installed = $false
            Version = $null
            IsTargetVersion = $false
        }
    }
    catch {
        Write-Log "Error checking installation status: $_" -Level Error
        return @{
            Installed = $false
            Version = $null
            IsTargetVersion = $false
        }
    }
}

function Get-ShareXInstaller {
    <#
    .SYNOPSIS
        Downloads the ShareX installer with retry logic
    #>
    param(
        [int]$MaxAttempts = $Config.RetryAttempts
    )
    
    Write-Log "Downloading ShareX installer from GitHub..."
    Write-Log "URL: $($Config.DownloadUrl)"
    
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($attempt -gt 1) {
                Write-Log "Retry attempt $attempt of $MaxAttempts" -Level Warning
                Start-Sleep -Seconds $Config.RetryDelay
            }
            
            # Download using .NET WebClient (PS 5.1 and 7 compatible)
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($Config.DownloadUrl, $Config.InstallerPath)
            $webClient.Dispose()
            
            # Verify download
            if (Test-Path $Config.InstallerPath) {
                $fileInfo = Get-Item $Config.InstallerPath
                $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
                
                if ($fileInfo.Length -lt $Config.MinFileSize) {
                    throw "Downloaded file is too small ($fileSizeMB MB). Possible corruption."
                }
                
                Write-Log "Download successful: $fileSizeMB MB" -Level Success
                return $true
            }
            else {
                throw "Installer file not found after download"
            }
        }
        catch {
            Write-Log "Download attempt $attempt failed: $_" -Level Error
            
            if ($attempt -eq $MaxAttempts) {
                Write-Log "All download attempts exhausted" -Level Error
                return $false
            }
        }
    }
    
    return $false
}

function Install-ShareX {
    <#
    .SYNOPSIS
        Performs silent installation of ShareX
    #>
    try {
        Write-Log "Starting ShareX installation..."
        
        # Inno Setup silent installation arguments
        $installArgs = @(
            '/VERYSILENT'
            '/SUPPRESSMSGBOXES'
            '/NORESTART'
            '/SP-'
            '/NOCANCEL'
        ) -join ' '
        
        Write-Log "Install arguments: $installArgs"
        
        # Start installation process (PS 5.1 and 7 compatible method)
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $Config.InstallerPath
        $processInfo.Arguments = $installArgs
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        $processInfo.RedirectStandardOutput = $false
        $processInfo.RedirectStandardError = $false
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        Write-Log "Launching installer..."
        $null = $process.Start()
        $process.WaitForExit()
        
        $exitCode = $process.ExitCode
        Write-Log "Installer exit code: $exitCode"
        
        # Interpret Inno Setup exit codes
        if ($exitCode -eq 0) {
            Write-Log "Installation completed successfully" -Level Success
            return $true
        }
        else {
            $errorMessage = switch ($exitCode) {
                1 { "Setup initialization error" }
                2 { "User cancelled (shouldn't happen in silent mode)" }
                3 { "Fatal error during installation" }
                4 { "Fatal error before installation" }
                5 { "User chose not to proceed" }
                6 { "Setup initialization error (variant)" }
                7 { "Installation aborted" }
                8 { "Insufficient disk space" }
                default { "Unknown error (code: $exitCode)" }
            }
            
            Write-Log "Installation failed: $errorMessage" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Installation exception: $_" -Level Error
        return $false
    }
}

function Test-ShareXValidation {
    <#
    .SYNOPSIS
        Validates ShareX installation post-deployment
    #>
    try {
        Write-Log "Validating installation..."
        
        # Wait briefly for registry to update
        Start-Sleep -Seconds 2
        
        # Check for executable
        $exeFound = $false
        $exePath = $null
        
        foreach ($path in $Config.ExePaths) {
            if (Test-Path $path) {
                $exeFound = $true
                $exePath = $path
                break
            }
        }
        
        if (!$exeFound) {
            Write-Log "Validation failed: ShareX.exe not found" -Level Error
            return $false
        }
        
        # Get file version
        try {
            $versionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($exePath)
            $fileVersion = $versionInfo.FileVersion
            Write-Log "Validated: ShareX.exe found at $exePath (v$fileVersion)" -Level Success
        }
        catch {
            Write-Log "Validated: ShareX.exe found at $exePath (version unknown)" -Level Success
        }
        
        # Check registry
        $regFound = $false
        foreach ($regPath in $Config.RegPaths) {
            if (Test-Path $regPath) {
                $apps = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
                foreach ($app in $apps) {
                    $appInfo = Get-ItemProperty -Path $app.PSPath -ErrorAction SilentlyContinue
                    if ($appInfo.DisplayName -eq $Config.AppName) {
                        Write-Log "Registry entry confirmed: $($appInfo.DisplayName) v$($appInfo.DisplayVersion)"
                        $regFound = $true
                        
                        # Save uninstall info
                        try {
                            $uninstallInfo = @"
Application: $($appInfo.DisplayName)
Version: $($appInfo.DisplayVersion)
Publisher: $($appInfo.Publisher)
Install Date: $($appInfo.InstallDate)
Install Location: $($appInfo.InstallLocation)
Uninstall String: $($appInfo.UninstallString)
"@
                            $uninstallInfoPath = Join-Path $Config.WorkRoot "uninstall-info.txt"
                            $uninstallInfo | Out-File -FilePath $uninstallInfoPath -Force
                        }
                        catch {
                            # Non-critical failure
                        }
                        
                        break
                    }
                }
            }
            if ($regFound) { break }
        }
        
        if (!$regFound) {
            Write-Log "Warning: Registry entry not found (may update on next system scan)" -Level Warning
        }
        
        return $true
    }
    catch {
        Write-Log "Validation exception: $_" -Level Error
        return $false
    }
}

function Remove-InstallerFiles {
    <#
    .SYNOPSIS
        Cleans up installer files after deployment
    #>
    try {
        if (Test-Path $Config.InstallerPath) {
            Remove-Item -Path $Config.InstallerPath -Force -ErrorAction Stop
            Write-Log "Cleaned up installer file"
        }
    }
    catch {
        Write-Log "Failed to remove installer: $_" -Level Warning
    }
}

function Save-DeploymentStatus {
    <#
    .SYNOPSIS
        Saves deployment status for tracking
    #>
    param(
        [string]$Status,
        [string]$Version
    )
    
    try {
        $statusInfo = @"
Last Run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Status: $Status
Version: $Version
Computer: $env:COMPUTERNAME
User Context: $env:USERNAME
PowerShell: $($PSVersionTable.PSVersion)
Mode: $Mode
"@
        $statusInfo | Out-File -FilePath $Config.LastRunPath -Force
    }
    catch {
        # Non-critical failure
    }
}

#endregion Functions

#region Main Execution

try {
    # Header
    Write-Output "=========================================="
    Write-Output "ShareX $($Config.AppVersion) Deployment"
    Write-Output "=========================================="
    Write-Log "Script started in $Mode mode"
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Log "Computer: $env:COMPUTERNAME"
    Write-Log "User Context: $env:USERNAME"
    Write-Log "Script Path: $PSCommandPath"
    
    # Initialize environment
    if (!(Initialize-Environment)) {
        Write-Log "Failed to initialize environment" -Level Error
        exit 99
    }
    
    # Check current installation status
    $installStatus = Test-ShareXInstalled
    
    #region DetectOnly Mode
    if ($Mode -eq 'DetectOnly') {
        Write-Output "------------------------------------------"
        Write-Log "Running in detection-only mode"
        
        if ($installStatus.IsTargetVersion) {
            Write-Log "DETECTED: ShareX $($Config.AppVersion) is installed" -Level Success
            Save-DeploymentStatus -Status "Detected" -Version $installStatus.Version
            exit 0
        }
        elseif ($installStatus.Installed) {
            Write-Log "Found ShareX version $($installStatus.Version), but target is $($Config.AppVersion)" -Level Warning
            Save-DeploymentStatus -Status "Wrong Version" -Version $installStatus.Version
            exit 10
        }
        else {
            Write-Log "NOT DETECTED: ShareX is not installed" -Level Warning
            Save-DeploymentStatus -Status "Not Installed" -Version "None"
            exit 10
        }
    }
    #endregion DetectOnly Mode
    
    #region Auto and ForceInstall Modes
    Write-Output "------------------------------------------"
    
    # Check if installation is needed
    if ($installStatus.IsTargetVersion -and $Mode -ne 'ForceInstall') {
        Write-Log "ShareX $($Config.AppVersion) is already installed" -Level Success
        Write-Log "No action required"
        Save-DeploymentStatus -Status "Already Installed" -Version $installStatus.Version
        exit 0
    }
    
    if ($installStatus.Installed -and $Mode -ne 'ForceInstall') {
        Write-Log "Upgrading from version $($installStatus.Version) to $($Config.AppVersion)" -Level Info
    }
    elseif ($Mode -eq 'ForceInstall') {
        Write-Log "Force install requested - proceeding with installation" -Level Warning
    }
    else {
        Write-Log "ShareX not detected - proceeding with installation"
    }
    
    # Download installer
    Write-Output "------------------------------------------"
    if (!(Get-ShareXInstaller)) {
        Write-Log "Deployment failed: Unable to download installer" -Level Error
        Save-DeploymentStatus -Status "Download Failed" -Version "N/A"
        exit 1
    }
    
    # Install ShareX
    Write-Output "------------------------------------------"
    if (!(Install-ShareX)) {
        Write-Log "Deployment failed: Installation error" -Level Error
        Save-DeploymentStatus -Status "Installation Failed" -Version "N/A"
        exit 2
    }
    
    # Validate installation
    Write-Output "------------------------------------------"
    if (!(Test-ShareXValidation)) {
        Write-Log "Deployment failed: Post-installation validation failed" -Level Error
        Save-DeploymentStatus -Status "Validation Failed" -Version "N/A"
        exit 3
    }
    
    # Cleanup
    Remove-InstallerFiles
    
    # Success
    Write-Output "------------------------------------------"
    Write-Log "ShareX $($Config.AppVersion) deployment completed successfully" -Level Success
    Write-Output "=========================================="
    Save-DeploymentStatus -Status "Success" -Version $Config.AppVersion
    exit 0
    
    #endregion Auto and ForceInstall Modes
}
catch {
    # Catch-all error handler
    Write-Log "CRITICAL ERROR: $_" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
    Save-DeploymentStatus -Status "Critical Error" -Version "N/A"
    exit 99
}
finally {
    Write-Output ""
    Write-Log "Log file: $($Config.LogPath)"
    Write-Log "Script execution completed"
}

#endregion Main Execution

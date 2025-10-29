<# 
.SYNOPSIS
    Silent installer for ShareX 18.0.1 - Intune Compatible
.DESCRIPTION
    - Intune Win32 app compatible with proper exit codes and console output
    - Works on PowerShell 5.1+ and PowerShell 7+
    - Checks if already installed (skips if current version exists)
    - Downloads and installs silently with comprehensive logging
    - Outputs status to console for Intune capture
.NOTES
    Requires: PowerShell 5.1+ | Runs in SYSTEM context via Intune
    Exit Codes:
        0   = Success (installation completed or already installed)
        1   = Download failed
        2   = Installation failed
        3   = Post-validation failed
.EXAMPLE
    powershell.exe -ExecutionPolicy Bypass -File Deploy-ShareX.ps1
#>

#Requires -Version 5.1

#region Configuration
$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$Config = @{
    AppName       = "ShareX"
    AppVersion    = "18.0.1"
    DownloadUrl   = "https://github.com/ShareX/ShareX/releases/download/v18.0.1/ShareX-18.0.1-setup.exe"
    WorkRoot      = "$env:ProgramData\ShareXDeploy"
    RetryAttempts = 3
    RetryDelay    = 5
}

$Config.InstallerPath = Join-Path $Config.WorkRoot "ShareX-$($Config.AppVersion)-setup.exe"
$Config.LogPath       = Join-Path $Config.WorkRoot "install-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
#endregion Configuration

#region Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $Config.LogPath -Value $logEntry -ErrorAction SilentlyContinue
    
    # Write to console (Intune captures this)
    Write-Host $logEntry
    
    # Also write to appropriate stream for Intune
    switch ($Level) {
        'Error'   { Write-Error $Message -ErrorAction Continue }
        'Warning' { Write-Warning $Message }
        'Success' { Write-Host "SUCCESS: $Message" -ForegroundColor Green }
    }
}

function Initialize-Environment {
    try {
        # Create working directory
        if (!(Test-Path $Config.WorkRoot)) {
            $null = New-Item -Path $Config.WorkRoot -ItemType Directory -Force
            Write-Log "Created working directory: $($Config.WorkRoot)"
        }
        
        # Ensure TLS 1.2+ for secure downloads (PS 5.1 and 7 compatible)
        $securityProtocol = [Net.ServicePointManager]::SecurityProtocol
        if ($securityProtocol -notmatch 'Tls12') {
            try {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Write-Log "Enabled TLS 1.2 for secure downloads"
            }
            catch {
                # Fallback for older systems
                [Net.ServicePointManager]::SecurityProtocol = 3072
                Write-Log "Enabled TLS 1.2 (compatibility mode)"
            }
        }
        
        return $true
    }
    catch {
        Write-Log "Environment initialization failed: $_" -Level Error
        return $false
    }
}

function Test-AlreadyInstalled {
    try {
        # Check registry for installed version (PS 5/7 compatible method)
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        
        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                $apps = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
                foreach ($app in $apps) {
                    $appInfo = Get-ItemProperty -Path $app.PSPath -ErrorAction SilentlyContinue
                    if ($appInfo.DisplayName -eq $Config.AppName) {
                        $installedVersion = $appInfo.DisplayVersion
                        Write-Log "Found existing installation: $($Config.AppName) v$installedVersion"
                        
                        if ($installedVersion -eq $Config.AppVersion) {
                            Write-Log "Target version $($Config.AppVersion) already installed" -Level Success
                            return $true
                        }
                        Write-Log "Upgrading from v$installedVersion to v$($Config.AppVersion)" -Level Info
                    }
                }
            }
        }
        return $false
    }
    catch {
        Write-Log "Error checking installation status: $_" -Level Warning
        return $false
    }
}

function Get-Installer {
    Write-Log "Downloading installer from $($Config.DownloadUrl)"
    
    $attempt = 0
    $downloaded = $false
    
    while ($attempt -lt $Config.RetryAttempts -and !$downloaded) {
        $attempt++
        try {
            if ($attempt -gt 1) {
                Write-Log "Retry attempt $attempt of $($Config.RetryAttempts)" -Level Warning
                Start-Sleep -Seconds $Config.RetryDelay
            }
            
            # PS 5/7 compatible download method
            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFile($Config.DownloadUrl, $Config.InstallerPath)
            $webClient.Dispose()
            
            # Verify file exists and has content
            if (Test-Path $Config.InstallerPath) {
                $fileInfo = Get-Item $Config.InstallerPath
                if ($fileInfo.Length -lt 1MB) {
                    throw "Downloaded file is too small ($($fileInfo.Length) bytes)"
                }
                
                $sizeMB = [math]::Round($fileInfo.Length/1MB, 2)
                Write-Log "Download complete. Size: $sizeMB MB" -Level Success
                $downloaded = $true
            }
            else {
                throw "Installer file not found after download"
            }
        }
        catch {
            Write-Log "Download failed: $_" -Level Error
            if ($attempt -eq $Config.RetryAttempts) {
                throw "Failed to download after $($Config.RetryAttempts) attempts"
            }
        }
    }
    
    return $downloaded
}

function Install-Application {
    # Inno Setup silent switches
    $installArgs = @(
        "/VERYSILENT"
        "/SUPPRESSMSGBOXES"
        "/NORESTART"
        "/SP-"
        "/NOCANCEL"
    )
    
    $installArgsString = $installArgs -join " "
    Write-Log "Starting installation with arguments: $installArgsString"
    
    try {
        # PS 5/7 compatible process start
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $Config.InstallerPath
        $processInfo.Arguments = $installArgsString
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $null = $process.Start()
        $process.WaitForExit()
        
        $exitCode = $process.ExitCode
        Write-Log "Installer exited with code: $exitCode"
        
        # Inno Setup exit codes
        if ($exitCode -ne 0) {
            $errorMsg = switch ($exitCode) {
                1 { "Setup initialization error" }
                2 { "User cancelled installation" }
                3 { "Fatal error during installation" }
                4 { "Fatal error before installation started" }
                5 { "User chose not to proceed" }
                6 { "Setup initialization error (variant)" }
                7 { "Installation aborted" }
                8 { "Disk space error" }
                default { "Unknown error code $exitCode" }
            }
            throw "Installation failed: $errorMsg (Exit Code: $exitCode)"
        }
        
        Write-Log "Installation completed successfully" -Level Success
        return $true
    }
    catch {
        Write-Log "Installation exception: $_" -Level Error
        throw
    }
}

function Test-Installation {
    Write-Log "Validating installation..."
    
    # Check for ShareX.exe in standard locations
    $exePaths = @(
        "$env:ProgramFiles\ShareX\ShareX.exe",
        "${env:ProgramFiles(x86)}\ShareX\ShareX.exe"
    )
    
    $foundPath = $null
    foreach ($path in $exePaths) {
        if (Test-Path $path) {
            $foundPath = $path
            break
        }
    }
    
    if (!$foundPath) {
        Write-Log "ShareX.exe not found in expected locations" -Level Error
        return $false
    }
    
    # Get file version
    $fileVersionInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($foundPath)
    $exeVersion = $fileVersionInfo.FileVersion
    Write-Log "ShareX validated at: $foundPath (Version: $exeVersion)" -Level Success
    
    # Capture uninstall information
    try {
        $uninstallInfoPath = Join-Path $Config.WorkRoot "uninstall-info.txt"
        $uninstallInfo = @()
        
        $regPaths = @(
            "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        
        foreach ($regPath in $regPaths) {
            if (Test-Path $regPath) {
                $apps = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
                foreach ($app in $apps) {
                    $appInfo = Get-ItemProperty -Path $app.PSPath -ErrorAction SilentlyContinue
                    if ($appInfo.DisplayName -eq $Config.AppName) {
                        $uninstallInfo += "DisplayName: $($appInfo.DisplayName)"
                        $uninstallInfo += "DisplayVersion: $($appInfo.DisplayVersion)"
                        $uninstallInfo += "Publisher: $($appInfo.Publisher)"
                        $uninstallInfo += "InstallLocation: $($appInfo.InstallLocation)"
                        $uninstallInfo += "UninstallString: $($appInfo.UninstallString)"
                        $uninstallInfo += "InstallDate: $($appInfo.InstallDate)"
                    }
                }
            }
        }
        
        if ($uninstallInfo.Count -gt 0) {
            $uninstallInfo | Out-File -FilePath $uninstallInfoPath -Force
            Write-Log "Uninstall information saved"
        }
    }
    catch {
        Write-Log "Warning: Could not capture uninstall info: $_" -Level Warning
    }
    
    return $true
}

function Remove-InstallerFile {
    try {
        if (Test-Path $Config.InstallerPath) {
            Remove-Item -Path $Config.InstallerPath -Force -ErrorAction Stop
            Write-Log "Cleaned up installer file"
        }
    }
    catch {
        Write-Log "Could not remove installer file: $_" -Level Warning
    }
}
#endregion Functions

#region Main Execution
$exitCode = 0

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "ShareX $($Config.AppVersion) Deployment" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Log "Deployment started"
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Log "Running as: $env:USERNAME on $env:COMPUTERNAME"
    
    # Step 1: Initialize environment
    if (!(Initialize-Environment)) {
        throw "Environment initialization failed"
    }
    
    # Step 2: Check if already installed
    if (Test-AlreadyInstalled) {
        Write-Log "Deployment skipped - application already installed at target version" -Level Success
        Write-Host "INTUNE_STATUS: SUCCESS - Already Installed" -ForegroundColor Green
        exit 0
    }
    
    # Step 3: Download installer
    if (!(Get-Installer)) {
        throw "Download failed"
    }
    
    # Step 4: Install application
    if (!(Install-Application)) {
        throw "Installation failed"
    }
    
    # Step 5: Validate installation
    if (!(Test-Installation)) {
        Write-Log "Post-installation validation failed" -Level Error
        Write-Host "INTUNE_STATUS: FAILED - Validation Error" -ForegroundColor Red
        $exitCode = 3
    }
    else {
        # Step 6: Cleanup
        Remove-InstallerFile
        
        Write-Log "Deployment completed successfully" -Level Success
        Write-Host "INTUNE_STATUS: SUCCESS - Installation Complete" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
        $exitCode = 0
    }
}
catch {
    $errorDetails = $_.Exception.Message
    Write-Log "DEPLOYMENT FAILED: $errorDetails" -Level Error
    Write-Host "INTUNE_STATUS: FAILED - $errorDetails" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Cyan
    
    # Return appropriate exit code based on error context
    if ($errorDetails -match "download|Download") { 
        $exitCode = 1 
        Write-Log "Exit Code: 1 (Download Failure)" -Level Error
    }
    elseif ($errorDetails -match "install|Install") { 
        $exitCode = 2 
        Write-Log "Exit Code: 2 (Installation Failure)" -Level Error
    }
    else { 
        $exitCode = 99 
        Write-Log "Exit Code: 99 (General Failure)" -Level Error
    }
}
finally {
    Write-Log "Log file location: $($Config.LogPath)"
    Write-Host "Log saved to: $($Config.LogPath)"
}

# Exit with appropriate code for Intune
exit $exitCode
#endregion Main Execution

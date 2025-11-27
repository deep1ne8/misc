#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Lacerte VC++ Runtime Repair Script - Detects, downloads, and installs required Visual C++ redistributables.

.DESCRIPTION
    Automatically detects missing VC++ runtime packages required by Lacerte and installs them silently.
    Supports VC++ 2008, 2010, 2013, and 2015-2022 (x86 versions).

.NOTES
    Version: 2.0
    Author: Earl's MSP Toolkit
    Requires: Administrator privileges
#>

[CmdletBinding()]
param()

# Configuration
$LogFile = "$env:SystemRoot\Temp\Lacerte_VC_Redistributable_Repair.log"
$TempPath = "$env:SystemRoot\Temp\LacerteVC"
$MaxRetries = 3
$TimeoutSeconds = 300

# Enhanced logging function
function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    try {
        Add-Content -Path $LogFile -Value $logEntry -ErrorAction Stop
        
        # Console output with color coding
        switch ($Level) {
            'Error'   { Write-Host $logEntry -ForegroundColor Red }
            'Warning' { Write-Host $logEntry -ForegroundColor Yellow }
            'Success' { Write-Host $logEntry -ForegroundColor Green }
            default   { Write-Host $logEntry }
        }
    }
    catch {
        Write-Warning "Failed to write to log file: $_"
    }
}

# VC++ Package definitions with improved version detection
$VCPackages = @(
    @{
        Name = "Microsoft Visual C++ 2008 x86"
        DetectVersion = "9.0.30729"
        DetectName = "Microsoft Visual C++ 2008 Redistributable"
        Url = "https://download.microsoft.com/download/5/D/8/5D8C65DF-C2CA-4628-A41D-1D324379A3DA/vcredist_x86.exe"
        FileName = "vcredist_x86_2008.exe"
        InstallArgs = @("/q", "/norestart")
        MinVersion = [version]"9.0.30729.6161"
    },
    @{
        Name = "Microsoft Visual C++ 2010 x86"
        DetectVersion = "10.0.40219"
        DetectName = "Microsoft Visual C++ 2010"
        Url = "https://download.microsoft.com/download/1/6/5/165255E7-1014-4D0A-B094-B6A430A6BFFC/vcredist_x86.exe"
        FileName = "vcredist_x86_2010.exe"
        InstallArgs = @("/q", "/norestart")
        MinVersion = [version]"10.0.40219.325"
    },
    @{
        Name = "Microsoft Visual C++ 2013 x86"
        DetectVersion = "12.0.40664"
        DetectName = "Microsoft Visual C++ 2013"
        Url = "https://aka.ms/highdpimfc2013x86enu"
        FileName = "vcredist_x86_2013.exe"
        InstallArgs = @("/install", "/quiet", "/norestart")
        MinVersion = [version]"12.0.40664.0"
    },
    @{
        Name = "Microsoft Visual C++ 2015-2022 x86"
        DetectVersion = "14."
        DetectName = "Microsoft Visual C++ 2015-2022 Redistributable"
        Url = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
        FileName = "vc_redist_x86_2015_2022.exe"
        InstallArgs = @("/install", "/quiet", "/norestart")
        MinVersion = [version]"14.38.33130.0"
    }
)

# Enhanced version detection function
function Get-InstalledVCVersion {
    param(
        [string]$DetectPattern,
        [string]$DetectName
    )
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    
    $installedVersions = @()
    
    foreach ($path in $registryPaths) {
        try {
            $items = Get-ItemProperty $path -ErrorAction SilentlyContinue | 
                Where-Object { 
                    $_.DisplayName -like "*$DetectName*" -and 
                    $_.DisplayVersion 
                }
            
            foreach ($item in $items) {
                if ($item.DisplayVersion -like "$DetectPattern*") {
                    $installedVersions += [PSCustomObject]@{
                        Name = $item.DisplayName
                        Version = $item.DisplayVersion
                        Publisher = $item.Publisher
                    }
                }
            }
        }
        catch {
            Write-Log "Error checking registry path $path : $_" -Level Warning
        }
    }
    
    return $installedVersions
}

# Download function with retry logic
function Get-VCPackage {
    param(
        [string]$Url,
        [string]$Destination,
        [int]$RetryCount = $MaxRetries
    )
    
    $attempt = 0
    $success = $false
    
    while (-not $success -and $attempt -lt $RetryCount) {
        $attempt++
        try {
            Write-Log "Download attempt $attempt of $RetryCount for $Url"
            
            # Use BITS transfer for better reliability
            if (Get-Command Start-BitsTransfer -ErrorAction SilentlyContinue) {
                Start-BitsTransfer -Source $Url -Destination $Destination -ErrorAction Stop
            }
            else {
                # Fallback to Invoke-WebRequest
                $ProgressPreference = 'SilentlyContinue'
                Invoke-WebRequest -Uri $Url -OutFile $Destination -UseBasicParsing -TimeoutSec $TimeoutSeconds -ErrorAction Stop
                $ProgressPreference = 'Continue'
            }
            
            if (Test-Path $Destination) {
                $fileSize = (Get-Item $Destination).Length
                Write-Log "Download successful. File size: $([math]::Round($fileSize/1MB, 2)) MB"
                $success = $true
            }
        }
        catch {
            Write-Log "Download attempt $attempt failed: $_" -Level Warning
            if ($attempt -lt $RetryCount) {
                Start-Sleep -Seconds (5 * $attempt)
            }
        }
    }
    
    return $success
}

# Installation function with validation
function Install-VCPackage {
    param(
        [string]$Path,
        [array]$Arguments,
        [string]$PackageName
    )
    
    try {
        Write-Log "Installing $PackageName..."
        
        $process = Start-Process -FilePath $Path -ArgumentList $Arguments -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
            Write-Log "$PackageName installed successfully. Exit code: $($process.ExitCode)" -Level Success
            return $true
        }
        elseif ($process.ExitCode -eq 1638) {
            Write-Log "$PackageName already installed (Exit code: 1638)" -Level Info
            return $true
        }
        else {
            Write-Log "$PackageName installation failed. Exit code: $($process.ExitCode)" -Level Error
            return $false
        }
    }
    catch {
        Write-Log "Exception during installation of $PackageName : $_" -Level Error
        return $false
    }
}

# Main execution
try {
    Write-Log "========================================" -Level Info
    Write-Log "Starting Lacerte VC++ Runtime Repair" -Level Info
    Write-Log "========================================" -Level Info
    
    # Verify administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Log "Script must be run as Administrator!" -Level Error
        exit 1
    }
    
    # Create temp directory
    if (-not (Test-Path $TempPath)) {
        New-Item -ItemType Directory -Path $TempPath -Force | Out-Null
        Write-Log "Created temporary directory: $TempPath"
    }
    
    $packagesInstalled = 0
    $packagesFailed = 0
    $packagesSkipped = 0
    
    foreach ($package in $VCPackages) {
        Write-Log "`nProcessing: $($package.Name)" -Level Info
        
        # Check if already installed
        $installedVersions = Get-InstalledVCVersion -DetectPattern $package.DetectVersion -DetectName $package.DetectName
        
        if ($installedVersions.Count -gt 0) {
            Write-Log "$($package.Name) detected: Version $($installedVersions[0].Version)" -Level Success
            $packagesSkipped++
            continue
        }
        
        Write-Log "$($package.Name) not found. Attempting installation..."
        
        # Download package
        $downloadPath = Join-Path $TempPath $package.FileName
        $downloadSuccess = Get-VCPackage -Url $package.Url -Destination $downloadPath
        
        if (-not $downloadSuccess) {
            Write-Log "Failed to download $($package.Name) after $MaxRetries attempts" -Level Error
            $packagesFailed++
            continue
        }
        
        # Install package
        $installSuccess = Install-VCPackage -Path $downloadPath -Arguments $package.InstallArgs -PackageName $package.Name
        
        if ($installSuccess) {
            $packagesInstalled++
            
            # Verify installation
            Start-Sleep -Seconds 2
            $verifyInstall = Get-InstalledVCVersion -DetectPattern $package.DetectVersion -DetectName $package.DetectName
            if ($verifyInstall.Count -gt 0) {
                Write-Log "Installation verified: $($package.Name) version $($verifyInstall[0].Version)" -Level Success
            }
        }
        else {
            $packagesFailed++
        }
        
        # Cleanup installer
        try {
            Remove-Item $downloadPath -Force -ErrorAction SilentlyContinue
        }
        catch {
            Write-Log "Could not remove installer file: $downloadPath" -Level Warning
        }
    }
    
    # Summary
    Write-Log "`n========================================" -Level Info
    Write-Log "VC++ Runtime Repair Summary" -Level Info
    Write-Log "========================================" -Level Info
    Write-Log "Packages already installed: $packagesSkipped" -Level Info
    Write-Log "Packages newly installed: $packagesInstalled" -Level Success
    Write-Log "Packages failed: $packagesFailed" -Level $(if ($packagesFailed -gt 0) { "Error" } else { "Info" })
    Write-Log "========================================" -Level Info
    
    # Cleanup temp directory
    try {
        if (Test-Path $TempPath) {
            Remove-Item $TempPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Cleaned up temporary directory"
        }
    }
    catch {
        Write-Log "Could not remove temp directory: $_" -Level Warning
    }
    
    if ($packagesFailed -eq 0) {
        Write-Log "`nAll required VC++ packages are installed successfully!" -Level Success
        exit 0
    }
    else {
        Write-Log "`nSome packages failed to install. Review log for details." -Level Warning
        exit 1
    }
}
catch {
    Write-Log "Critical error during execution: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    exit 1
}

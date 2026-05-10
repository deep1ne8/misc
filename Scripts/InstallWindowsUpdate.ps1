#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Windows Update Management and Installation Script
.DESCRIPTION
    This script handles Windows Update configuration, BitLocker management, and update installation with detailed logging
.NOTES
    Version: 2.0
    Author: Earl Daniels
    Last Updated: 2025-02-26
#>

function Write-CustomLog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'ERROR')]
        [string]$Level = 'INFO',
        
        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )
    
    # Define log path
    $logPath = "C:\Windows\Temp\WinUpgradeLog.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    
    # Format message
    $logMessage = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logPath -Value $logMessage
    
    # Only write to console if not suppressed
    if (-not $NoConsole) {
        # Set appropriate color based on level
        $color = switch ($Level) {
            'INFO'    { 'White' }
            'SUCCESS' { 'Green' }
            'WARNING' { 'Yellow' }
            'ERROR'   { 'Red' }
            default   { 'White' }
        }
        
        Write-Host $logMessage -ForegroundColor $color
    }
}

function Show-BannerMessage {
    param (
        [string]$Message
    )
    
    $bannerWidth = 60
    $padding = [Math]::Max(0, ($bannerWidth - $Message.Length - 2) / 2)
    $leftPad = [Math]::Floor($padding)
    $rightPad = [Math]::Ceiling($padding)
    
    Write-Host ""
    Write-Host ("=" * $bannerWidth) -ForegroundColor Cyan
    Write-Host ("|" + (" " * $leftPad) + $Message + (" " * $rightPad) + "|") -ForegroundColor Cyan
    Write-Host ("=" * $bannerWidth) -ForegroundColor Cyan
    Write-Host ""
}

function Remove-WindowsUpdatePolicies {
    Show-BannerMessage "Removing Windows Update Policies"
    
    $paths = @(
        "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate", 
        "HKCU:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
    )
    
    foreach ($path in $paths) {
        if (Test-Path $path) {
            try {
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Write-CustomLog -Message "Successfully removed: $path" -Level 'SUCCESS'
            }
            catch {
                Write-CustomLog -Message "Failed to remove $path : $_" -Level 'ERROR'
            }
        }
        else {
            Write-CustomLog -Message "Path not found (skipping): $path" -Level 'INFO'
        }
    }
}

function Install-RequiredModules {
    Show-BannerMessage "Installing Required Modules"
    
    try {
        # Set execution policy
        Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser -Force -ErrorAction Stop
        Write-CustomLog -Message "Execution policy set to Unrestricted for current user" -Level 'SUCCESS'
        
        # Install NuGet if needed
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Write-CustomLog -Message "Installing NuGet package provider..." -Level 'INFO'
            Install-PackageProvider -Name NuGet -Force -ErrorAction Stop | Out-Null
            Write-CustomLog -Message "NuGet package provider installed successfully" -Level 'SUCCESS'
        }
        else {
            Write-CustomLog -Message "NuGet package provider already installed" -Level 'INFO'
        }
        
        # Install PSWindowsUpdate if needed
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-CustomLog -Message "Installing PSWindowsUpdate module..." -Level 'INFO'
            Install-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
            Write-CustomLog -Message "PSWindowsUpdate module installed successfully" -Level 'SUCCESS'
        }
        else {
            Write-CustomLog -Message "PSWindowsUpdate module already installed" -Level 'INFO'
        }
        
        # Import module
        Import-Module -Name PSWindowsUpdate -Force -ErrorAction Stop
        Write-CustomLog -Message "PSWindowsUpdate module imported successfully" -Level 'SUCCESS'
        
        return $true
    }
    catch {
        Write-CustomLog -Message "Failed to install required modules: $_" -Level 'ERROR'
        return $false
    }
}

function Reset-WindowsUpdateComponents {
    Show-BannerMessage "Cleaning Windows Update Components"
    
    try {
        Write-CustomLog -Message "Stopping Windows Update related services..." -Level 'INFO'
        
        $updateServices = @("wuauserv", "cryptSvc", "bits", "msiserver")
        foreach ($service in $updateServices) {
            try {
                Stop-Service -Name $service -Force -ErrorAction SilentlyContinue
                Write-CustomLog -Message "Service stopped: $service" -Level 'INFO' -NoConsole
            }
            catch {
                Write-CustomLog -Message "Could not stop service $service" -Level 'WARNING' -NoConsole
            }
        }
        
        Import-Module -Name PSWindowsUpdate -Force
        Reset-WUComponents -Verbose
        
        Write-CustomLog -Message "Windows Update components reset successfully" -Level 'SUCCESS'
        return $true
    }
    catch {
        Write-CustomLog -Message "Failed to reset Windows Update components: $_" -Level 'ERROR'
        return $false
    }
}

function Suspend-BitLockerEncryption {
    Show-BannerMessage "Managing BitLocker Encryption"
    
    try {
        # Check if BitLocker module is available
        if (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue) {
            $bitLockerVolume = Get-BitLockerVolume -MountPoint C: -ErrorAction SilentlyContinue
            
            if ($bitLockerVolume) {
                $bitLockerStatus = $bitLockerVolume.VolumeStatus
                
                if ($bitLockerStatus -eq 'FullyEncrypted' -or $bitLockerStatus -eq 'EncryptionInProgress') {
                    Write-CustomLog -Message "BitLocker is enabled on C:\ - Suspending protection" -Level 'INFO'
                    Suspend-BitLocker -MountPoint C: -RebootCount 1 -ErrorAction Stop
                    Write-CustomLog -Message "BitLocker protection suspended successfully for 1 reboot" -Level 'SUCCESS'
                }
                else {
                    Write-CustomLog -Message "BitLocker is not in a state requiring suspension (Status: $bitLockerStatus)" -Level 'INFO'
                }
            }
            else {
                Write-CustomLog -Message "BitLocker is not enabled on C:\" -Level 'INFO'
            }
        }
        else {
            Write-CustomLog -Message "BitLocker cmdlets not available on this system" -Level 'WARNING'
        }
        
        return $true
    }
    catch {
        Write-CustomLog -Message "Failed to manage BitLocker: $_" -Level 'ERROR'
        return $false
    }
}

function Install-WindowsUpdates {
    Show-BannerMessage "Installing Windows Updates"
    
    try {
        Write-CustomLog -Message "Scanning for available updates..." -Level 'INFO'
        $availableUpdates = Get-WindowsUpdate -ErrorAction Stop
        
        if ($availableUpdates.Count -eq 0) {
            Write-CustomLog -Message "No updates available for installation" -Level 'INFO'
        }
        else {
            Write-CustomLog -Message "Found $($availableUpdates.Count) update(s) to install" -Level 'INFO'
            Write-CustomLog -Message "Beginning update installation - this may take some time..." -Level 'INFO'
            
            # Redirect verbose output to a temporary file for logging purposes
            $tempLogFile = [System.IO.Path]::GetTempFileName()
            $null = Get-WindowsUpdate -AcceptAll -Install -Verbose 4> $tempLogFile
            
            # Read the verbose output and log it properly
            Get-Content $tempLogFile | ForEach-Object {
                Write-CustomLog -Message $_ -Level 'INFO' -NoConsole
            }
            
            # Clean up
            Remove-Item $tempLogFile -Force -ErrorAction SilentlyContinue
            
            Write-CustomLog -Message "Windows Updates installed successfully" -Level 'SUCCESS'
        }
        
        return $true
    }
    catch {
        Write-CustomLog -Message "Failed to install Windows Updates: $_" -Level 'ERROR'
        return $false
    }
}

function Set-WindowsUpdateConfiguration {
    Show-BannerMessage "Configuring Windows Update Settings"
    
    try {
        # Configure automatic update settings
        $auPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update"
        
        if (-not (Test-Path $auPath)) {
            New-Item -Path $auPath -Force | Out-Null
            Write-CustomLog -Message "Created Windows Update Auto Update registry path" -Level 'INFO'
        }
        
        # AUOptions = 4 means "Auto download and schedule the install"
        Set-ItemProperty -Path $auPath -Name "AUOptions" -Value 4 -Type DWord -Force
        Write-CustomLog -Message "Windows Update set to automatically download and install updates" -Level 'SUCCESS'
        
        # Enable Windows Update service and set to automatic startup
        Set-Service -Name wuauserv -StartupType Automatic
        Start-Service -Name wuauserv
        Write-CustomLog -Message "Windows Update service enabled and set to automatic startup" -Level 'SUCCESS'
        
        return $true
    }
    catch {
        Write-CustomLog -Message "Failed to configure Windows Update settings: $_" -Level 'ERROR'
        return $false
    }
}

function Invoke-UpdateProcess {
    # Create log file if it doesn't exist
    $logPath = "C:\Windows\Temp\WinUpgradeLog.txt"
    if (-not (Test-Path $logPath)) {
        New-Item -Path $logPath -ItemType File -Force | Out-Null
    }
    
    # Script starting banner
    Show-BannerMessage "Windows Update Management Script"
    Write-CustomLog -Message "Script execution started" -Level 'INFO'
    
    # Step 1: Remove existing Windows Update policies
    $step1 = Remove-WindowsUpdatePolicies
    
    # Step 2: Install required modules
    $step2 = Install-RequiredModules
    if (-not $step2) {
        Write-CustomLog -Message "Terminating script due to module installation failure" -Level 'ERROR'
        return
    }
    
    # Step 3: Reset Windows Update components
    $step3 = Reset-WindowsUpdateComponents
    
    # Step 4: Manage BitLocker
    $step4 = Suspend-BitLockerEncryption
    
    # Step 5: Install Windows Updates
    $step5 = Install-WindowsUpdates
    
    # Step 6: Configure Windows Update settings
    $step6 = Set-WindowsUpdateConfiguration
    
    # Script completion
    Show-BannerMessage "Windows Update Process Complete"
    
    # Summary
    $successCount = @($step1, $step2, $step3, $step4, $step5, $step6).Where({$_ -eq $true}).Count
    $totalSteps = 6
    
    Write-CustomLog -Message "Successfully completed $successCount out of $totalSteps steps" -Level 'INFO'
    Write-CustomLog -Message "Script execution completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level 'SUCCESS'
    
    return $true
}

# Execute the main function
Invoke-UpdateProcess
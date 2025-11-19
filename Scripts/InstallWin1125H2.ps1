#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Installs Windows 11 25H2 update
.DESCRIPTION
    Searches for and installs Windows 11 version 25H2 update using PSWindowsUpdate module
#>

[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARN"){"Yellow"}else{"Green"})
}

try {
    Write-Log "Starting Windows 11 25H2 update installation"
    
    # Check Windows version
    $osInfo = Get-CimInstance Win32_OperatingSystem
    Write-Log "Current OS: $($osInfo.Caption) Build $($osInfo.BuildNumber)"
    
    if ($osInfo.Caption -notmatch "Windows 11") {
        throw "This script requires Windows 11"
    }
    
    # Install NuGet provider if needed
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Write-Log "Installing NuGet provider"
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope AllUsers | Out-Null
    }
    
    # Install PSWindowsUpdate module
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Write-Log "Installing PSWindowsUpdate module"
        Install-Module -Name PSWindowsUpdate -Force -Scope AllUsers -AllowClobber
    }
    
    Import-Module PSWindowsUpdate -Force
    Write-Log "PSWindowsUpdate module loaded"
    
    # Search for Windows 11 25H2 update
    Write-Log "Searching for Windows 11 25H2 feature update..."
    $updates = Get-WindowsUpdate -MicrosoftUpdate -Title "*Windows 11*25H2*" -Verbose:$false
    
    if (-not $updates) {
        Write-Log "Searching alternative patterns..."
        $updates = Get-WindowsUpdate -MicrosoftUpdate | Where-Object { 
            $_.Title -match "Windows 11.*version 25H2" -or 
            $_.Title -match "Feature update.*Windows 11.*25H2" 
        }
    }
    
    if (-not $updates) {
        Write-Log "No Windows 11 25H2 update found. Checking all available updates:" "WARN"
        Get-WindowsUpdate -MicrosoftUpdate | Select-Object Title, KB, Size | Format-Table -AutoSize
        throw "Windows 11 25H2 update not available"
    }
    
    Write-Log "Found update(s):"
    $updates | ForEach-Object { Write-Log "  - $($_.Title) (KB$($_.KB)) - $([math]::Round($_.Size/1MB, 2)) MB" }
    
    # Install update
    Write-Log "Installing Windows 11 25H2 update (this may take a while)..."
    $result = Install-WindowsUpdate -MicrosoftUpdate -Title "*Windows 11*25H2*" -AcceptAll -IgnoreReboot -Verbose:$false
    
    if ($result) {
        Write-Log "Update installation completed successfully"
        Write-Log "Result: $($result.Result)"
        
        if ($result.RebootRequired) {
            Write-Log "SYSTEM REBOOT REQUIRED" "WARN"
            $reboot = Read-Host "Reboot now? (Y/N)"
            if ($reboot -eq 'Y') {
                Write-Log "Initiating restart in 30 seconds..."
                shutdown /r /t 30 /c "Restarting to complete Windows 11 25H2 update installation"
            }
        }
    }
    
    Write-Log "Process completed"
    
} catch {
    Write-Log "ERROR: $($_.Exception.Message)" "ERROR"
    exit 1
}

<#
.SYNOPSIS
    Complete Office 365 removal script with blocker resolution
.DESCRIPTION
    Removes all Office 365 versions, remnants, and resolves common uninstallation blockers
.NOTES
    Author: MSP Automation
    Requires: Administrator privileges
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$RemoveUserData,
    [switch]$Force
)

$ErrorActionPreference = "Continue"
$ProgressPreference = "SilentlyContinue"

$LogPath = "$env:ProgramData\OfficeRemoval"
$LogFile = "$LogPath\OfficeRemoval_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

if (!(Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogMessage
    Write-Host $LogMessage -ForegroundColor $(if($Level -eq "ERROR"){"Red"}elseif($Level -eq "WARNING"){"Yellow"}else{"Green"})
}

function Stop-OfficeProcesses {
    Write-Log "Stopping Office processes..."
    $officeProcesses = @(
        "winword", "excel", "powerpnt", "outlook", "onenote", "msaccess", 
        "mspub", "lync", "teams", "onedrive", "officeclicktorun", 
        "integrator", "firstrun", "setup", "appvshnotify"
    )
    
    foreach ($proc in $officeProcesses) {
        Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 3
}

function Stop-OfficeServices {
    Write-Log "Stopping Office services..."
    $services = @(
        "ClickToRunSvc",
        "OfficeSvc",
        "Office 365 Service",
        "Microsoft Office Click-to-Run Service"
    )
    
    foreach ($svc in $services) {
        $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($service) {
            try {
                Stop-Service -Name $svc -Force -ErrorAction Stop
                Set-Service -Name $svc -StartupType Disabled -ErrorAction Stop
                Write-Log "Stopped and disabled service: $svc"
            } catch {
                Write-Log "Failed to stop service: $svc - $($_.Exception.Message)" "WARNING"
            }
        }
    }
}

function Remove-OfficeScheduledTasks {
    Write-Log "Removing Office scheduled tasks..."
    Get-ScheduledTask | Where-Object {$_.TaskName -like "*Office*"} | Unregister-ScheduledTask -Confirm:$false -ErrorAction SilentlyContinue
}

function Uninstall-OfficeProducts {
    Write-Log "Uninstalling Office products via WMI..."
    
    $products = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -like "*Microsoft 365*" -or 
        $_.Name -like "*Office 365*" -or 
        $_.Name -like "*Microsoft Office*" -or
        $_.Name -like "*Office 16*" -or
        $_.Name -like "*Office 19*"
    }
    
    foreach ($product in $products) {
        Write-Log "Uninstalling: $($product.Name)"
        try {
            $product.Uninstall() | Out-Null
            Write-Log "Successfully uninstalled: $($product.Name)"
        } catch {
            Write-Log "Failed to uninstall: $($product.Name) - $($_.Exception.Message)" "ERROR"
        }
    }
}

function Remove-OfficeClickToRun {
    Write-Log "Removing Click-to-Run installation..."
    
    $c2rPaths = @(
        "${env:ProgramFiles}\Microsoft Office\root\integration\integrator.exe",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\integration\integrator.exe"
    )
    
    foreach ($path in $c2rPaths) {
        if (Test-Path $path) {
            Write-Log "Found integrator at: $path"
            try {
                Start-Process -FilePath $path -ArgumentList "/U" -Wait -NoNewWindow -ErrorAction Stop
                Write-Log "Executed integrator /U successfully"
            } catch {
                Write-Log "Failed to execute integrator: $($_.Exception.Message)" "WARNING"
            }
        }
    }
    
    $setupPaths = @(
        "${env:ProgramFiles}\Microsoft Office\Office16\SETUP.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\SETUP.EXE"
    )
    
    foreach ($path in $setupPaths) {
        if (Test-Path $path) {
            Write-Log "Running setup cleanup from: $path"
            try {
                Start-Process -FilePath $path -ArgumentList "/uninstall ProPlus /config `"$env:Temp\uninstall.xml`"" -Wait -NoNewWindow -ErrorAction Stop
            } catch {
                Write-Log "Setup cleanup attempt failed: $($_.Exception.Message)" "WARNING"
            }
        }
    }
}

function Remove-OfficeRegistryKeys {
    Write-Log "Removing Office registry keys..."
    
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\O365*",
        "HKCU:\SOFTWARE\Microsoft\Office",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365*"
    )
    
    foreach ($regPath in $regPaths) {
        if (Test-Path $regPath) {
            try {
                Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                Write-Log "Removed registry key: $regPath"
            } catch {
                Write-Log "Failed to remove registry key: $regPath - $($_.Exception.Message)" "WARNING"
            }
        }
    }
}

function Remove-OfficeFolders {
    Write-Log "Removing Office installation folders..."
    
    $folders = @(
        "${env:ProgramFiles}\Microsoft Office",
        "${env:ProgramFiles(x86)}\Microsoft Office",
        "${env:ProgramFiles}\Microsoft Office 15",
        "${env:ProgramFiles(x86)}\Microsoft Office 15",
        "${env:ProgramFiles}\Microsoft Office 16",
        "${env:ProgramFiles(x86)}\Microsoft Office 16",
        "${env:ProgramData}\Microsoft\Office",
        "${env:ProgramData}\Microsoft\ClickToRun",
        "${env:ProgramData}\Microsoft\OfficeSoftwareProtectionPlatform"
    )
    
    if ($RemoveUserData) {
        $folders += "${env:LOCALAPPDATA}\Microsoft\Office"
        $folders += "${env:APPDATA}\Microsoft\Office"
        $folders += "${env:APPDATA}\Microsoft\Templates"
    }
    
    foreach ($folder in $folders) {
        if (Test-Path $folder) {
            try {
                Remove-Item -Path $folder -Recurse -Force -ErrorAction Stop
                Write-Log "Removed folder: $folder"
            } catch {
                Write-Log "Failed to remove folder: $folder - $($_.Exception.Message)" "WARNING"
                if ($Force) {
                    cmd /c "rd /s /q `"$folder`"" 2>&1 | Out-Null
                }
            }
        }
    }
}

function Remove-OfficeCache {
    Write-Log "Clearing Office cache and temporary files..."
    
    $cachePaths = @(
        "${env:LOCALAPPDATA}\Microsoft\Office\16.0",
        "${env:LOCALAPPDATA}\Microsoft\Office\15.0",
        "${env:TEMP}\*.tmp"
    )
    
    foreach ($cache in $cachePaths) {
        if (Test-Path $cache) {
            Remove-Item -Path $cache -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function Use-OfficeScrubber {
    Write-Log "Attempting to use official Office removal tool..."
    
    $scrubberUrl = "https://aka.ms/SaRA_CommandLineVersion"
    $scrubberPath = "$env:TEMP\OfficeUninstaller.exe"
    
    try {
        Write-Log "Downloading Office scrubber tool..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $scrubberUrl -OutFile $scrubberPath -UseBasicParsing -ErrorAction Stop
        
        if (Test-Path $scrubberPath) {
            Write-Log "Running Office scrubber..."
            Start-Process -FilePath $scrubberPath -ArgumentList "-S -O365 -Q" -Wait -NoNewWindow
            Remove-Item -Path $scrubberPath -Force -ErrorAction SilentlyContinue
            Write-Log "Office scrubber completed"
        }
    } catch {
        Write-Log "Could not download/run Office scrubber: $($_.Exception.Message)" "WARNING"
    }
}

function Test-OfficeRemaining {
    Write-Log "Checking for remaining Office components..."
    
    $remaining = @()
    
    $checkPaths = @(
        "${env:ProgramFiles}\Microsoft Office",
        "${env:ProgramFiles(x86)}\Microsoft Office"
    )
    
    foreach ($path in $checkPaths) {
        if (Test-Path $path) {
            $remaining += $path
        }
    }
    
    $regCheck = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
                Where-Object {$_.DisplayName -like "*Office*" -or $_.DisplayName -like "*365*"}
    
    if ($regCheck) {
        $remaining += "Registry entries found"
    }
    
    if ($remaining.Count -gt 0) {
        Write-Log "Remaining Office components detected:" "WARNING"
        $remaining | ForEach-Object { Write-Log "  - $_" "WARNING" }
        return $false
    } else {
        Write-Log "No Office components detected - removal successful!"
        return $true
    }
}

# Main execution
try {
    Write-Log "=== Office 365 Complete Removal Started ==="
    Write-Log "Script version: 1.0"
    Write-Log "Computer: $env:COMPUTERNAME"
    Write-Log "User: $env:USERNAME"
    
    Stop-OfficeProcesses
    Stop-OfficeServices
    Remove-OfficeScheduledTasks
    Uninstall-OfficeProducts
    Remove-OfficeClickToRun
    
    Stop-OfficeProcesses
    
    Remove-OfficeRegistryKeys
    Remove-OfficeFolders
    Remove-OfficeCache
    
    Use-OfficeScrubber
    
    $cleanRemoval = Test-OfficeRemaining
    
    Write-Log "=== Office 365 Removal Completed ==="
    Write-Log "Log file saved: $LogFile"
    
    if ($cleanRemoval) {
        Write-Log "SUCCESS: Office has been completely removed" "INFO"
        Write-Host "`nA system restart is recommended to complete the cleanup." -ForegroundColor Cyan
    } else {
        Write-Log "WARNING: Some Office components may remain - manual cleanup may be required" "WARNING"
        Write-Host "`nSome components remain. Review log file for details: $LogFile" -ForegroundColor Yellow
    }
    
} catch {
    Write-Log "Critical error during removal: $($_.Exception.Message)" "ERROR"
    throw
}

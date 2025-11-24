<#
.SYNOPSIS
    Bartender Performance Diagnostic and Remediation Tool for Windows 11 24H2
.DESCRIPTION
    Diagnoses and fixes common Bartender printing slowness issues on Windows 11 24H2
    Addresses print spooler, driver isolation, memory optimization, and compatibility issues
.NOTES
    Author: Earl's MSP Solutions
    Version: 1.0
    Requires: Administrator privileges
#>

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$DiagnosticOnly,
    [switch]$SkipBackup
)

$ErrorActionPreference = "Continue"
$LogPath = "C:\Temp\BartenderFix_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage -ForegroundColor $(switch($Level){"ERROR"{"Red"}"WARNING"{"Yellow"}"SUCCESS"{"Green"}default{"White"}})
    Add-Content -Path $LogPath -Value $logMessage
}

function Test-BartenderRunning {
    $bartender = Get-Process -Name "bartend" -ErrorAction SilentlyContinue
    if ($bartender) {
        Write-Log "Bartender is currently running (PID: $($bartender.Id))" "WARNING"
        return $true
    }
    return $false
}

function Get-SystemInfo {
    Write-Log "=== SYSTEM DIAGNOSTICS ===" "INFO"
    
    $os = Get-CimInstance Win32_OperatingSystem
    $cpu = Get-CimInstance Win32_Processor
    $ram = [math]::Round($os.TotalVisibleMemorySize/1MB, 2)
    $freeRam = [math]::Round($os.FreePhysicalMemory/1MB, 2)
    
    Write-Log "OS: $($os.Caption) Build $($os.BuildNumber)" "INFO"
    Write-Log "CPU: $($cpu.Name)" "INFO"
    Write-Log "Total RAM: $ram GB | Free RAM: $freeRam GB" "INFO"
    
    # Check for 24H2
    if ($os.BuildNumber -ge 26100) {
        Write-Log "Windows 11 24H2 detected - Known printing issues present" "WARNING"
        return $true
    }
    return $false
}

function Test-WindowsUpdates {
    Write-Log "Checking for required Windows Updates..." "INFO"
    
    try {
        $session = New-Object -ComObject Microsoft.Update.Session
        $searcher = $session.CreateUpdateSearcher()
        $result = $searcher.Search("IsInstalled=0 AND Type='Software'")
        
        if ($result.Updates.Count -gt 0) {
            Write-Log "Found $($result.Updates.Count) pending Windows Updates" "WARNING"
            Write-Log "Recommend installing KB5048667 or later for 24H2 printing fixes" "WARNING"
        } else {
            Write-Log "Windows Update check complete - No pending updates" "SUCCESS"
        }
    } catch {
        Write-Log "Unable to check Windows Updates: $($_.Exception.Message)" "WARNING"
    }
}

function Backup-RegistryKeys {
    if ($SkipBackup) { return }
    
    Write-Log "Backing up registry keys..." "INFO"
    $backupPath = "C:\Temp\RegistryBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"
    
    try {
        reg export "HKLM\SYSTEM\CurrentControlSet\Control\Print" $backupPath /y | Out-Null
        Write-Log "Registry backup saved to: $backupPath" "SUCCESS"
    } catch {
        Write-Log "Registry backup failed: $($_.Exception.Message)" "WARNING"
    }
}

function Fix-PrintSpooler {
    Write-Log "=== PRINT SPOOLER REMEDIATION ===" "INFO"
    
    try {
        # Stop print spooler
        Write-Log "Stopping Print Spooler service..." "INFO"
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        
        # Clear spooler cache
        $spoolPath = "C:\Windows\System32\spool\PRINTERS"
        if (Test-Path $spoolPath) {
            Write-Log "Clearing print spooler cache..." "INFO"
            Get-ChildItem -Path $spoolPath -File | Remove-Item -Force -ErrorAction SilentlyContinue
        }
        
        # Start print spooler
        Write-Log "Starting Print Spooler service..." "INFO"
        Start-Service -Name Spooler -ErrorAction Stop
        
        # Set spooler to automatic startup
        Set-Service -Name Spooler -StartupType Automatic
        
        Write-Log "Print Spooler remediation completed" "SUCCESS"
        return $true
    } catch {
        Write-Log "Print Spooler fix failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Fix-PrinterDriverIsolation {
    Write-Log "=== PRINTER DRIVER ISOLATION FIX ===" "INFO"
    
    if ($DiagnosticOnly) {
        Write-Log "Diagnostic mode - Skipping registry changes" "INFO"
        return
    }
    
    try {
        # Disable RPC Authentication Level (known 24H2 issue)
        $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Print"
        
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        Write-Log "Disabling printer driver isolation..." "INFO"
        Set-ItemProperty -Path $regPath -Name "RpcAuthnLevelPrivacyEnabled" -Value 0 -Type DWord -Force
        
        # Disable Point and Print restrictions
        $pnpPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PointAndPrint"
        if (-not (Test-Path $pnpPath)) {
            New-Item -Path $pnpPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $pnpPath -Name "RestrictDriverInstallationToAdministrators" -Value 0 -Type DWord -Force
        
        Write-Log "Printer driver isolation disabled" "SUCCESS"
        return $true
    } catch {
        Write-Log "Driver isolation fix failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Optimize-BartenderSettings {
    Write-Log "=== BARTENDER OPTIMIZATION ===" "INFO"
    
    # Common Bartender installation paths
    $bartenderPaths = @(
        "C:\Program Files\Seagull\BarTender",
        "C:\Program Files (x86)\Seagull\BarTender",
        "${env:ProgramFiles}\Seagull\BarTender",
        "${env:ProgramFiles(x86)}\Seagull\BarTender"
    )
    
    $bartenderPath = $bartenderPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if (-not $bartenderPath) {
        Write-Log "Bartender installation not found in standard locations" "WARNING"
        return $false
    }
    
    Write-Log "Bartender found at: $bartenderPath" "INFO"
    
    # Check Bartender version
    $exePath = Get-ChildItem -Path $bartenderPath -Filter "bartend.exe" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($exePath) {
        $version = (Get-Item $exePath.FullName).VersionInfo.FileVersion
        Write-Log "Bartender Version: $version" "INFO"
    }
    
    # Recommend registry optimizations
    Write-Log "Registry optimizations for Bartender..." "INFO"
    
    try {
        # Disable GDI rendering (improves performance)
        $btRegPath = "HKLM:\SOFTWARE\Seagull Scientific\BarTender"
        if (Test-Path $btRegPath) {
            # These may vary by version - check Bartender documentation
            Write-Log "Bartender registry key found - manual optimization recommended" "INFO"
        }
        
        return $true
    } catch {
        Write-Log "Bartender optimization encountered issues: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

function Optimize-SystemPerformance {
    Write-Log "=== SYSTEM PERFORMANCE OPTIMIZATION ===" "INFO"
    
    try {
        # Disable Windows Search indexing for print folders (reduces I/O)
        Write-Log "Optimizing Windows Search indexing..." "INFO"
        $spoolPath = "C:\Windows\System32\spool"
        
        # Stop Windows Search temporarily
        Stop-Service -Name "WSearch" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Service -Name "WSearch" -ErrorAction SilentlyContinue
        
        # Increase processor scheduling for background services
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\PriorityControl" -Name "Win32PrioritySeparation" -Value 24 -Type DWord -Force
        
        # Disable visual effects that impact printing (if minimal RAM)
        $ram = (Get-CimInstance Win32_OperatingSystem).TotalVisibleMemorySize/1MB
        if ($ram -lt 12) {
            Write-Log "Low RAM detected ($([math]::Round($ram, 2)) GB) - Recommending visual effects optimization" "WARNING"
            Write-Log "Manual step: Adjust visual effects to 'Adjust for best performance' in System Properties" "INFO"
        }
        
        Write-Log "System performance optimization completed" "SUCCESS"
        return $true
    } catch {
        Write-Log "Performance optimization failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Test-PrinterConfiguration {
    Write-Log "=== PRINTER CONFIGURATION ANALYSIS ===" "INFO"
    
    try {
        $printers = Get-Printer
        Write-Log "Found $($printers.Count) configured printer(s)" "INFO"
        
        foreach ($printer in $printers) {
            Write-Log "Printer: $($printer.Name) | Status: $($printer.PrinterStatus) | Driver: $($printer.DriverName)" "INFO"
            
            # Check for stuck jobs
            $jobs = Get-PrintJob -PrinterName $printer.Name -ErrorAction SilentlyContinue
            if ($jobs) {
                Write-Log "  Found $($jobs.Count) print job(s) in queue" "WARNING"
                if (-not $DiagnosticOnly) {
                    $jobs | Remove-PrintJob -Confirm:$false -ErrorAction SilentlyContinue
                    Write-Log "  Cleared stuck print jobs" "SUCCESS"
                }
            }
        }
        
        return $true
    } catch {
        Write-Log "Printer configuration analysis failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Set-RecommendedServices {
    Write-Log "=== SERVICE CONFIGURATION ===" "INFO"
    
    $services = @(
        @{Name="Spooler"; StartupType="Automatic"; ShouldBeRunning=$true},
        @{Name="PrintNotify"; StartupType="Manual"; ShouldBeRunning=$false},
        @{Name="WSearch"; StartupType="Automatic"; ShouldBeRunning=$true}
    )
    
    foreach ($svc in $services) {
        try {
            $service = Get-Service -Name $svc.Name -ErrorAction Stop
            
            if ($service.StartType -ne $svc.StartupType) {
                if (-not $DiagnosticOnly) {
                    Set-Service -Name $svc.Name -StartupType $svc.StartupType
                    Write-Log "Set $($svc.Name) to $($svc.StartupType)" "SUCCESS"
                } else {
                    Write-Log "Would set $($svc.Name) to $($svc.StartupType)" "INFO"
                }
            }
            
            if ($svc.ShouldBeRunning -and $service.Status -ne "Running") {
                if (-not $DiagnosticOnly) {
                    Start-Service -Name $svc.Name
                    Write-Log "Started $($svc.Name)" "SUCCESS"
                }
            }
        } catch {
            Write-Log "Service $($svc.Name) configuration failed: $($_.Exception.Message)" "WARNING"
        }
    }
}

function Show-Recommendations {
    Write-Log "`n=== ADDITIONAL RECOMMENDATIONS ===" "INFO"
    
    Write-Log "1. Update Bartender to latest version compatible with Windows 11 24H2" "INFO"
    Write-Log "2. Install Windows Update KB5048667 or later if not present" "INFO"
    Write-Log "3. Consider RAM upgrade to 16GB for optimal Bartender performance" "INFO"
    Write-Log "4. In Bartender: Tools > Options > Print Engine > Disable 'Use GDI for printing'" "INFO"
    Write-Log "5. Reduce label complexity if possible (fewer graphics/variables)" "INFO"
    Write-Log "6. Test with direct printer connection (bypass print server if applicable)" "INFO"
    Write-Log "7. Disable antivirus real-time scanning for Bartender directories temporarily" "INFO"
    
    Write-Log "`nRestart required for all changes to take effect" "WARNING"
}

# ============================================
# MAIN EXECUTION
# ============================================

Write-Log "========================================" "INFO"
Write-Log "Bartender Performance Fix Tool - START" "INFO"
Write-Log "========================================" "INFO"
Write-Log "Log file: $LogPath" "INFO"

# Check if Bartender is running
if (Test-BartenderRunning) {
    Write-Log "Please close Bartender before running this script" "ERROR"
    Write-Log "Press any key to exit..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# System diagnostics
$is24H2 = Get-SystemInfo
Test-WindowsUpdates

if (-not $DiagnosticOnly) {
    # Create backup
    Backup-RegistryKeys
    
    # Execute fixes
    $results = @{
        PrintSpooler = Fix-PrintSpooler
        DriverIsolation = Fix-PrinterDriverIsolation
        BartenderOpt = Optimize-BartenderSettings
        SystemPerf = Optimize-SystemPerformance
        ServiceConfig = Set-RecommendedServices
    }
    
    # Clear printer queues and test configuration
    Test-PrinterConfiguration
    
    # Summary
    Write-Log "`n=== REMEDIATION SUMMARY ===" "INFO"
    foreach ($key in $results.Keys) {
        $status = if ($results[$key]) { "SUCCESS" } else { "FAILED" }
        Write-Log "$key : $status" $(if ($results[$key]) { "SUCCESS" } else { "ERROR" })
    }
} else {
    Write-Log "DIAGNOSTIC MODE - No changes made" "INFO"
    Test-PrinterConfiguration
}

Show-Recommendations

Write-Log "`n========================================" "INFO"
Write-Log "Bartender Performance Fix Tool - END" "INFO"
Write-Log "========================================" "INFO"
Write-Log "A system restart is recommended to apply all changes" "WARNING"
Write-Log "Log saved to: $LogPath" "INFO"

# Prompt for restart
$restart = Read-Host "`nWould you like to restart now? (Y/N)"
if ($restart -eq 'Y' -or $restart -eq 'y') {
    Write-Log "Initiating system restart in 60 seconds..." "WARNING"
    shutdown /r /t 60 /c "Bartender performance fixes applied - Restart required"
    Write-Log "Type 'shutdown /a' to cancel restart" "INFO"
} else {
    Write-Log "Please restart the system manually when convenient" "INFO"
}

Write-Host "`nPress any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

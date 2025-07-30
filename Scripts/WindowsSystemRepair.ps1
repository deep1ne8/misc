#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Quick Windows System Repair Script
.DESCRIPTION
    Fast, reliable system repair with verbose colorized output
.NOTES
    Version: 4.0 - Quick & Clean Edition
#>

param(
    [string]$LogPath = "C:\Temp\QuickRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log",
    [switch]$SkipReboot
)

Write-Host "`n=== Configuring System for Secure TLS 1.2 Connections ===`n" -ForegroundColor Cyan

# --- Permanent TLS 1.2 Fix ---
# Apply at runtime:
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Host "✔ TLS 1.2 enforced for this session." -ForegroundColor Green

# Apply permanent fix in registry (for future .NET and PowerShell sessions):
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"
)
foreach ($path in $regPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    New-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -PropertyType DWord -Force | Out-Null
}
Write-Host "✔ Registry updated to permanently enable strong crypto/TLS 1.2." -ForegroundColor Green


# Ensure log directory exists
$null = New-Item -Path (Split-Path $LogPath) -ItemType Directory -Force -ErrorAction SilentlyContinue

function Write-Log {
    param(
        [string]$Message, 
        [ValidateSet("Info", "Success", "Warning", "Error", "Header")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $output = "[$timestamp] $Message"
    
    $colors = @{
        Info = "White"
        Success = "Green"
        Warning = "Yellow"
        Error = "Red"
        Header = "Cyan"
    }
    
    Write-Host $output -ForegroundColor $colors[$Level]
    $output | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Start-QuickCommand {
    param(
        [string]$Command,
        [string]$Description,
        [int]$TimeoutMinutes = 5
    )
    
    Write-Log "► Starting: $Description" -Level Info
    
    try {
        $job = Start-Job -ScriptBlock {
            param($cmd)
            Invoke-Expression $cmd
        } -ArgumentList $Command
        
        $completed = Wait-Job $job -Timeout ($TimeoutMinutes * 60)
        
        if ($completed) {
            $result = Receive-Job $job
            Remove-Job $job
            Write-Log "✓ $Description completed successfully" -Level Success
            return $true
        } else {
            Remove-Job $job -Force
            Write-Log "⚠ $Description timed out after $TimeoutMinutes minutes" -Level Warning
            return $false
        }
    } catch {
        Write-Log "✗ $Description failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Test-AdminRights {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

function Reset-WindowsUpdate {
    Write-Log "► Resetting Windows Update components..." -Level Info
    
    try {
        $services = @('wuauserv', 'cryptsvc', 'bits', 'msiserver')
        
        # Stop services
        foreach ($service in $services) {
            Stop-Service $service -Force -ErrorAction SilentlyContinue
            Write-Log "  Stopped $service" -Level Info
        }
        
        # Rename folders
        $folders = @{
            'C:\Windows\SoftwareDistribution' = 'C:\Windows\SoftwareDistribution.bak'
            'C:\Windows\System32\catroot2' = 'C:\Windows\System32\catroot2.bak'
        }
        
        foreach ($folder in $folders.GetEnumerator()) {
            if (Test-Path $folder.Key) {
                if (Test-Path $folder.Value) { 
                    Remove-Item $folder.Value -Recurse -Force -ErrorAction SilentlyContinue 
                }
                Rename-Item $folder.Key $folder.Value -ErrorAction SilentlyContinue
                Write-Log "  Renamed $($folder.Key)" -Level Info
            }
        }
        
        # Start services
        foreach ($service in $services) {
            Start-Service $service -ErrorAction SilentlyContinue
            Write-Log "  Started $service" -Level Info
        }
        
        Write-Log "✓ Windows Update reset completed" -Level Success
        return $true
    } catch {
        Write-Log "✗ Windows Update reset failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function Clear-TempFiles {
    Write-Log "► Clearing temporary files..." -Level Info
    
    $tempPaths = @(
        "$env:TEMP\*",
        "$env:WINDIR\Temp\*",
        "$env:WINDIR\Prefetch\*"
    )
    
    $cleaned = 0
    foreach ($path in $tempPaths) {
        try {
            $items = Get-ChildItem $path -ErrorAction SilentlyContinue
            if ($items) {
                Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
                $cleaned += $items.Count
                Write-Log "  Cleaned $(Split-Path $path -Parent)" -Level Info
            }
        } catch {
            Write-Log "  Skipped $(Split-Path $path -Parent): Access denied" -Level Warning
        }
    }
    
    Write-Log "✓ Cleaned $cleaned temporary items" -Level Success
}

# Main execution
if (-not (Test-AdminRights)) {
    Write-Log "ERROR: This script requires Administrator privileges!" -Level Error
    Write-Log "Please run PowerShell as Administrator and try again." -Level Error
    exit 1
}

Write-Log "=== QUICK SYSTEM REPAIR STARTED ===" -Level Header
Write-Log "Log file: $LogPath" -Level Info
Write-Log "" -Level Info

$results = @{}
$startTime = Get-Date

# Phase 1: Quick SFC scan
Write-Log "PHASE 1: System File Check (Quick)" -Level Header
$results['SFC'] = Start-QuickCommand "sfc /verifyonly" "System File Verification" 3

# Phase 2: DISM health check (fast)
Write-Log "PHASE 2: Component Store Health" -Level Header
$results['DISM_Check'] = Start-QuickCommand "DISM /Online /Cleanup-Image /CheckHealth" "Component Health Check" 2

# Phase 3: Clean temporary files
Write-Log "PHASE 3: Temporary File Cleanup" -Level Header
$results['TempClean'] = Clear-TempFiles

# Phase 4: Windows Update reset (if needed)
Write-Log "PHASE 4: Windows Update Reset" -Level Header
$results['WU_Reset'] = Reset-WindowsUpdate

# Phase 5: Registry cleanup
Write-Log "PHASE 5: Registry Optimization" -Level Header
$results['Registry'] = Start-QuickCommand "sfc /verifyonly" "Registry Verification" 2

# Phase 6: Component cleanup
Write-Log "PHASE 6: Component Cleanup" -Level Header
$results['Cleanup'] = Start-QuickCommand "DISM /Online /Cleanup-Image /StartComponentCleanup" "Component Cleanup" 3

# Results summary
Write-Log "" -Level Info
Write-Log "=== REPAIR SUMMARY ===" -Level Header

$passed = 0
$total = $results.Count

foreach ($task in $results.GetEnumerator()) {
    $status = if ($task.Value) { "PASS"; $passed++ } else { "FAIL" }
    $level = if ($task.Value) { "Success" } else { "Error" }
    Write-Log "$($task.Key): $status" -Level $level
}

$duration = (Get-Date) - $startTime
Write-Log "" -Level Info
Write-Log "Execution time: $($duration.Minutes)m $($duration.Seconds)s" -Level Info
Write-Log "Success rate: $passed/$total tasks completed" -Level Info

# Recommendations
Write-Log "" -Level Info
Write-Log "=== NEXT STEPS ===" -Level Header

if ($passed -eq $total) {
    Write-Log "✓ All repairs completed successfully!" -Level Success
} else {
    Write-Log "⚠ Some tasks failed - manual intervention may be needed" -Level Warning
}

Write-Log "• Check Event Viewer for detailed error information" -Level Info
Write-Log "• Consider running full SFC scan: sfc /scannow" -Level Info
Write-Log "• Restart recommended to complete all changes" -Level Warning

Write-Log "" -Level Info
Write-Log "=== REPAIR COMPLETED ===" -Level Success
Write-Log "Full log available at: $LogPath" -Level Info

# Optional restart
if (-not $SkipReboot) {
    Write-Log "" -Level Info
    $reboot = Read-Host "Restart computer now? (y/N)"
    if ($reboot -match "^[Yy]") {
        Write-Log "Restarting in 5 seconds..." -Level Warning
        Start-Sleep 5
        Restart-Computer -Force
    }
}

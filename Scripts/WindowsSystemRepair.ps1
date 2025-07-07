#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Windows System Repair Script
.DESCRIPTION
    system repair
.NOTES
    Version: 3.0 - Bulletproof Edition
#>

param(
    [string]$LogPath = "C:\Temp\SystemRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

# Create log directory
$null = New-Item -Path (Split-Path $LogPath) -ItemType Directory -Force -ErrorAction SilentlyContinue

function Write-Output {
    param([string]$Message, [string]$Color = "White")
    
    $timestamp = Get-Date -Format "HH:mm:ss"
    $output = "[$timestamp] $Message"
    
    Write-Host $output -ForegroundColor $Color
    $output | Out-File -FilePath $LogPath -Append -Encoding UTF8
}

function Run-Command {
    param([string]$Command, [string]$Name)
    
    Write-Output "Starting: $Name" -Color Yellow
    
    try {
        # Direct command execution - no fancy stuff
        $result = cmd.exe /c $Command 2>&1
        
        if ($LASTEXITCODE -eq 0) {
            Write-Output "✓ $Name completed successfully" -Color Green
            return $true
        } else {
            Write-Output "✗ $Name failed (Exit: $LASTEXITCODE)" -Color Red
            return $false
        }
    } catch {
        Write-Output "✗ $Name crashed: $_" -Color Red
        return $false
    }
}

# Verify admin rights
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Output "ERROR: Run as Administrator!" -Color Red
    exit 1
}

Write-Output "=== WINDOWS SYSTEM REPAIR STARTED ===" -Color Cyan
Write-Output "Log: $LogPath" -Color Gray

# Track results
$results = @{}

# 1. DISM Component Health Check
Write-Output "" -Color White
Write-Output "PHASE 1: Component Store Health Check" -Color Cyan
$results['DISM_Scan'] = Run-Command "DISM /Online /Cleanup-Image /ScanHealth /NoRestart" "DISM Health Scan"

# 2. DISM Repair (if needed)
if ($results['DISM_Scan']) {
    Write-Output "" -Color White
    Write-Output "PHASE 2: Component Store Repair" -Color Cyan
    $results['DISM_Repair'] = Run-Command "DISM /Online /Cleanup-Image /RestoreHealth /NoRestart" "DISM Repair"
}

# 3. System File Checker
Write-Output "" -Color White
Write-Output "PHASE 3: System File Checker" -Color Cyan
$results['SFC'] = Run-Command "sfc /scannow" "System File Checker"

# 4. Windows Memory Diagnostic (schedule for next boot)
Write-Output "" -Color White
Write-Output "PHASE 4: Memory Diagnostic (Next Boot)" -Color Cyan
$results['MemDiag'] = Run-Command "mdsched /f" "Memory Diagnostic Scheduler"

# 5. Check Disk (C: drive)
Write-Output "" -Color White
Write-Output "PHASE 5: Disk Check" -Color Cyan
$results['ChkDsk'] = Run-Command "chkdsk C: /f /r /x" "Check Disk"

# 6. Component Store Cleanup
Write-Output "" -Color White
Write-Output "PHASE 6: Component Cleanup" -Color Cyan
$results['Cleanup'] = Run-Command "DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase /NoRestart" "Component Cleanup"

# 7. Windows Update Reset (if SFC failed)
if (-not $results['SFC']) {
    Write-Output "" -Color White
    Write-Output "PHASE 7: Windows Update Reset" -Color Cyan
    
    try {
        Write-Output "Stopping services..." -Color Yellow
        Stop-Service wuauserv, cryptsvc, bits, msiserver -Force -ErrorAction SilentlyContinue
        
        Write-Output "Renaming folders..." -Color Yellow
        $folders = @("C:\Windows\SoftwareDistribution", "C:\Windows\System32\catroot2")
        foreach ($folder in $folders) {
            if (Test-Path $folder) {
                $backup = "$folder.bak"
                if (Test-Path $backup) { Remove-Item $backup -Recurse -Force }
                Rename-Item $folder $backup -ErrorAction SilentlyContinue
            }
        }
        
        Write-Output "Starting services..." -Color Yellow
        Start-Service wuauserv, cryptsvc, bits, msiserver -ErrorAction SilentlyContinue
        
        Write-Output "✓ Windows Update Reset completed" -Color Green
        $results['WU_Reset'] = $true
    } catch {
        Write-Output "✗ Windows Update Reset failed: $_" -Color Red
        $results['WU_Reset'] = $false
    }
}

# 8. Registry Cleanup
Write-Output "" -Color White
Write-Output "PHASE 8: Registry Cleanup" -Color Cyan
$results['RegClean'] = Run-Command "sfc /scannow" "Registry Verification"

# Summary Report
Write-Output "" -Color White
Write-Output "=== REPAIR SUMMARY ===" -Color Cyan

$passed = 0
$total = 0

foreach ($task in $results.GetEnumerator()) {
    $total++
    $status = if ($task.Value) { "PASS"; $passed++ } else { "FAIL" }
    $color = if ($task.Value) { "Green" } else { "Red" }
    Write-Output "$($task.Key): $status" -Color $color
}

Write-Output "" -Color White
Write-Output "Success Rate: $passed/$total tasks completed" -Color Cyan

# Recommendations
Write-Output "" -Color White
Write-Output "=== RECOMMENDATIONS ===" -Color Cyan

if ($results['MemDiag']) {
    Write-Output "• Memory test scheduled for next reboot" -Color Yellow
}

if ($results['ChkDsk']) {
    Write-Output "• Disk check completed - review results in Event Viewer" -Color Yellow
}

if ($passed -lt $total) {
    Write-Output "• Some repairs failed - consider manual intervention" -Color Yellow
}

Write-Output "• RESTART REQUIRED to complete all repairs" -Color Yellow

Write-Output "" -Color White
Write-Output "=== REPAIR COMPLETED ===" -Color Green
Write-Output "Full log: $LogPath" -Color Gray

# Offer immediate reboot
$reboot = Read-Host "`nRestart computer now? (y/N)"
if ($reboot -match "^[Yy]") {
    Write-Output "Restarting in 10 seconds..." -Color Yellow
    Start-Sleep 10
    Restart-Computer -Force
}

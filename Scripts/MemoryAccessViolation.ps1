# Memory Access Violation Troubleshooting Script
# Run as Administrator

Write-Host "=== Memory Access Violation Troubleshooting ===" -ForegroundColor Cyan
Write-Host "Starting comprehensive system diagnostics..." -ForegroundColor Yellow

# Function to run commands with error handling
function Invoke-SafeCommand {
    param([string]$Command, [string]$Description)
    Write-Host "`n[$Description]" -ForegroundColor Green
    try {
        Invoke-Expression $Command
        Write-Host "✓ Completed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# 1. System File Checker
Invoke-SafeCommand "sfc /scannow" "Running System File Checker"

# 2. DISM Health Check and Repair
Invoke-SafeCommand "DISM /Online /Cleanup-Image /CheckHealth" "Checking Windows Image Health"
Invoke-SafeCommand "DISM /Online /Cleanup-Image /RestoreHealth" "Repairing Windows Image"

# 3. Memory Diagnostic
Write-Host "`n[Memory Diagnostic Test]" -ForegroundColor Green
Write-Host "Scheduling memory test for next reboot..."
mdsched.exe

# 4. Check for Windows Updates
Write-Host "`n[Windows Update Check]" -ForegroundColor Green
try {
    Import-Module PSWindowsUpdate -ErrorAction SilentlyContinue
    if (Get-Module PSWindowsUpdate) {
        Get-WindowsUpdate
    } else {
        Write-Host "Run Windows Update manually via Settings > Update & Security"
    }
}
catch {
    Write-Host "Check Windows Updates manually via Settings"
}

# 5. Registry Cleanup for Common Issues
Write-Host "`n[Registry Cleanup]" -ForegroundColor Green
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDLLs",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
)

foreach ($path in $regPaths) {
    if (Test-Path $path) {
        Write-Host "Checking registry path: $path" -ForegroundColor Yellow
        # Add specific registry checks here if needed
    }
}

# 6. Application-Specific Troubleshooting
Write-Host "`n[Application Troubleshooting]" -ForegroundColor Green

# Get recently crashed applications from Event Log
$crashEvents = Get-WinEvent -FilterHashtable @{LogName='Application'; ID=1000; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction SilentlyContinue

if ($crashEvents) {
    Write-Host "Recent application crashes found:" -ForegroundColor Yellow
    foreach ($event in $crashEvents) {
        $eventXML = [xml]$event.ToXml()
        $appName = $eventXML.Event.EventData.Data[0].'#text'
        $faultModule = $eventXML.Event.EventData.Data[3].'#text'
        Write-Host "  • App: $appName | Module: $faultModule | Time: $($event.TimeCreated)" -ForegroundColor White
    }
}

# 7. Hardware Diagnostic Commands
Write-Host "`n[Hardware Diagnostics]" -ForegroundColor Green

# Check disk health
Invoke-SafeCommand "chkdsk C: /f /r" "Disk Check (requires reboot confirmation)"

# Temperature and hardware info (if available)
try {
    $temp = Get-WmiObject -Class Win32_TemperatureProbe -ErrorAction SilentlyContinue
    if ($temp) {
        Write-Host "System temperature monitoring available" -ForegroundColor Green
    }
}
catch {
    Write-Host "Hardware monitoring not available via WMI" -ForegroundColor Yellow
}

# 8. Create System Restore Point
Write-Host "`n[System Restore Point]" -ForegroundColor Green
try {
    Checkpoint-Computer -Description "Before Memory Error Troubleshooting" -RestorePointType "MODIFY_SETTINGS"
    Write-Host "✓ Restore point created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Could not create restore point: $($_.Exception.Message)" -ForegroundColor Yellow
}

# 9. Clean Temporary Files
Write-Host "`n[Cleaning Temporary Files]" -ForegroundColor Green
$tempPaths = @(
    "$env:TEMP\*",
    "$env:LOCALAPPDATA\Temp\*",
    "C:\Windows\Temp\*",
    "C:\Windows\Prefetch\*"
)

foreach ($tempPath in $tempPaths) {
    try {
        Remove-Item -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "✓ Cleaned: $tempPath" -ForegroundColor Green
    }
    catch {
        Write-Host "Could not clean: $tempPath" -ForegroundColor Yellow
    }
}

# 10. Final Recommendations
Write-Host "`n=== FINAL RECOMMENDATIONS ===" -ForegroundColor Cyan
Write-Host "1. Restart your computer to complete repairs" -ForegroundColor White
Write-Host "2. Run memory diagnostic test (scheduled above)" -ForegroundColor White
Write-Host "3. If error persists, try running the problematic application:" -ForegroundColor White
Write-Host "   • As Administrator" -ForegroundColor Yellow
Write-Host "   • In Compatibility Mode" -ForegroundColor Yellow
Write-Host "   • After a clean boot" -ForegroundColor Yellow
Write-Host "4. Consider reinstalling the problematic application" -ForegroundColor White
Write-Host "5. If hardware-related, run extended memory tests" -ForegroundColor White

Write-Host "`n=== SCRIPT COMPLETED ===" -ForegroundColor Cyan
Read-Host "Press Enter to exit"

# Windows System Repair Script
$logPath = "C:\Logs\SystemRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
New-Item -ItemType Directory -Path (Split-Path $logPath) -Force | Out-Null

function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Tee-Object -FilePath $logPath -Append
}

Write-Log "Starting Windows System Repair"

# Basic DISM repairs
Write-Log "Running DISM scan..."
$dismScan = Invoke-Expression "DISM /Online /Cleanup-Image /ScanHealth" | Out-String
Write-Log $dismScan

Write-Log "Running DISM restore..."
$dismRestore = Invoke-Expression "DISM /Online /Cleanup-Image /RestoreHealth" | Out-String
Write-Log $dismRestore

# SFC scan
Write-Log "Running System File Checker..."
$sfcOutput = Invoke-Expression "sfc /scannow" | Out-String
Write-Log $sfcOutput

# Volume health check
Write-Log "Checking volume health..."
$volumeCheck = Get-Volume | Where-Object { $_.DriveLetter -eq 'C' } | Repair-Volume -OfflineScanAndFix
Write-Log "Volume health check initiated: $volumeCheck"

# Component Store Cleanup
Write-Log "Cleaning up component store..."
$cleanup = Invoke-Expression "DISM /Online /Cleanup-Image /StartComponentCleanup" | Out-String
Write-Log $cleanup

Write-Log "System repair operations completed"
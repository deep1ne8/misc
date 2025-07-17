<#
.SYNOPSIS
Automated health check and repair for Windows image issues.
.DESCRIPTION
Performs environment validation, SFC, DISM checks, and optional servicing stack verification.
Run in elevated PowerShell.
#>

Write-Host "`n=== Starting Windows Image Health Check and Repair ===`n" -ForegroundColor Cyan

# Validate SystemRoot
if ($env:SystemRoot -ne "C:\Windows") {
    Write-Warning "⚠ SystemRoot is not pointing to C:\Windows. Current value: $($env:SystemRoot)"
} else {
    Write-Host "✔ SystemRoot correctly points to C:\Windows"
}

# Validate that Windows folder exists
if (-not (Test-Path "C:\Windows\System32")) {
    Write-Error "❌ C:\Windows\System32 not found. This environment may not be valid."
    exit 1
} else {
    Write-Host "✔ Windows directory structure validated."
}

# Run SFC
Write-Host "`n--- Running SFC scan ---`n" -ForegroundColor Yellow
Start-Process -FilePath "cmd.exe" -ArgumentList "/c sfc /scannow" -Verb RunAs -Wait

# Run DISM /CheckHealth
Write-Host "`n--- Running DISM CheckHealth ---`n" -ForegroundColor Yellow
Start-Process -FilePath "dism.exe" -ArgumentList "/Online","/Cleanup-Image","/CheckHealth" -Verb RunAs -Wait

# Run DISM /ScanHealth
Write-Host "`n--- Running DISM ScanHealth ---`n" -ForegroundColor Yellow
Start-Process -FilePath "dism.exe" -ArgumentList "/Online","/Cleanup-Image","/ScanHealth" -Verb RunAs -Wait

# Run DISM /RestoreHealth
Write-Host "`n--- Running DISM RestoreHealth ---`n" -ForegroundColor Yellow
Start-Process -FilePath "dism.exe" -ArgumentList "/Online","/Cleanup-Image","/RestoreHealth" -Verb RunAs -Wait

Write-Host "`n=== Operations completed. Please review the output above. ===`n" -ForegroundColor Green

# Optional: Verify Servicing Stack Update presence
$ssu = Get-HotFix | Where-Object { $_.Description -like "*Servicing Stack*" }
if ($ssu) {
    Write-Host "`n✔ Servicing Stack Update(s) found:" -ForegroundColor Green
    $ssu | ForEach-Object { Write-Host "  $($_.HotFixID) installed on $($_.InstalledOn)" }
} else {
    Write-Warning "`n⚠ No Servicing Stack Update detected. You may need to manually install the latest SSU from Microsoft."
}

Write-Host "`n✅ Script execution finished."

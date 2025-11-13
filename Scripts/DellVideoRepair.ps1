#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Updates graphics driver using Dell Command | Update CLI
.DESCRIPTION
    Scans for current graphics driver, downloads and installs the latest version
#>

$ErrorActionPreference = 'Stop'

# Verify DCU-CLI is installed
if (-not (Get-Command dcu-cli.exe -ErrorAction SilentlyContinue)) {
    Write-Error "Dell Command | Update CLI not found. Install from: https://www.dell.com/support/dcu"
    exit 1
}

Write-Host "Checking current graphics driver..." -ForegroundColor Cyan

# Get current driver info
$currentDriver = dcu-cli /report | Select-String -Pattern "Video|Graphics" -Context 2
if ($currentDriver) {
    Write-Host "`nCurrent Graphics Driver:" -ForegroundColor Yellow
    $currentDriver | ForEach-Object { Write-Host $_.Line }
}

Write-Host "`nScanning for graphics driver updates..." -ForegroundColor Cyan

# Scan specifically for video driver updates
$scanResult = dcu-cli /scan -category=video -silent

if ($LASTEXITCODE -eq 0) {
    Write-Host "Scan completed successfully" -ForegroundColor Green
    
    # Apply updates (download and install)
    Write-Host "`nDownloading and installing latest graphics driver..." -ForegroundColor Cyan
    
    $updateResult = dcu-cli /applyUpdates -category=video -reboot=disable -silent
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nGraphics driver update completed successfully!" -ForegroundColor Green
        Write-Host "A system restart may be required for changes to take effect." -ForegroundColor Yellow
        
        # Check if reboot is needed
        $rebootRequired = dcu-cli /rebootRequired
        if ($LASTEXITCODE -eq 500) {
            Write-Host "`nREBOOT REQUIRED" -ForegroundColor Red
            $response = Read-Host "Restart now? (Y/N)"
            if ($response -eq 'Y') {
                Restart-Computer -Force
            }
        }
    } else {
        Write-Warning "Update installation encountered issues. Exit code: $LASTEXITCODE"
    }
} elseif ($LASTEXITCODE -eq 500) {
    Write-Host "No graphics driver updates available. You're up to date!" -ForegroundColor Green
} else {
    Write-Warning "Scan encountered issues. Exit code: $LASTEXITCODE"
}

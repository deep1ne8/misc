# Enable System Restore for Drive C:
Write-Host "Enabling System Restore on drive C..." -ForegroundColor Cyan
try {
    Enable-ComputerRestore -Drive C:
    Write-Host "System Restore enabled successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to enable System Restore: $($_.Exception.Message)" -ForegroundColor Red
}
Start-Sleep -Seconds 2

# Set Registry Path and Properties
Write-Host "Configuring registry settings for upgrade eligibility..." -ForegroundColor Cyan
try {
    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
    $RegPath = "HKU:\S-1-5-18\Software\Microsoft\PCHC"
    $UpgradeEligibilityValue = "1"
    New-Item $RegPath -Force | Out-Null
    New-ItemProperty -Path $RegPath -Name UpgradeEligibility -Value $UpgradeEligibilityValue -PropertyType DWORD -Force | Out-Null
    Write-Host "Registry configured successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to configure registry: $($_.Exception.Message)" -ForegroundColor Red
}
Start-Sleep -Seconds 2

# File Download and Installation
$Link = "https://go.microsoft.com/fwlink/?linkid=2171764"
$Path = "C:\Users\Public\Downloads"
$File = "windows11installationassistant.exe"

Write-Host "Downloading Windows 11 Installation Assistant..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $Link -OutFile "$Path\$File" -UseBasicParsing
    Write-Host "Download completed successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to download file: $($_.Exception.Message)" -ForegroundColor Red
    exit
}
Start-Sleep -Seconds 3

Write-Host "Starting installation process..." -ForegroundColor Cyan
try {
    Start-Process -FilePath "$Path\$File" -ArgumentList "/QuietInstall /SkipEULA /auto upgrade /NoRestartUI /copylogs $Path" -Wait
    Write-Host "Installation initiated successfully." -ForegroundColor Green
} catch {
    Write-Host "Failed to start installation: $($_.Exception.Message)" -ForegroundColor Red
}
Start-Sleep -Seconds 2

# Monitor Setup Log
Write-Host "Monitoring setup log for updates..." -ForegroundColor Cyan
try {
    Get-Content "C:\`$WINDOWS.~BT\Sources\Panther\setupact.log" -Wait -Tail 10
} catch {
    Write-Host "Failed to read setup log: $($_.Exception.Message)" -ForegroundColor Red
}

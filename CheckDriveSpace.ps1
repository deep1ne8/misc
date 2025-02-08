# Check drive space
$drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$freeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
$percentageFree = [math]::Round(($freeSpace / $drive.Size) * 100, 2)

Write-Host "`nChecking drive space..."
Write-Host "Drive: $($drive.DeviceID)"
Write-Host "Free space: $freeSpace GB"
Write-Host "Percentage free: $percentageFree%" -ForegroundColor Green

if ($percentageFree -lt 20) {
    Write-Host "Low disk space! Please free up some space." -ForegroundColor Red
    Start-Sleep -Seconds 5
}

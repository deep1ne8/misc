# DownloadFile
Write-Host "`n"
# Check if the destination path exists
Write-Host "Enter destination path: ==>  " -ForegroundColor Blue -NoNewline
$DestinationPath = Read-Host
Start-Sleep -Seconds 2
Write-Host "`n"
Write-Host "Enter URL: ==>  " -ForegroundColor Blue -NoNewline
$Url = Read-Host
Start-Sleep -Seconds 2
Write-Host "`n"
Write-Host "Download in progress..." -ForegroundColor White -BackgroundColor Green
Write-Host "Please wait..." -ForegroundColor White -BackgroundColor Green
Start-Sleep -Seconds 2
Write-Host "`n"

# Check if the destination path exists and is writable
if (!(Test-Path $DestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $DestinationPath | Out-Null
    } catch {
        Write-Host "Failed to create destination path: $_" -ForegroundColor Red
        return
    }
}

$WebClient = New-Object System.Net.WebClient

# Event handler for progress updates
$WebClient.DownloadProgressChanged += {
    param ($sender, $e)
    Write-Host "Downloading: $($e.ProgressPercentage)% completed" -ForegroundColor Yellow
}

# Event handler for download completion
$WebClient.DownloadFileCompleted += {
    Write-Host "Download completed successfully!" -ForegroundColor Green
}

# Start asynchronous download
$WebClient.DownloadFileAsync($Url, $DestinationPath)

# Keep script running until download completes
while ($WebClient.IsBusy) { Start-Sleep -Milliseconds 500 }



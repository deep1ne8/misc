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

$ProgressPreference = 'SilentlyContinue'

# Download file with progress bar
try {
    Invoke-RestMethod -Uri $Url -OutFile $DestinationPath -Method Get -Progress {$PercentComplete = $_.ProgressPercentage; Write-Host "Downloading: $PercentComplete% completed" -ForegroundColor Yellow}
    Write-Host "Download completed successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}



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

# Check for null pointer references
if ($null -eq $WebClient) {
    Write-Host "Failed to create WebClient object" -ForegroundColor Red
    return
}

# Event handler for progress updates
$WebClient.DownloadProgressChanged += {
    param ($s, $e)
    if ($null -ne $e) {
        Write-Host "Downloading: $($e.ProgressPercentage)% completed" -ForegroundColor Yellow
    } else {
        Write-Host "Error: null reference in DownloadProgressChanged event" -ForegroundColor Red
    }
}

# Event handler for download completion
$WebClient.DownloadFileCompleted += {
    param ($s, $e)
    if ($null -ne $e) {
        if ($e.Error -eq $null -and $e.Cancelled -eq $false) {
            Write-Host "Download completed successfully!" -ForegroundColor Green
        } else {
            Write-Host "Error: $($e.Error.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "Error: null reference in DownloadFileCompleted event" -ForegroundColor Red
    }
}

# Start asynchronous download
try {
    $WebClient.DownloadFileAsync($Url, $DestinationPath)
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}

# Keep script running until download completes
while ($WebClient.IsBusy) { Start-Sleep -Milliseconds 500 }



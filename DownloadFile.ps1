# Download file with BITS
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
Write-Host "Starting BITS download..." -ForegroundColor White -BackgroundColor Green
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

# Ensure running in Windows PowerShell (NOT PowerShell 7)
if ($PSVersionTable.PSEdition -eq "Core") {
    Write-Host "BITS only works in Windows PowerShell 5.1. Switching..." -ForegroundColor Yellow
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& { Start-BitsTransfer -Source '$Url' -Destination '$DestinationPath' -DisplayName '$jobName' -Asynchronous }"

}

try {
    $job = Start-BitsTransfer -Source $Url -Destination $DestinationPath -DisplayName $jobName -Asynchronous
} catch {
    Write-Host "Error starting BITS transfer: $_" -ForegroundColor Red
    return
}

# Monitor the BITS job and resume if suspended
Start-Sleep -Seconds 5  # Allow time for BITS to start processing

try {
    $job = Get-BitsTransfer | Where-Object { $_.DisplayName -eq $jobName }
    
    if ($job.JobState -eq "Suspended") {
        Write-Host "BITS job is suspended, attempting to resume..." -ForegroundColor Yellow
        Resume-BitsTransfer -BitsJob $job
    }

    # Display progress bar
    $progress = $job | Get-BitsTransfer
    Write-Progress -Activity "Downloading $Url to $DestinationPath" -Status "$($progress.BytesTransferred / 1MB) MB of $($progress.BytesTotal / 1MB) MB" -PercentComplete ($progress.BytesTransferred / $progress.BytesTotal) * 100
} catch {
    Write-Host "Failed to resume BITS transfer: $_" -ForegroundColor Red
}



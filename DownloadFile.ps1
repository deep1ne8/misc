Write-Host "File Downloader!"
Write-Host "`n"
# Check if the destination path exists
#Write-Host "Enter destination path: ==>  " -ForegroundColor Blue -NoNewline
#$DestinationPath = Read-Host
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

$BasePath = C:\temp

# Check if the destination path exists and is writable
if (!(Test-Path $BasePath)) {
    try {
        New-Item -ItemType Directory -Path $BasePath | Out-Null
    } catch {
        Write-Host "Failed to create destination path: $_" -ForegroundColor Red
        return
    }
}

# Use regex to get the file name from the URL
if ($Url -match '([^/]+)$') {
    $FileName = $matches[0]
    $FullPath = Join-Path -Path $BasePath -ChildPath $FileName
    Write-Host "Full file path: $FullPath"
}

$ProgressPreference = 'Continue'

try {
    $webResponse = Invoke-WebRequest -Uri $Url -OutFile $FullPath -Method Get -UseBasicParsing -Verbose
    $totalLength = [System.Convert]::ToInt64($webResponse.Headers["Content-Length"])
    $stream = [System.IO.File]::OpenRead($FullPath)
    $buffer = New-Object byte[] 8192
    $bytesRead = 0
    do {
        $read = $stream.Read($buffer, 0, $buffer.Length)
        $bytesRead += $read
        $percentComplete = ($bytesRead / $totalLength) * 100
        Write-Progress -PercentComplete $percentComplete -Activity "Downloading" -Status "$([math]::Round($percentComplete, 2))% completed"
    } while ($read -gt 0)
    $stream.Close()
    Write-Host "Download completed successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}



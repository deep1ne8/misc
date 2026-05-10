param (
    [string] $DestinationPath = "C:\Downloads"
)

Write-Host "Autobyte File Downloader!" -ForegroundColor Cyan
Write-Host ""

$Url = Read-Host "Enter URL:   "

if ([string]::IsNullOrWhiteSpace($Url) -or -not ($Url -match '^https?://')) {
    Write-Host "Invalid or empty URL input." -ForegroundColor Red
    return
}

$FileName = [System.IO.Path]::GetFileName($Url)
if ([string]::IsNullOrWhiteSpace($FileName)) {
    Write-Host "URL does not contain a valid file name." -ForegroundColor Red
    return
}
$FullPath = Join-Path -Path $DestinationPath -ChildPath $FileName

if (-not (Test-Path $DestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $DestinationPath -ErrorAction Stop | Out-Null
    } catch {
        Write-Host "Failed to create destination path: $_" -ForegroundColor Red
        return
    }
}

Write-Host "`n"
Write-Host "Downloading to: $FullPath" -ForegroundColor White
Write-Host "`n"

$maxRetries = 3
$retryCount = 0
$success = $false

while (-not $success -and $retryCount -lt $maxRetries) {
    try {
        Invoke-WebRequest -Uri $Url -OutFile $FullPath -Method Get -ErrorAction Stop
        Write-Host "Download completed successfully!" -ForegroundColor Green
        $success = $true
    } catch {
        $retryCount++
        Write-Host "Failed to download the file (Attempt $retryCount of $maxRetries): $_" -ForegroundColor Red
        Start-Sleep -Seconds 2
    }
}

if (-not $success) {
    Write-Host "Failed to download the file after $maxRetries attempts." -ForegroundColor Red
}


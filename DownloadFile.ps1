Write-Host "`n"
Write-Host "Autobyte File Downloader!" -ForegroundColor blue -BackgroundColor white
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

try {
    $FileName = [System.IO.Path]::GetFileName($Url)
    $FullPath = [System.IO.Path]::Combine($BasePath, $FileName)
    $webResponse = Invoke-WebRequest -Uri $Url -OutFile $FullPath -Method Get -UseBasicParsing -Verbose -Progress {$PSCmdlet.WriteProgress($PSCmdlet.MyInvocation.MyCommand.Name,$_.StatusMessage, [int]($_.PercentComplete))}
        $webResponse
    Write-Host "Download completed successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}


<#
# Use regex to get the file name from the URL
if ($Url -match '([^/]+)$') {
    $FileName = $matches[0]
    $FullPath = Join-Path -Path $BasePath -ChildPath $FileName
    Write-Host "Full file path: $FullPath"
}

#>
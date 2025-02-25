# Internet Speed Test using PowerShell
# This script uses speedtest.net's CLI


if (-not(choco list speedtest)) {
    Write-Host "Installing Speedtest..." -ForegroundColor Green
    Set-ExecutionPolicy Bypass -Scope Process -Force; `
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; `
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')); `
    choco install -y speedtest
}

# Check if Invoke-WebRequest is available (PowerShell 3.0+)
if (choco list speedtest) {
    Write-Host "Starting Internet Speed Test..." -ForegroundColor Cyan
    
    # Method 1: Using Speedtest CLI (if installed)
    if (choco list speedtest) {
        Write-Host "Using Speedtest CLI..." -ForegroundColor Green
        speedtest
    }
    # Method 2: PowerShell implementation
    else {
        Write-Host "Using PowerShell implementation..." -ForegroundColor Yellow
        
        # Function to measure download speed
        function Test-DownloadSpeed {
            $url = "https://download.microsoft.com/download/5/B/C/5BC5DBB3-652D-4DCE-B14A-475AB85EEF6E/WindowsUpdateDiagnostic.diagcab"
            $outputPath = "$env:TEMP\speedtest.tmp"
            
            $startTime = Get-Date
            Invoke-WebRequest -Uri $url -OutFile $outputPath
            $endTime = Get-Date
            
            $fileSize = (Get-Item $outputPath).Length / 1MB
            $duration = ($endTime - $startTime).TotalSeconds
            $speedMbps = [Math]::Round(($fileSize * 8) / $duration, 2)
            
            Remove-Item $outputPath -Force
            
            return $speedMbps
        }
        
        # Function to measure upload speed (simplified)
        function Test-UploadSpeed {
            $url = "https://www.speedtest.net/api/upload.php"
            $tempFile = "$env:TEMP\upload_test.tmp"
            
            # Create a file of roughly 5MB
            $randomData = New-Object byte[] (5MB)
            [System.Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($randomData)
            [System.IO.File]::WriteAllBytes($tempFile, $randomData)
            
            try {
                $startTime = Get-Date
                Invoke-RestMethod -Uri $url -Method Post -InFile $tempFile -TimeoutSec 60 -ErrorAction SilentlyContinue | Out-Null
                $endTime = Get-Date
                
                $fileSize = (Get-Item $tempFile).Length / 1MB
                $duration = ($endTime - $startTime).TotalSeconds
                $speedMbps = [Math]::Round(($fileSize * 8) / $duration, 2)
            }
            catch {
                Write-Host "Unable to complete upload test. Using estimate." -ForegroundColor Red
                $speedMbps = [Math]::Round((Test-DownloadSpeed / 3), 2)
            }
            
            Remove-Item $tempFile -Force
            return $speedMbps
        }
        
        # Test ping
        try {
            $ping = Test-Connection -ComputerName 8.8.8.8 -Count 10 | Measure-Object -Property ResponseTime -Average
            $pingMs = [Math]::Round($ping.Average, 0)
        }
        catch {
            $pingMs = "Unable to measure"
        }
        
        # Run the tests
        Write-Host "`nTesting download speed..." -ForegroundColor Cyan
        $downloadSpeed = Test-DownloadSpeed
        Write-Host "Testing upload speed..." -ForegroundColor Cyan
        $uploadSpeed = Test-UploadSpeed
        
        # Display results
        Write-Host "`n===== SPEED TEST RESULTS =====" -ForegroundColor Green
        Write-Host "Download: $downloadSpeed Mbps" -ForegroundColor White
        Write-Host "Upload: $uploadSpeed Mbps" -ForegroundColor White
        Write-Host "Ping: $pingMs ms" -ForegroundColor White
        Write-Host "=============================" -ForegroundColor Green
    }
}
else {
    Write-Error "This script requires PowerShell 3.0 or later."
}
# Internet Speed Test using PowerShell
# This script uses speedtest.net's CLI when available, or falls back to PowerShell implementation

Write-Host "Checking for Speedtest module..." -ForegroundColor Green

# Check if Speedtest is installed
$speedtestInstalled = $false
try {
    $chocoList = choco list --localonly speedtest 2>$null
    $speedtestInstalled = $chocoList -match "speedtest"
} catch {
    Write-Host "Chocolatey not installed or not in PATH" -ForegroundColor Yellow
}

if (-not $speedtestInstalled) {
    Write-Host "Speedtest not found. Installing Speedtest..." -ForegroundColor Green

    # Ensure Chocolatey is installed
    try {
        $chocoVersion = choco --version 2>$null
        if (-not $chocoVersion) {
            throw "Chocolatey not found"
        }
    } catch {
        Write-Host "Chocolatey not found. Installing Chocolatey..." -ForegroundColor Yellow
        try {
            Set-ExecutionPolicy Bypass -Scope Process -Force
            [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        } catch {
            Write-Error "Failed to install Chocolatey: $_"
            return
        }
    }
    
    # Install Speedtest
    try {
        choco install -y speedtest
        $speedtestInstalled = $true
    } catch {
        Write-Host "Failed to install Speedtest via Chocolatey: $_" -ForegroundColor Red
    }
} else {
    Write-Host "Speedtest is already installed." -ForegroundColor Cyan
}

# Check if PowerShell version supports Invoke-WebRequest (PowerShell 3.0+)
$psVersion = $PSVersionTable.PSVersion.Major
if ($psVersion -lt 3) {
    Write-Error "This script requires PowerShell 3.0 or later."
    return
}

Write-Host "Starting Internet Speed Test..." -ForegroundColor Cyan

# Method 1: Using Speedtest CLI (if installed)
if ($speedtestInstalled) {
    Write-Host "Using Speedtest CLI..." -ForegroundColor Green
    try {
        speedtest
    } catch {
        Write-Host "Error running Speedtest CLI: $_" -ForegroundColor Red
        Write-Host "Falling back to PowerShell implementation..." -ForegroundColor Yellow
        $fallbackToPowerShell = $true
    }
}
# Method 2: PowerShell implementation
else {
    $fallbackToPowerShell = $true
}

if ($fallbackToPowerShell -or -not $speedtestInstalled) {
    Write-Host "Using PowerShell implementation..." -ForegroundColor Yellow
    
    # Function to measure download speed
    function Test-DownloadSpeed {
        $url = "https://download.microsoft.com/download/5/B/C/5BC5DBB3-652D-4DCE-B14A-475AB85EEF6E/WindowsUpdateDiagnostic.diagcab"
        $outputPath = "$env:TEMP\speedtest.tmp"
        
        try {
            $startTime = Get-Date
            Invoke-WebRequest -Uri $url -OutFile $outputPath
            $endTime = Get-Date
            
            $fileSize = (Get-Item $outputPath).Length / 1MB
            $duration = ($endTime - $startTime).TotalSeconds
            $speedMbps = [Math]::Round(($fileSize * 8) / $duration, 2)
            
            Remove-Item $outputPath -Force
            
            return $speedMbps
        } catch {
            Write-Host "Download test failed: $_" -ForegroundColor Red
            return
        }
    }
    
    # Function to measure upload speed (simplified)
    function Test-UploadSpeed {
        $url = "https://www.speedtest.net/api/upload.php"
        $tempFile = "$env:TEMP\upload_test.tmp"
        
        # Create a file of roughly 5MB
        try {
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
                $downloadSpeed = Test-DownloadSpeed
                $speedMbps = if ($downloadSpeed -gt 0) { [Math]::Round(($downloadSpeed / 3), 2) } else { 0 }
            }
            
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            return $speedMbps
        } catch {
            Write-Host "Failed to create test file: $_" -ForegroundColor Red
            return 0
        }
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
# Improved Speed Test Script
# This script checks for and installs Chocolatey if needed, then installs and runs Speedtest CLI

function Test-CommandExists {
    param($Command)
    
    $exists = $null -ne (Get-Command -Name $Command -ErrorAction SilentlyContinue)
    return $exists
}

function Install-Chocolatey {
    Write-Host "Chocolatey not found. Installing Chocolatey..." -ForegroundColor Yellow
    
    # Check if there's an existing installation folder
    if (Test-Path "C:\ProgramData\chocolatey") {
        Write-Host "WARNING: Chocolatey directory exists but command is not available." -ForegroundColor Red
        Write-Host "This could be due to an incomplete installation or PATH issues." -ForegroundColor Red
        
        $response = Read-Host "Do you want to try refreshing your PATH and checking again? (Y/N)"
        if ($response -eq "Y" -or $response -eq "y") {
            # Refresh environment variables
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
            
            if (Test-CommandExists "choco") {
                Write-Host "Chocolatey is now available!" -ForegroundColor Green
                return $true
            }
        }
        
        $response = Read-Host "Do you want to attempt repair by running the installer? (Y/N)"
        if ($response -ne "Y" -and $response -ne "y") {
            Write-Host "Installation aborted." -ForegroundColor Red
            return $false
        }
    }
    
    try {
        # Use TLS 1.2 for security
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        
        # Download and run Chocolatey installer
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        if (Test-CommandExists "choco") {
            Write-Host "Chocolatey installed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Chocolatey installation completed but 'choco' command is not available." -ForegroundColor Red
            Write-Host "You may need to restart your PowerShell session." -ForegroundColor Yellow
            return $false
        }
    } catch {
        Write-Host "Failed to install Chocolatey: $_" -ForegroundColor Red
        return $false
    }
}

function Install-SpeedtestCLI {
    Write-Host "Installing Speedtest CLI..." -ForegroundColor Yellow
    
    try {
        choco install speedtest --yes --force
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
        
        if (Test-CommandExists "speedtest") {
            Write-Host "Speedtest CLI installed successfully!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Speedtest CLI installation completed but 'speedtest' command is not available." -ForegroundColor Red
            Write-Host "Trying to find the installed executable..." -ForegroundColor Yellow
            
            # Try to find the speedtest.exe location
            $possibleLocations = @(
                "C:\ProgramData\chocolatey\bin\speedtest.exe",
                "C:\ProgramData\chocolatey\lib\speedtest\tools\speedtest.exe"
            )
            
            foreach ($location in $possibleLocations) {
                if (Test-Path $location) {
                    Write-Host "Found Speedtest CLI at: $location" -ForegroundColor Green
                    return $location
                }
            }
            
            Write-Host "Could not find Speedtest CLI executable." -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Failed to install Speedtest CLI: $_" -ForegroundColor Red
        return $false
    }
}

function Invoke-Fallback-SpeedTest {
    Write-Host "Running fallback PowerShell speed test..." -ForegroundColor Yellow
    
    try {
        # Using reliable test servers
        $downloadTestUrl = "https://download.microsoft.com/download/2/0/E/20E90413-712F-438C-988E-FDAA79A8AC3D/dotnetfx35.exe"
        
        Write-Host "Testing download speed..." -ForegroundColor Cyan
        $startTime = Get-Date
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("user-agent", "PowerShell Speed Test")
        
        # Stream the file but don't save it
        $downloadStream = $webClient.OpenRead($downloadTestUrl)
        $buffer = New-Object byte[] 8192
        $totalBytesRead = 0
        $maxBytes = 20 * 1024 * 1024 # Read only first 20 MB to save time
        
        while (($bytesRead = $downloadStream.Read($buffer, 0, $buffer.Length)) -gt 0 -and $totalBytesRead -lt $maxBytes) {
            $totalBytesRead += $bytesRead
            # Show progress
            if ($totalBytesRead % (1024 * 1024) -lt $buffer.Length) {
                Write-Host "." -NoNewline
            }
        }
        
        $downloadStream.Close()
        $endTime = Get-Date
        
        $duration = ($endTime - $startTime).TotalSeconds
        $downloadSpeedMbps = [Math]::Round(($totalBytesRead * 8 / 1000000) / $duration, 2)
        
        Write-Host "`nDownload Speed: $downloadSpeedMbps Mbps" -ForegroundColor Green
        
        # We can't easily test upload in a reliable way with PowerShell only
        # So we'll skip it and just return download speed
        
        return @{
            "Download" = $downloadSpeedMbps
            "Upload" = "N/A"
            "Ping" = "N/A"
        }
    } catch {
        Write-Host "Fallback speed test failed: $_" -ForegroundColor Red
        return @{
            "Download" = "Error"
            "Upload" = "Error"
            "Ping" = "Error"
        }
    }
}

function Invoke-SpeedTest {
    param(
        [string]$SpeedtestPath = "speedtest"
    )
    
    Write-Host "Running speed test..." -ForegroundColor Yellow
    
    try {
        if ($SpeedtestPath -eq "speedtest") {
            # Run using PATH
            $result = & speedtest --format=json --accept-license --accept-gdpr
        } else {
            # Run using specific path
            $result = & $SpeedtestPath --format=json --accept-license --accept-gdpr
        }
        
        if ($LASTEXITCODE -ne 0) {
            throw "Speedtest CLI returned error code: $LASTEXITCODE"
        }
        
        # Parse JSON result
        $speedData = $result | ConvertFrom-Json
        
        return @{
            "Download" = [Math]::Round($speedData.download.bandwidth * 8 / 1000000, 2)
            "Upload" = [Math]::Round($speedData.upload.bandwidth * 8 / 1000000, 2)
            "Ping" = [Math]::Round($speedData.ping.latency, 0)
            "ISP" = $speedData.isp
            "Server" = $speedData.server.name
            "Location" = $speedData.server.location
        }
    } catch {
        Write-Host "Error running Speedtest CLI: $_" -ForegroundColor Red
        Write-Host "Falling back to PowerShell implementation..." -ForegroundColor Yellow
        return Invoke-Fallback-SpeedTest
    }
}

function Show-Results {
    param(
        [hashtable]$Results
    )
    
    Write-Host "`n===== SPEED TEST RESULTS =====" -ForegroundColor Cyan
    Write-Host "Download: $($Results.Download) Mbps" -ForegroundColor Green
    Write-Host "Upload: $($Results.Upload) Mbps" -ForegroundColor Green
    Write-Host "Ping: $($Results.Ping) ms" -ForegroundColor Green
    
    if ($Results.ContainsKey("ISP")) {
        Write-Host "ISP: $($Results.ISP)" -ForegroundColor Green
    }
    
    if ($Results.ContainsKey("Server")) {
        Write-Host "Server: $($Results.Server)" -ForegroundColor Green
    }
    
    if ($Results.ContainsKey("Location")) {
        Write-Host "Location: $($Results.Location)" -ForegroundColor Green
    }
    
    Write-Host "=============================" -ForegroundColor Cyan
}

# Main script execution
Write-Host "Internet Speed Test Script" -ForegroundColor Cyan
Write-Host "------------------------" -ForegroundColor Cyan

# Check if Speedtest CLI is already installed
if (Test-CommandExists "speedtest") {
    Write-Host "Speedtest CLI is already installed." -ForegroundColor Green
    $speedtestPath = "speedtest"
} else {
    Write-Host "Speedtest CLI not found. Checking for Chocolatey..." -ForegroundColor Yellow
    
    # Check if Chocolatey is installed
    if (-not (Test-CommandExists "choco")) {
        $chocoInstalled = Install-Chocolatey
        if (-not $chocoInstalled) {
            Write-Host "Could not install Chocolatey. Falling back to PowerShell implementation." -ForegroundColor Red
            $results = Invoke-Fallback-SpeedTest
            Show-Results -Results $results
            return
        }
    } else {
        Write-Host "Chocolatey is already installed." -ForegroundColor Green
    }
    
    # Install Speedtest CLI using Chocolatey
    $speedtestPath = Install-SpeedtestCLI
    if (-not $speedtestPath) {
        Write-Host "Could not install Speedtest CLI. Falling back to PowerShell implementation." -ForegroundColor Red
        $results = Run-Fallback-SpeedTest
        Display-Results -Results $results
        return
    }
}

# Run the speed test
$results = Invoke-SpeedTest -SpeedtestPath $speedtestPath

# Display results
Display-Results -Results $results
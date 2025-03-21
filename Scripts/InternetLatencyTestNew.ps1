$targets = @("8.8.8.8", "1.1.1.1", "www.google.com", "www.microsoft.com")
$pingCount = 10
$results = @()
$wifiSignal = $null

# Function to calculate jitter
function Get-Jitter {
    param([int[]]$latencies)
    if ($latencies.Count -lt 2) { return 0 }
    $differences = @()
    for ($i = 1; $i -lt $latencies.Count; $i++) {
        $differences += [math]::Abs($latencies[$i] - $latencies[$i - 1])
    }
    return [math]::Round(($differences | Measure-Object -Average).Average, 2)
}

# Function to perform latency test
function Test-Latency {
    param([string]$target)
    $pings = Test-Connection -ComputerName $target -Count $pingCount -ErrorAction SilentlyContinue
    $latencies = @()
    
    if ($pings) {
        foreach ($ping in $pings) {
            $latencies += $ping.ResponseTime
        }
    }

    # Calculate statistics
    $min = ($latencies | Measure-Object -Minimum).Minimum
    $max = ($latencies | Measure-Object -Maximum).Maximum
    $avg = ($latencies | Measure-Object -Average).Average
    $loss = 100 - (($latencies.Count / $pingCount) * 100)
    $jitter = Get-Jitter -latencies $latencies

    return [PSCustomObject]@{
        Target = $target
        Avg = [math]::Round($avg, 2)
        Min = $min
        Max = $max
        Loss = "$loss%"
        Jitter = "$jitter ms"
    }
}

# Function to check Wi-Fi signal strength
function Get-WiFi-Signal {
    $wifiInfo = netsh wlan show interfaces | Select-String "Signal"
    if ($wifiInfo) {
        $signalStrength = $wifiInfo -replace ".*Signal\s*:\s*", "" -replace "%", ""
        $rssi = [math]::Round(($signalStrength / 2) - 100, 2) # Convert % to RSSI (dBm)
        return $rssi
    }
    return $null
}

# Check if on Wi-Fi and get signal strength
$wifiSignal = Get-WiFi-Signal

# Run tests for each target
foreach ($target in $targets) {
    $results += Test-Latency -target $target
}

# Display results
$results | Format-Table -AutoSize -Property Target, Avg, Min, Max, Loss, Jitter

# Wi-Fi Signal Analysis
if ($null -ne $wifiSignal) {
    Write-Host "`n--- Wi-Fi Signal Analysis ---" -ForegroundColor Cyan
    Write-Host "Signal Strength: $wifiSignal dBm" -ForegroundColor Yellow
    if ($wifiSignal -gt -50) {
        Write-Host "✅ Excellent Signal Strength (Strong Connection)" -ForegroundColor Green
    } elseif ($wifiSignal -gt -60) {
        Write-Host "✅ Good Signal Strength (Stable Connection)" -ForegroundColor Green
    } elseif ($wifiSignal -gt -70) {
        Write-Host "⚠️ Moderate Signal Strength (May cause occasional lag)" -ForegroundColor Yellow
    } else {
        Write-Host "🚨 Weak Signal Strength! (High chance of packet loss and lag)" -ForegroundColor Red
        Write-Host "   🔎 Possible Causes: Distance from router, interference from walls/devices."
        Write-Host "   🛠 Potential Solutions: Move closer to router, switch to 5GHz band, reduce interference, or upgrade router."
    }
}

# Network Analysis with Causes & Solutions
Write-Host "`n--- Network Analysis ---" -ForegroundColor Cyan
foreach ($result in $results) {
    Write-Host "`nAnalyzing: $($result.Target)" -ForegroundColor Green

    # Packet Loss Analysis
    if ($result.Loss -match "\d+" -and [int]$result.Loss -gt 2) {
        Write-Host "⚠️ Packet Loss Detected: $($result.Loss)" -ForegroundColor Yellow
        Write-Host "   🔎 Possible Causes: ISP issues, router congestion, weak WiFi signal, firewall interference."
        Write-Host "   🛠 Potential Solutions: Restart router, switch to a wired connection, check ISP status, reduce network load."
    }

    # High Latency Spikes
    if ($result.Max -gt 150) {
        Write-Host "⚠️ High Latency Spikes: Max = $($result.Max) ms" -ForegroundColor Red
        Write-Host "   🔎 Possible Causes: Network congestion, ISP throttling, poor routing, VPN interference."
        Write-Host "   🛠 Potential Solutions: Restart modem, change DNS servers, disable VPN, check for background downloads."
    }

    # High Jitter Detection
    if ($result.Jitter -match "\d+" -and [int]$result.Jitter -gt 10) {
        Write-Host "⚠️ High Jitter: $($result.Jitter)" -ForegroundColor Yellow
        Write-Host "   🔎 Possible Causes: Wireless interference, overloaded router, ISP instability."
        Write-Host "   🛠 Potential Solutions: Use wired connection, upgrade router, reduce connected devices."
    }

    # Stable Connection Analysis
    if ($result.Loss -eq "0%" -and $result.Max -le 150 -and [int]$result.Jitter -le 10) {
        Write-Host "✅ Connection is stable for $($result.Target)" -ForegroundColor Green
    }
}

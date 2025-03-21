$targets = @("8.8.8.8", "1.1.1.1", "www.google.com", "www.microsoft.com")
$pingCount = 10
$results = @()
$wifiSignal = $null

# Function to calculate jitter manually (PS 5.1 does not support .Average on arrays)
function Get-Jitter {
    param([int[]]$latencies)
    if ($latencies.Count -lt 2) { return 0 }

    $differences = @()
    for ($i = 1; $i -lt $latencies.Count; $i++) {
        $differences += [math]::Abs($latencies[$i] - $latencies[$i - 1])
    }

    return [math]::Round(($differences | Measure-Object -Average).Average, 2)
}

# Function to perform latency test with live progress output
function Test-Latency {
    param([string]$target)
    
    Write-Host "`nTesting latency to: $target..." -ForegroundColor Cyan
    $latencies = @()
    $lostPackets = 0

    for ($i = 1; $i -le $pingCount; $i++) {
        $ping = Test-Connection -ComputerName $target -Count 1 -ErrorAction SilentlyContinue

        if ($ping) {
            $latency = $ping.ResponseTime
            $latencies += $latency
            Write-Host "[Ping $i] ✅ $latency ms" -ForegroundColor Green -NoNewline
        } else {
            $lostPackets++
            Write-Host "[Ping $i] ❌ Request timed out" -ForegroundColor Red -NoNewline
        }

        Write-Host " | Progress: [$i/$pingCount]"
        Start-Sleep -Milliseconds 500  # Small delay for readability
    }

    if ($latencies.Count -gt 0) {
        $min = ($latencies | Measure-Object -Minimum).Minimum
        $max = ($latencies | Measure-Object -Maximum).Maximum
        $avg = ($latencies | Measure-Object -Average).Average
        $loss = [math]::Round(($lostPackets / $pingCount) * 100, 2)
        $jitter = Get-Jitter -latencies $latencies
    } else {
        $min = "N/A"
        $max = "N/A"
        $avg = "N/A"
        $loss = "100%"
        $jitter = "N/A"
    }

    return [PSCustomObject]@{
        Target = $target
        Avg = "$avg ms"
        Min = "$min ms"
        Max = "$max ms"
        Loss = "$loss%"
        Jitter = "$jitter ms"
    }
}

# Function to check Wi-Fi signal strength (PS 5.1 Compatible)
function Get-WiFi-Signal {
    try {
        $wifiInfo = netsh wlan show interfaces | Select-String "Signal"
        if ($wifiInfo) {
            if ($wifiInfo -match "Signal\s*:\s*(\d+)") {
                $signalStrength = $matches[1] -as [int]
                if ($signalStrength -ne $null) {
                    $rssi = [math]::Round(($signalStrength / 2) - 100, 2) # Convert % to RSSI (dBm)
                    return $rssi
                }
            }
        }
        return $null
    } catch {
        Write-Host "Error retrieving Wi-Fi signal strength: $_" -ForegroundColor Red
        return $null
    }
}

# Check if on Wi-Fi and get signal strength
$wifiSignal = Get-WiFi-Signal

# Run tests for each target
foreach ($target in $targets) {
    $results += Test-Latency -target $target
}

# Display results
Write-Host "`n--- Final Latency Results ---" -ForegroundColor Cyan
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

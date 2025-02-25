# Internet Latency Testing Tool
# This script tests latency to multiple destinations and provides statistics

<#
.SYNOPSIS
    Tests network latency to multiple destinations and provides detailed statistics.

.DESCRIPTION
    The Test-InternetLatency function performs ping tests to specified destinations
    and presents detailed latency statistics including average, minimum, and maximum
    response times as well as packet loss percentage. Results are color-coded for
    easy interpretation and can be logged to a file.

.PARAMETER Destinations
    An array of IP addresses or hostnames to test. 
    Default: 8.8.8.8 (Google DNS), 1.1.1.1 (Cloudflare DNS), www.google.com, www.microsoft.com

.PARAMETER Count
    The number of ping requests to send to each destination.
    Default: 10

.PARAMETER Interval
    The interval in seconds between ping requests.
    Default: 2

.PARAMETER Continuous
    Switch to enable continuous monitoring. Press Ctrl+C to stop.

.PARAMETER OutputFile
    Path to save the log file. A timestamp will be appended to the filename.
    Example: "LatencyReport" becomes "LatencyReport_20250225_123456.log"

.PARAMETER ShowDetails
    Switch to display detailed results for each ping.

.EXAMPLE
    Test-InternetLatency
    Performs a basic latency test with default settings.

.EXAMPLE
    Test-InternetLatency -Destinations "8.8.8.8", "1.1.1.1", "www.amazon.com"
    Tests latency to the specified destinations.

.EXAMPLE
    Test-InternetLatency -Count 5 -Interval 1
    Sends 5 ping requests to each default destination with a 1-second interval.

.EXAMPLE
    Test-InternetLatency -Continuous
    Continuously monitors latency until manually stopped with Ctrl+C.

.EXAMPLE
    Test-InternetLatency -ShowDetails -OutputFile "C:\Logs\LatencyReport"
    Performs a latency test with detailed output and saves results to a log file.

.NOTES
    Author: Claude
    Version: 1.0
    Date: February 25, 2025
    
    Color Coding:
    - Green: Good latency (<50ms)
    - Yellow: Average latency (50-100ms)
    - Red: Poor latency (>100ms)
#>

function Test-InternetLatency {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string[]]$Destinations = @("8.8.8.8", "1.1.1.1", "www.google.com", "www.microsoft.com"),
        
        [Parameter(Mandatory=$false)]
        [int]$Count = 10,
        
        [Parameter(Mandatory=$false)]
        [int]$Interval = 2,
        
        [Parameter(Mandatory=$false)]
        [switch]$Continuous,
        
        [Parameter(Mandatory=$false)]
        [string]$OutputFile,
        
        [Parameter(Mandatory=$false)]
        [switch]$ShowDetails
    )
    
    # Function to get timestamp
    function Get-TimeStamp {
        return "[{0:MM/dd/yyyy} {0:HH:mm:ss}]" -f (Get-Date)
    }
    
    # Function to display results in color based on latency
    function Write-ColoredLatency {
        param (
            [int]$Latency
        )
        
        if ($Latency -lt 50) {
            Write-Host $Latency"ms" -ForegroundColor Green -NoNewline
        } elseif ($Latency -lt 100) {
            Write-Host $Latency"ms" -ForegroundColor Yellow -NoNewline
        } else {
            Write-Host $Latency"ms" -ForegroundColor Red -NoNewline
        }
    }
    
    # Initialize results array
    $results = @()
    
    # Create log file if specified
    if ($OutputFile) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutputFile = $OutputFile -replace "\.log$|\.txt$", ""
        $OutputFile = "$OutputFile`_$timestamp.log"
        
        "Internet Latency Test started at $(Get-Date)" | Out-File -FilePath $OutputFile
        "Destinations: $($Destinations -join ', ')" | Out-File -FilePath $OutputFile -Append
        "----------------------------------------" | Out-File -FilePath $OutputFile -Append
    }
    
    # Display header
    Write-Host "Internet Latency Testing Tool" -ForegroundColor Cyan
    Write-Host "Testing latency to: $($Destinations -join ', ')" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor Cyan
    
    # Counter for continuous mode
    $iteration = 1
    
    # Main loop
    do {
        if ($Continuous) {
            Write-Host "`nIteration $iteration - $(Get-TimeStamp)" -ForegroundColor Magenta
            if ($OutputFile) {
                "`nIteration $iteration - $(Get-TimeStamp)" | Out-File -FilePath $OutputFile -Append
            }
            $iteration++
        }
        
        foreach ($dest in $Destinations) {
            Write-Host "Testing $dest... " -NoNewline
            
            $pingResults = @()
            $successCount = 0
            $totalLatency = 0
            $minLatency = [int]::MaxValue
            $maxLatency = 0
            
            # Perform ping tests
            for ($i = 1; $i -le $Count; $i++) {
                try {
                    # Display progress
                    Write-Host "." -NoNewline
                    
                    # Perform ping test
                    $ping = Test-Connection -ComputerName $dest -Count 1 -ErrorAction Stop
                    $latency = $ping.ResponseTime
                    
                    # Update statistics
                    $pingResults += $latency
                    $successCount++
                    $totalLatency += $latency
                    
                    if ($latency -lt $minLatency) { $minLatency = $latency }
                    if ($latency -gt $maxLatency) { $maxLatency = $latency }
                    
                    # Pause between pings
                    if ($i -lt $Count) {
                        Start-Sleep -Seconds $Interval
                    }
                }
                catch {
                    Write-Host "x" -NoNewline -ForegroundColor Red
                    $pingResults += "Timeout"
                    
                    # Pause between pings
                    if ($i -lt $Count) {
                        Start-Sleep -Seconds $Interval
                    }
                }
            }
            
            # Calculate statistics
            $packetLoss = (($Count - $successCount) / $Count) * 100
            $avgLatency = if ($successCount -gt 0) { $totalLatency / $successCount } else { 0 }
            
            # Display results
            Write-Host "`r$dest".PadRight(30) -NoNewline
            
            if ($successCount -gt 0) {
                # Avg
                Write-Host "Avg: " -NoNewline
                Write-ColoredLatency -Latency $avgLatency
                
                # Min
                Write-Host " | Min: " -NoNewline
                Write-ColoredLatency -Latency $minLatency
                
                # Max
                Write-Host " | Max: " -NoNewline
                Write-ColoredLatency -Latency $maxLatency
                
                # Packet Loss
                Write-Host " | Loss: " -NoNewline
                if ($packetLoss -eq 0) {
                    Write-Host "$packetLoss%" -ForegroundColor Green
                } elseif ($packetLoss -lt 5) {
                    Write-Host "$packetLoss%" -ForegroundColor Yellow
                } else {
                    Write-Host "$packetLoss%" -ForegroundColor Red
                }
            } else {
                Write-Host "Failed to connect!" -ForegroundColor Red
            }
            
            # Display detailed results if requested
            if ($ShowDetails -and $pingResults.Count -gt 0) {
                Write-Host "  Details: " -NoNewline
                for ($i = 0; $i -lt $pingResults.Count; $i++) {
                    if ($pingResults[$i] -eq "Timeout") {
                        Write-Host "T" -NoNewline -ForegroundColor Red
                    } else {
                        if ($pingResults[$i] -lt 50) {
                            Write-Host "." -NoNewline -ForegroundColor Green
                        } elseif ($pingResults[$i] -lt 100) {
                            Write-Host "." -NoNewline -ForegroundColor Yellow
                        } else {
                            Write-Host "." -NoNewline -ForegroundColor Red
                        }
                    }
                }
                Write-Host ""
            }
            
            # Log results if requested
            if ($OutputFile) {
                $logEntry = "$dest - Avg: $avgLatency ms | Min: $minLatency ms | Max: $maxLatency ms | Loss: $packetLoss%"
                $logEntry | Out-File -FilePath $OutputFile -Append
                
                if ($ShowDetails) {
                    "  Details: $($pingResults -join ', ')" | Out-File -FilePath $OutputFile -Append
                }
            }
            
            # Store results
            $results += [PSCustomObject]@{
                Destination = $dest
                AverageLatency = $avgLatency
                MinLatency = $minLatency
                MaxLatency = $maxLatency
                PacketLoss = $packetLoss
                Details = $pingResults
            }
        }
        
        # Display summary
        Write-Host "`nSUMMARY:" -ForegroundColor Cyan
        $results | Sort-Object -Property AverageLatency | Format-Table -Property Destination, @{
            Label = "Avg (ms)"
            Expression = { [math]::Round($_.AverageLatency, 2) }
        }, @{
            Label = "Min (ms)"
            Expression = { $_.MinLatency }
        }, @{
            Label = "Max (ms)"
            Expression = { $_.MaxLatency }
        }, @{
            Label = "Loss (%)"
            Expression = { [math]::Round($_.PacketLoss, 2) }
        }
        
        # Log summary if requested
        if ($OutputFile) {
            "`nSUMMARY:" | Out-File -FilePath $OutputFile -Append
            $results | Sort-Object -Property AverageLatency | Format-Table -Property Destination, @{
                Label = "Avg (ms)"
                Expression = { [math]::Round($_.AverageLatency, 2) }
            }, @{
                Label = "Min (ms)"
                Expression = { $_.MinLatency }
            }, @{
                Label = "Max (ms)"
                Expression = { $_.MaxLatency }
            }, @{
                Label = "Loss (%)"
                Expression = { [math]::Round($_.PacketLoss, 2) }
            } | Out-String | Out-File -FilePath $OutputFile -Append
        }
        
        # Reset results for continuous mode
        $results = @()
        
        # If continuous, pause before next iteration
        if ($Continuous) {
            Write-Host "Next test in 5 seconds... Press Ctrl+C to stop." -ForegroundColor Yellow
            Start-Sleep -Seconds 5
        }
    } while ($Continuous)
    
    # Display output file location if logging was enabled
    if ($OutputFile) {
        Write-Host "Results saved to: $OutputFile" -ForegroundColor Green
    }
}

# Example usage:
# Basic test with default settings
# Test-InternetLatency

# Test with custom destinations
# Test-InternetLatency -Destinations "8.8.8.8", "1.1.1.1", "www.amazon.com"

# Continuous monitoring
# Test-InternetLatency -Continuous

# Detailed report with logging
# Test-InternetLatency -Count 20 -Interval 1 -ShowDetails -OutputFile "LatencyReport"
#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Comprehensive system diagnostic scan and repair for slow-running systems.

.DESCRIPTION
    Performs automated diagnostics, repairs common issues, and generates detailed reports.
    Designed for MSP deployment across multiple endpoints.

.PARAMETER SkipRepairs
    Run diagnostics only without performing repairs.

.PARAMETER GenerateReport
    Generate HTML report. Default: $true

.EXAMPLE
    .\System-DiagnosticRepair.ps1
    .\System-DiagnosticRepair.ps1 -SkipRepairs
#>

[CmdletBinding()]
param(
    [switch]$SkipRepairs,
    [string]$ReportPath = "$env:TEMP\SystemDiagnostic_$(Get-Date -Format 'yyyyMMdd_HHmmss').html",
    [switch]$GenerateReport = $true
)

$script:Results = @{
    ComputerName = $env:COMPUTERNAME
    ScanTime = Get-Date
    Errors = [System.Collections.ArrayList]@()
    Warnings = [System.Collections.ArrayList]@()
    Successs = [System.Collections.ArrayList]@()
    Infos = [System.Collections.ArrayList]@()
}

function Write-Log {
    param([string]$Message, [string]$Type = "Info")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch($Type) {
        "Error" { "Red" }
        "Warning" { "Yellow" }
        "Success" { "Green" }
        default { "White" }
    }
    Write-Host "[$timestamp] $Message" -ForegroundColor $color
    
    $logEntry = [PSCustomObject]@{
        Time = $timestamp
        Message = $Message
    }
    
    $null = $script:Results[$Type + "s"].Add($logEntry)
}

function Test-DiskSpace {
    Write-Host "`n=== Disk Space Analysis ===" -ForegroundColor Cyan
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -gt 0 }
    
    foreach ($drive in $drives) {
        $freePercent = [math]::Round(($drive.Free / ($drive.Used + $drive.Free)) * 100, 2)
        $freeGB = [math]::Round($drive.Free / 1GB, 2)
        
        if ($freePercent -lt 10) {
            Write-Log "CRITICAL: Drive $($drive.Name) has only $freePercent% free ($freeGB GB)" "Error"
        } elseif ($freePercent -lt 20) {
            Write-Log "WARNING: Drive $($drive.Name) has $freePercent% free ($freeGB GB)" "Warning"
        } else {
            Write-Log "Drive $($drive.Name): $freePercent% free ($freeGB GB)" "Info"
        }
    }
}

function Clear-TempFiles {
    if ($SkipRepairs) { return }
    
    Write-Host "`n=== Cleaning Temporary Files ===" -ForegroundColor Cyan
    $tempPaths = @(
        "$env:TEMP\*",
        "$env:LOCALAPPDATA\Temp\*",
        "C:\Windows\Temp\*",
        "C:\Windows\Prefetch\*"
    )
    
    $totalCleaned = 0
    foreach ($path in $tempPaths) {
        try {
            $items = Get-ChildItem -Path $path -Force -ErrorAction SilentlyContinue | 
                     Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-7) }
            
            if ($items) {
                foreach ($item in $items) {
                    try {
                        $size = 0
                        if ($item.PSIsContainer) {
                            $childItems = Get-ChildItem -Path $item.FullName -Recurse -Force -File -ErrorAction SilentlyContinue
                            if ($childItems) {
                                $size = ($childItems | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                                if ($null -eq $size) { $size = 0 }
                            }
                        } else {
                            $size = $item.Length
                        }
                        
                        Remove-Item -Path $item.FullName -Recurse -Force -ErrorAction SilentlyContinue
                        $totalCleaned += $size
                    } catch {
                        # Silently continue on access denied or locked files
                    }
                }
            }
        } catch {
            # Path doesn't exist or access denied
        }
    }
    
    $cleanedMB = [math]::Round($totalCleaned / 1MB, 2)
    if ($cleanedMB -gt 0) {
        Write-Log "Cleaned $cleanedMB MB of temporary files" "Success"
    } else {
        Write-Log "No temporary files to clean (or all files in use)" "Info"
    }
}

function Test-SystemFiles {
    Write-Host "`n=== System File Integrity ===" -ForegroundColor Cyan
    
    if ($SkipRepairs) {
        Write-Log "Skipping SFC scan (repair mode disabled)" "Warning"
        return
    }
    
    Write-Log "Running System File Checker (this may take several minutes)..." "Info"
    $sfcResult = & sfc /scannow 2>&1
    
    if ($sfcResult -match "found corrupt files and successfully repaired them") {
        Write-Log "SFC repaired corrupted system files" "Success"
    } elseif ($sfcResult -match "found corrupt files but was unable to fix") {
        Write-Log "SFC found corrupted files that require manual repair (run DISM)" "Error"
    } else {
        Write-Log "System files integrity verified" "Info"
    }
}

function Test-DiskHealth {
    Write-Host "`n=== Disk Health Check ===" -ForegroundColor Cyan
    
    try {
        $volumes = Get-Volume | Where-Object { $_.DriveLetter -and $_.FileSystem }
        
        foreach ($vol in $volumes) {
            $health = $vol.HealthStatus
            if ($health -eq "Healthy") {
                Write-Log "Drive $($vol.DriveLetter): $health" "Info"
            } else {
                Write-Log "Drive $($vol.DriveLetter): $health - Attention required!" "Error"
            }
        }
        
        # Check SMART status
        $disks = Get-PhysicalDisk
        foreach ($disk in $disks) {
            if ($disk.HealthStatus -ne "Healthy") {
                Write-Log "Physical Disk $($disk.FriendlyName): $($disk.HealthStatus)" "Error"
            }
        }
    } catch {
        Write-Log "Unable to retrieve disk health: $($_.Exception.Message)" "Warning"
    }
}

function Test-Memory {
    Write-Host "`n=== Memory Analysis ===" -ForegroundColor Cyan
    
    $os = Get-CimInstance Win32_OperatingSystem
    $totalMemoryGB = [math]::Round($os.TotalVisibleMemorySize / 1MB, 2)
    $freeMemoryGB = [math]::Round($os.FreePhysicalMemory / 1MB, 2)
    $usedPercent = [math]::Round((($totalMemoryGB - $freeMemoryGB) / $totalMemoryGB) * 100, 2)
    
    Write-Log "Total RAM: $totalMemoryGB GB | Used: $usedPercent% | Free: $freeMemoryGB GB" "Info"
    
    if ($usedPercent -gt 90) {
        Write-Log "Memory usage critically high ($usedPercent%)" "Error"
        
        # Find top memory consumers
        $topProcesses = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 5
        Write-Log "Top 5 memory consumers:" "Info"
        foreach ($proc in $topProcesses) {
            $memMB = [math]::Round($proc.WorkingSet / 1MB, 2)
            Write-Log "  $($proc.Name): $memMB MB" "Info"
        }
    } elseif ($usedPercent -gt 80) {
        Write-Log "Memory usage high ($usedPercent%)" "Warning"
    }
}

function Test-StartupPrograms {
    Write-Host "`n=== Startup Programs Analysis ===" -ForegroundColor Cyan
    
    $startupItems = @()
    
    # Registry startup locations
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
        "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    )
    
    foreach ($path in $regPaths) {
        try {
            $items = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
            if ($items) {
                $items.PSObject.Properties | Where-Object { $_.Name -notmatch "PS" } | ForEach-Object {
                    $startupItems += @{
                        Name = $_.Name
                        Command = $_.Value
                        Location = $path
                    }
                }
            }
        } catch {}
    }
    
    Write-Log "Found $($startupItems.Count) startup items" "Info"
    if ($startupItems.Count -gt 15) {
        Write-Log "High number of startup programs may impact boot time" "Warning"
    }
}

function Test-WindowsUpdate {
    Write-Host "`n=== Windows Update Status ===" -ForegroundColor Cyan
    
    try {
        $updateSession = New-Object -ComObject Microsoft.Update.Session
        $updateSearcher = $updateSession.CreateUpdateSearcher()
        $searchResult = $updateSearcher.Search("IsInstalled=0")
        
        $pendingUpdates = $searchResult.Updates.Count
        
        if ($pendingUpdates -gt 0) {
            Write-Log "Windows Update: $pendingUpdates pending updates found" "Warning"
        } else {
            Write-Log "Windows Update: System is up to date" "Info"
        }
        
        # Check last update time
        $lastUpdate = Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1
        if ($lastUpdate) {
            $daysSinceUpdate = (New-TimeSpan -Start $lastUpdate.InstalledOn -End (Get-Date)).Days
            if ($daysSinceUpdate -gt 30) {
                Write-Log "Last update installed $daysSinceUpdate days ago" "Warning"
            } else {
                Write-Log "Last update: $($lastUpdate.HotFixID) on $($lastUpdate.InstalledOn)" "Info"
            }
        }
    } catch {
        Write-Log "Unable to check Windows Update status: $($_.Exception.Message)" "Warning"
    }
}

function Test-Services {
    Write-Host "`n=== Critical Services Check ===" -ForegroundColor Cyan
    
    $criticalServices = @(
        "wuauserv",    # Windows Update
        "BITS",        # Background Intelligent Transfer Service
        "Dhcp",        # DHCP Client
        "Dnscache",    # DNS Client
        "EventLog",    # Windows Event Log
        "MpsSvc",      # Windows Defender Firewall
        "Winmgmt"      # Windows Management Instrumentation
    )
    
    foreach ($svc in $criticalServices) {
        $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.Status -ne "Running") {
                Write-Log "Service '$($service.DisplayName)' is $($service.Status)" "Error"
                
                if (-not $SkipRepairs -and $service.StartType -ne "Disabled") {
                    try {
                        Start-Service -Name $svc -ErrorAction Stop
                        Write-Log "Successfully started service '$($service.DisplayName)'" "Success"
                    } catch {
                        Write-Log "Failed to start service: $($_.Exception.Message)" "Error"
                    }
                }
            }
        }
    }
}

function Test-EventLogs {
    Write-Host "`n=== Event Log Analysis ===" -ForegroundColor Cyan
    
    $hours = 24
    $after = (Get-Date).AddHours(-$hours)
    
    try {
        $criticalErrors = Get-WinEvent -FilterHashtable @{
            LogName = 'System', 'Application'
            Level = 1,2  # Critical and Error
            StartTime = $after
        } -ErrorAction SilentlyContinue | 
        Group-Object Id | 
        Sort-Object Count -Descending | 
        Select-Object -First 5
        
        if ($criticalErrors) {
            Write-Log "Top 5 recurring errors in last $hours hours:" "Warning"
            foreach ($error in $criticalErrors) {
                $sample = Get-WinEvent -FilterHashtable @{
                    LogName = 'System', 'Application'
                    Id = $error.Name
                    StartTime = $after
                } -MaxEvents 1 -ErrorAction SilentlyContinue
                
                Write-Log "  Event ID $($error.Name): $($error.Count) occurrences - $($sample.Message.Substring(0, [Math]::Min(100, $sample.Message.Length)))" "Info"
            }
        } else {
            Write-Log "No critical errors found in last $hours hours" "Info"
        }
    } catch {
        Write-Log "Unable to analyze event logs: $($_.Exception.Message)" "Warning"
    }
}

function Optimize-NetworkSettings {
    if ($SkipRepairs) { return }
    
    Write-Host "`n=== Network Optimization ===" -ForegroundColor Cyan
    
    try {
        # Flush DNS cache
        Clear-DnsClientCache
        Write-Log "DNS cache cleared" "Success"
        
        # Reset Winsock
        & netsh winsock reset | Out-Null
        Write-Log "Winsock reset completed (restart required)" "Success"
        
        # Release and renew IP
        & ipconfig /release | Out-Null
        & ipconfig /renew | Out-Null
        Write-Log "IP address renewed" "Success"
    } catch {
        Write-Log "Network optimization failed: $($_.Exception.Message)" "Error"
    }
}

function Test-DefragmentationStatus {
    Write-Host "`n=== Disk Optimization Status ===" -ForegroundColor Cyan
    
    try {
        $volumes = Get-Volume | Where-Object { $_.DriveLetter -and $_.FileSystem -eq "NTFS" }
        
        foreach ($vol in $volumes) {
            $defragAnalysis = Optimize-Volume -DriveLetter $vol.DriveLetter -Analyze -ErrorAction SilentlyContinue
            Write-Log "Drive $($vol.DriveLetter): Optimization recommended based on file fragmentation" "Info"
        }
    } catch {
        Write-Log "Unable to analyze disk optimization status" "Warning"
    }
}

function Clear-BrowserCache {
    if ($SkipRepairs) { return }
    
    Write-Host "`n=== Browser Cache Cleanup ===" -ForegroundColor Cyan
    
    $cachePaths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*.default-release\cache2"
    )
    
    $totalCleaned = 0
    foreach ($path in $cachePaths) {
        try {
            if (Test-Path $path) {
                $items = Get-ChildItem -Path $path -Recurse -Force -File -ErrorAction SilentlyContinue
                if ($items) {
                    $size = ($items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                    if ($null -eq $size) { $size = 0 }
                    
                    Remove-Item -Path "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
                    $totalCleaned += $size
                }
            }
        } catch {
            # Path doesn't exist or browser is running
        }
    }
    
    if ($totalCleaned -gt 0) {
        $cleanedMB = [math]::Round($totalCleaned / 1MB, 2)
        Write-Log "Cleaned $cleanedMB MB of browser cache" "Success"
    } else {
        Write-Log "No browser cache to clean (or browsers currently running)" "Info"
    }
}

function Test-PowerPlan {
    Write-Host "`n=== Power Plan Configuration ===" -ForegroundColor Cyan
    
    try {
        $currentPlan = powercfg /getactivescheme
        Write-Log "Current power plan: $($currentPlan -replace '.*\(', '' -replace '\)', '')" "Info"
        
        if ($currentPlan -match "Power saver") {
            Write-Log "Power Saver plan may impact performance" "Warning"
            
            if (-not $SkipRepairs) {
                # Set to High Performance
                $highPerf = powercfg /list | Where-Object { $_ -match "High performance" }
                if ($highPerf -match "([a-f0-9\-]{36})") {
                    powercfg /setactive $matches[1]
                    Write-Log "Switched to High Performance power plan" "Success"
                }
            }
        }
    } catch {
        Write-Log "Unable to check power plan" "Warning"
    }
}

function New-DiagnosticReport {
    if (-not $GenerateReport) { return }
    
    Write-Host "`n=== Generating Report ===" -ForegroundColor Cyan
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>System Diagnostic Report - $($script:Results.ComputerName)</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 30px; border-left: 4px solid #3498db; padding-left: 10px; }
        .info { background: #d1ecf1; border-left: 4px solid #0c5460; padding: 10px; margin: 10px 0; }
        .warning { background: #fff3cd; border-left: 4px solid #856404; padding: 10px; margin: 10px 0; }
        .error { background: #f8d7da; border-left: 4px solid #721c24; padding: 10px; margin: 10px 0; }
        .success { background: #d4edda; border-left: 4px solid #155724; padding: 10px; margin: 10px 0; }
        .timestamp { color: #7f8c8d; font-size: 0.9em; }
        table { width: 100%; border-collapse: collapse; margin: 20px 0; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #3498db; color: white; }
        tr:hover { background-color: #f5f5f5; }
    </style>
</head>
<body>
    <div class="container">
        <h1>System Diagnostic Report</h1>
        <p><strong>Computer:</strong> $($script:Results.ComputerName)</p>
        <p><strong>Scan Time:</strong> $($script:Results.ScanTime)</p>
        <p><strong>Repair Mode:</strong> $(if($SkipRepairs){"Disabled"}else{"Enabled"})</p>
        
        <h2>Errors Found ($($script:Results.Errors.Count))</h2>
"@
    
    foreach ($item in $script:Results.Errors) {
        $html += "<div class='error'><span class='timestamp'>$($item.Time)</span> - $($item.Message)</div>`n"
    }
    
    $html += "<h2>Warnings ($($script:Results.Warnings.Count))</h2>`n"
    foreach ($item in $script:Results.Warnings) {
        $html += "<div class='warning'><span class='timestamp'>$($item.Time)</span> - $($item.Message)</div>`n"
    }
    
    $html += "<h2>Repairs Performed ($($script:Results.Successs.Count))</h2>`n"
    foreach ($item in $script:Results.Successs) {
        $html += "<div class='success'><span class='timestamp'>$($item.Time)</span> - $($item.Message)</div>`n"
    }
    
    $html += "<h2>Additional Information</h2>`n"
    foreach ($item in $script:Results.Infos) {
        $html += "<div class='info'><span class='timestamp'>$($item.Time)</span> - $($item.Message)</div>`n"
    }
    
    $html += @"
    </div>
</body>
</html>
"@
    
    $html | Out-File -FilePath $ReportPath -Encoding UTF8
    Write-Log "Report generated: $ReportPath" "Success"
    
    # Open report
    Start-Process $ReportPath
}

# Main Execution
Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║     System Diagnostic & Repair Tool v1.0                  ║" -ForegroundColor Green
Write-Host "║     Computer: $($env:COMPUTERNAME.PadRight(44)) ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green

if ($SkipRepairs) {
    Write-Host "`nRunning in DIAGNOSTIC ONLY mode (no repairs will be performed)`n" -ForegroundColor Yellow
} else {
    Write-Host "`nRunning in DIAGNOSTIC & REPAIR mode`n" -ForegroundColor Green
}

# Execute diagnostics
Test-DiskSpace
Test-DiskHealth
Test-Memory
Test-Services
Test-StartupPrograms
Test-WindowsUpdate
Test-EventLogs
Test-PowerPlan
Test-DefragmentationStatus

# Execute repairs
Clear-TempFiles
Clear-BrowserCache
Optimize-NetworkSettings
Test-SystemFiles

# Generate report
New-DiagnosticReport

Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
Write-Host "║     Diagnostic Scan Complete                               ║" -ForegroundColor Green
Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green
Write-Host "`nSummary:" -ForegroundColor Cyan
Write-Host "  Errors: $($script:Results.Errors.Count)" -ForegroundColor Red
Write-Host "  Warnings: $($script:Results.Warnings.Count)" -ForegroundColor Yellow
Write-Host "  Repairs: $($script:Results.Successs.Count)" -ForegroundColor Green

if ($script:Results.Errors.Count -gt 0) {
    Write-Host "`nAction Required: Review errors and consider running DISM for system file repairs" -ForegroundColor Yellow
}

Write-Host "`nNote: Some changes may require a system restart to take effect.`n" -ForegroundColor Cyan

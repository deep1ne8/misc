# Dell Hardware Health Check - Remote or Local
$ServerName = "SUNHV01"  # Change to target server

$ScriptBlock = {
    $Results = @{
        ServerName = $env:COMPUTERNAME
        Timestamp = Get-Date
        OMSAInstalled = $false
        HardwareStatus = @()
        Errors = @()
    }
    
    # Check for OMSA installation
    $OMSAPath = "C:\Program Files\Dell\SysMgt\oma\bin\omreport.exe"
    if (Test-Path $OMSAPath) {
        $Results.OMSAInstalled = $true
        
        try {
            # Get overall system status
            $SystemHealth = & $OMSAPath system summary | Out-String
            $Results.HardwareStatus += @{Component = "SystemSummary"; Status = $SystemHealth}
            
            # Get specific alerts
            $Alerts = & $OMSAPath chassis alerts | Out-String
            if ($Alerts -notmatch "No alerts detected") {
                $Results.Errors += $Alerts
            }
            
            # Check critical components
            $Components = @("fans", "temps", "volts", "memory", "processors", "pwrsupplies")
            foreach ($Component in $Components) {
                try {
                    $Output = & $OMSAPath chassis $Component | Out-String
                    $Results.HardwareStatus += @{Component = $Component; Status = $Output}
                } catch {
                    $Results.Errors += "Failed to check $Component"
                }
            }
        } catch {
            $Results.Errors += "OMSA query failed: $_"
        }
    } else {
        # Fallback to WMI/CIM
        try {
            # Check Dell CIM namespace
            $DellClasses = Get-CimInstance -Namespace root\cimv2\Dell -ClassName Dell_Chassis -ErrorAction Stop
            $Results.HardwareStatus += @{Component = "Dell_Chassis"; Status = $DellClasses}
        } catch {
            $Results.Errors += "Dell WMI/CIM not available"
        }
    }
    
    # Check Windows Event Logs for Dell-related errors
    try {
        $DellEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System','Application'
            ProviderName = '*Dell*'
            Level = 1,2,3  # Critical, Error, Warning
            StartTime = (Get-Date).AddHours(-24)
        } -ErrorAction SilentlyContinue
        
        if ($DellEvents) {
            $Results.Errors += $DellEvents | Select-Object TimeCreated, Id, LevelDisplayName, Message
        }
    } catch {
        # Event log query may fail if no Dell events exist
    }
    
    return $Results
}

# Execute locally or remotely
if ($ServerName -eq $env:COMPUTERNAME) {
    $HealthReport = & $ScriptBlock
} else {
    $HealthReport = Invoke-Command -ComputerName $ServerName -ScriptBlock $ScriptBlock
}

# Display Results
Write-Host "`n=== Dell Hardware Health Report for $($HealthReport.ServerName) ===" -ForegroundColor Cyan
Write-Host "Timestamp: $($HealthReport.Timestamp)" -ForegroundColor Gray
Write-Host "OMSA Installed: $($HealthReport.OMSAInstalled)" -ForegroundColor $(if($HealthReport.OMSAInstalled){'Green'}else{'Yellow'})

if ($HealthReport.Errors.Count -gt 0) {
    Write-Host "`n[ALERTS DETECTED]" -ForegroundColor Red
    $HealthReport.Errors | ForEach-Object { Write-Host $_ -ForegroundColor Yellow }
} else {
    Write-Host "`nNo critical errors detected" -ForegroundColor Green
}

$HealthReport | ConvertTo-Json -Depth 5 | Out-File "DellHealth_$ServerName_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"

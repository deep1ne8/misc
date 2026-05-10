<#
.SYNOPSIS
    Dell Hardware Health Check - Robust version with multiple detection methods
.DESCRIPTION
    Comprehensive Dell server hardware health monitoring with OMSA, WMI/CIM, and Event Log checks
.PARAMETER ComputerName
    Target server name (default: localhost)
.PARAMETER Credential
    PSCredential for remote connection
.PARAMETER OutputPath
    Path for JSON report output (default: current directory)
.EXAMPLE
    .\Get-DellHardwareHealth.ps1 -ComputerName "SUNHV01"
.EXAMPLE
    .\Get-DellHardwareHealth.ps1 -ComputerName "SUNHV01" -Credential (Get-Credential)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ComputerName = $env:COMPUTERNAME,
    
    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "."
)

#Requires -Version 5.1

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Colors = @{
        "INFO" = "Cyan"
        "SUCCESS" = "Green"
        "WARNING" = "Yellow"
        "ERROR" = "Red"
    }
    Write-Host "[$Level] $Message" -ForegroundColor $Colors[$Level]
}

$ScriptBlock = {
    param($VerbosePreference)
    
    $Results = [PSCustomObject]@{
        ServerName = $env:COMPUTERNAME
        Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        IsLocal = $true
        OMSAInstalled = $false
        OMSAVersion = $null
        HardwareStatus = [System.Collections.ArrayList]::new()
        ComponentDetails = [System.Collections.ArrayList]::new()
        Alerts = [System.Collections.ArrayList]::new()
        EventLogErrors = [System.Collections.ArrayList]::new()
        OverallStatus = "Unknown"
        ErrorMessages = [System.Collections.ArrayList]::new()
    }
    
    # Function to safely add items to ArrayLists
    function Add-Result {
        param($List, $Item)
        [void]$List.Add($Item)
    }
    
    # Check for OMSA installation (multiple possible paths)
    $OMSAPaths = @(
        "C:\Program Files\Dell\SysMgt\oma\bin\omreport.exe",
        "C:\Program Files (x86)\Dell\SysMgt\oma\bin\omreport.exe",
        "${env:ProgramFiles}\Dell\SysMgt\oma\bin\omreport.exe",
        "${env:ProgramFiles(x86)}\Dell\SysMgt\oma\bin\omreport.exe"
    )
    
    $OMSAPath = $OMSAPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
    
    if ($OMSAPath) {
        $Results.OMSAInstalled = $true
        Write-Verbose "OMSA found at: $OMSAPath"
        
        try {
            # Get OMSA version
            $VersionOutput = & $OMSAPath about 2>&1 | Out-String
            if ($VersionOutput -match "Version\s*:\s*([\d\.]+)") {
                $Results.OMSAVersion = $matches[1]
            }
            
            # Check OMSA service status
            $OMSAServices = Get-Service -Name "*DSM*", "*Dell*" -ErrorAction SilentlyContinue | 
                Where-Object { $_.DisplayName -match "Server Administrator" }
            
            $ServiceRunning = $OMSAServices | Where-Object { $_.Status -eq "Running" }
            
            if (-not $ServiceRunning) {
                Add-Result $Results.ErrorMessages "OMSA services not running"
                $Results.OverallStatus = "Warning"
            } else {
                # Get overall system health
                try {
                    $SystemSummary = & $OMSAPath system summary 2>&1 | Out-String
                    if ($LASTEXITCODE -eq 0) {
                        Add-Result $Results.HardwareStatus ([PSCustomObject]@{
                            Component = "SystemSummary"
                            Status = $SystemSummary.Trim()
                        })
                        
                        # Parse overall status
                        if ($SystemSummary -match "Main System Chassis\s+:\s+(\w+)") {
                            $Results.OverallStatus = $matches[1]
                        }
                    }
                } catch {
                    Add-Result $Results.ErrorMessages "Failed to get system summary: $($_.Exception.Message)"
                }
                
                # Get chassis alerts
                try {
                    $AlertsOutput = & $OMSAPath chassis alerts 2>&1 | Out-String
                    if ($LASTEXITCODE -eq 0) {
                        $AlertLines = $AlertsOutput -split "`n" | Where-Object { $_ -match "\S" -and $_ -notmatch "^(Alerts|----|\s*$)" }
                        
                        foreach ($Alert in $AlertLines) {
                            if ($Alert -notmatch "No alerts detected") {
                                Add-Result $Results.Alerts ([PSCustomObject]@{
                                    Type = "Chassis Alert"
                                    Message = $Alert.Trim()
                                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                                })
                            }
                        }
                    }
                } catch {
                    Add-Result $Results.ErrorMessages "Failed to get chassis alerts: $($_.Exception.Message)"
                }
                
                # Check critical components
                $Components = @(
                    @{Name = "fans"; DisplayName = "Cooling Fans"},
                    @{Name = "temps"; DisplayName = "Temperature Sensors"},
                    @{Name = "volts"; DisplayName = "Voltage Sensors"},
                    @{Name = "memory"; DisplayName = "Memory"},
                    @{Name = "processors"; DisplayName = "Processors"},
                    @{Name = "pwrsupplies"; DisplayName = "Power Supplies"},
                    @{Name = "batteries"; DisplayName = "Batteries"},
                    @{Name = "intrusion"; DisplayName = "Chassis Intrusion"}
                )
                
                foreach ($Comp in $Components) {
                    try {
                        $Output = & $OMSAPath chassis $($Comp.Name) 2>&1 | Out-String
                        if ($LASTEXITCODE -eq 0 -and $Output -match "\S") {
                            # Parse component status
                            $Status = "Unknown"
                            if ($Output -match "Status\s*:\s*(\w+)") {
                                $Status = $matches[1]
                            }
                            
                            Add-Result $Results.ComponentDetails ([PSCustomObject]@{
                                Component = $Comp.DisplayName
                                Status = $Status
                                Details = $Output.Trim()
                            })
                            
                            # Check for non-OK status
                            if ($Status -notmatch "^(Ok|Enabled|Ready)$" -and $Status -ne "Unknown") {
                                Add-Result $Results.Alerts ([PSCustomObject]@{
                                    Type = "Component Warning"
                                    Component = $Comp.DisplayName
                                    Message = "Status: $Status"
                                    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                                })
                            }
                        }
                    } catch {
                        Write-Verbose "Could not check $($Comp.DisplayName): $($_.Exception.Message)"
                    }
                }
            }
        } catch {
            Add-Result $Results.ErrorMessages "OMSA query failed: $($_.Exception.Message)"
        }
    } else {
        Write-Verbose "OMSA not found, trying WMI/CIM fallback"
        
        # Fallback to WMI/CIM
        try {
            $DellNamespaces = @("root\cimv2\Dell", "root\Dell")
            $NamespaceFound = $false
            
            foreach ($NS in $DellNamespaces) {
                try {
                    $TestClass = Get-CimInstance -Namespace $NS -ClassName CIM_ComputerSystem -ErrorAction Stop | Select-Object -First 1
                    if ($TestClass) {
                        $NamespaceFound = $true
                        Write-Verbose "Dell CIM namespace found: $NS"
                        
                        # Try to get Dell-specific classes
                        $DellClasses = @("Dell_Chassis", "Dell_Fan", "Dell_PowerSupply", "Dell_TemperatureProbe", "Dell_Sensor")
                        
                        foreach ($ClassName in $DellClasses) {
                            try {
                                $Instances = Get-CimInstance -Namespace $NS -ClassName $ClassName -ErrorAction SilentlyContinue
                                if ($Instances) {
                                    Add-Result $Results.ComponentDetails ([PSCustomObject]@{
                                        Component = $ClassName
                                        Status = "Available"
                                        Details = ($Instances | ConvertTo-Json -Compress -Depth 2)
                                    })
                                }
                            } catch {
                                # Class may not exist, continue
                            }
                        }
                        break
                    }
                } catch {
                    continue
                }
            }
            
            if (-not $NamespaceFound) {
                Add-Result $Results.ErrorMessages "Dell WMI/CIM namespace not available. Install Dell OpenManage Server Administrator."
            }
        } catch {
            Add-Result $Results.ErrorMessages "Dell WMI/CIM query failed: $($_.Exception.Message)"
        }
    }
    
    # Check Windows Event Logs for Dell-related errors
    try {
        $StartTime = (Get-Date).AddHours(-24)
        
        # Query System and Application logs separately to avoid provider wildcard issues
        $LogNames = @("System", "Application")
        $DellProviders = @("Dell", "DellSmbios", "DSM SA", "Server Administrator")
        
        foreach ($LogName in $LogNames) {
            foreach ($Provider in $DellProviders) {
                try {
                    $Events = Get-WinEvent -FilterHashtable @{
                        LogName = $LogName
                        ProviderName = $Provider
                        Level = 1,2,3  # Critical, Error, Warning
                        StartTime = $StartTime
                    } -ErrorAction SilentlyContinue -MaxEvents 50
                    
                    if ($Events) {
                        foreach ($Event in $Events) {
                            Add-Result $Results.EventLogErrors ([PSCustomObject]@{
                                TimeCreated = $Event.TimeCreated
                                Level = $Event.LevelDisplayName
                                EventId = $Event.Id
                                Provider = $Event.ProviderName
                                Message = $Event.Message.Substring(0, [Math]::Min(500, $Event.Message.Length))
                            })
                        }
                    }
                } catch {
                    # Provider may not exist in this log
                }
            }
        }
    } catch {
        Write-Verbose "Event log query warning: $($_.Exception.Message)"
    }
    
    # Determine final overall status
    if ($Results.Alerts.Count -gt 0) {
        $CriticalAlerts = $Results.Alerts | Where-Object { $_.Message -match "(Critical|Failed|Error)" }
        if ($CriticalAlerts) {
            $Results.OverallStatus = "Critical"
        } elseif ($Results.OverallStatus -eq "Unknown") {
            $Results.OverallStatus = "Warning"
        }
    } elseif ($Results.OverallStatus -eq "Unknown" -and $Results.OMSAInstalled) {
        $Results.OverallStatus = "Ok"
    }
    
    return $Results
}

# Main execution
try {
    Write-Log "Starting Dell Hardware Health Check" "INFO"
    Write-Log "Target: $ComputerName" "INFO"
    
    $IsLocal = ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq "localhost" -or $ComputerName -eq ".")
    
    if ($IsLocal) {
        Write-Log "Running local check..." "INFO"
        $HealthReport = & $ScriptBlock -VerbosePreference $VerbosePreference
    } else {
        # Test remote connectivity
        Write-Log "Testing remote connectivity..." "INFO"
        
        if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
            throw "Cannot reach $ComputerName. Server may be offline or unreachable."
        }
        
        # Test WinRM
        try {
            $TestSession = $Credential ? 
                (New-PSSession -ComputerName $ComputerName -Credential $Credential -ErrorAction Stop) :
                (New-PSSession -ComputerName $ComputerName -ErrorAction Stop)
            Remove-PSSession $TestSession
            Write-Log "Remote connection successful" "SUCCESS"
        } catch {
            throw "Cannot establish PowerShell remote session to $ComputerName. Ensure WinRM is enabled. Error: $($_.Exception.Message)"
        }
        
        # Execute remote check
        Write-Log "Running remote check..." "INFO"
        $InvokeParams = @{
            ComputerName = $ComputerName
            ScriptBlock = $ScriptBlock
            ArgumentList = $VerbosePreference
        }
        
        if ($Credential) {
            $InvokeParams['Credential'] = $Credential
        }
        
        $HealthReport = Invoke-Command @InvokeParams -ErrorAction Stop
        $HealthReport.IsLocal = $false
    }
    
    # Display Results
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Dell Hardware Health Report" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Server Name    : $($HealthReport.ServerName)" -ForegroundColor White
    Write-Host "Timestamp      : $($HealthReport.Timestamp)" -ForegroundColor Gray
    Write-Host "OMSA Installed : $($HealthReport.OMSAInstalled)" -ForegroundColor $(if($HealthReport.OMSAInstalled){'Green'}else{'Yellow'})
    
    if ($HealthReport.OMSAVersion) {
        Write-Host "OMSA Version   : $($HealthReport.OMSAVersion)" -ForegroundColor Gray
    }
    
    # Overall Status
    $StatusColor = switch ($HealthReport.OverallStatus) {
        "Ok" { "Green" }
        "Warning" { "Yellow" }
        "Critical" { "Red" }
        default { "Gray" }
    }
    Write-Host "Overall Status : $($HealthReport.OverallStatus)" -ForegroundColor $StatusColor
    Write-Host ""
    
    # Display Alerts
    if ($HealthReport.Alerts.Count -gt 0) {
        Write-Log "HARDWARE ALERTS DETECTED" "ERROR"
        Write-Host "----------------------------------------" -ForegroundColor Red
        $HealthReport.Alerts | ForEach-Object {
            Write-Host "  [$($_.Type)] $($_.Message)" -ForegroundColor Yellow
            if ($_.Component) {
                Write-Host "    Component: $($_.Component)" -ForegroundColor Gray
            }
        }
        Write-Host ""
    } else {
        Write-Log "No hardware alerts detected" "SUCCESS"
        Write-Host ""
    }
    
    # Display Component Status Summary
    if ($HealthReport.ComponentDetails.Count -gt 0) {
        Write-Host "Component Status Summary:" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor Cyan
        $HealthReport.ComponentDetails | ForEach-Object {
            $CompColor = if ($_.Status -match "^(Ok|Enabled|Ready|Available)$") { "Green" } else { "Yellow" }
            Write-Host "  $($_.Component.PadRight(20)) : $($_.Status)" -ForegroundColor $CompColor
        }
        Write-Host ""
    }
    
    # Display Recent Event Log Errors
    if ($HealthReport.EventLogErrors.Count -gt 0) {
        Write-Host "Recent Event Log Errors (Last 24 hours):" -ForegroundColor Yellow
        Write-Host "----------------------------------------" -ForegroundColor Yellow
        $HealthReport.EventLogErrors | Select-Object -First 5 | ForEach-Object {
            Write-Host "  [$($_.TimeCreated)] $($_.Level) - Event ID $($_.EventId)" -ForegroundColor Gray
            Write-Host "    $($_.Message.Substring(0, [Math]::Min(100, $_.Message.Length)))..." -ForegroundColor DarkGray
        }
        if ($HealthReport.EventLogErrors.Count -gt 5) {
            Write-Host "  ... and $($HealthReport.EventLogErrors.Count - 5) more events" -ForegroundColor DarkGray
        }
        Write-Host ""
    }
    
    # Display Error Messages
    if ($HealthReport.ErrorMessages.Count -gt 0) {
        Write-Log "Script Execution Warnings" "WARNING"
        Write-Host "----------------------------------------" -ForegroundColor Yellow
        $HealthReport.ErrorMessages | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # Export JSON Report
    $OutputFile = Join-Path $OutputPath "DellHealth_$($HealthReport.ServerName)_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    
    try {
        $HealthReport | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputFile -Encoding UTF8 -Force
        Write-Log "Report saved to: $OutputFile" "SUCCESS"
    } catch {
        Write-Log "Failed to save report: $($_.Exception.Message)" "ERROR"
    }
    
    # Return status code
    if ($HealthReport.OverallStatus -eq "Critical") {
        exit 2
    } elseif ($HealthReport.OverallStatus -eq "Warning" -or $HealthReport.Alerts.Count -gt 0) {
        exit 1
    } else {
        exit 0
    }
    
} catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" "ERROR"
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 99
}

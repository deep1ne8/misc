# Printer Supply Monitor PowerShell Script
# Based on the C# PrinterStatus application
# Requires SNMP to be enabled on target printers

param(
    [string]$OutputPath = ".\PrinterStatus.csv",
    [int]$TonerWarningThreshold = 20,
    [int]$TonerCriticalThreshold = 5,
    [int]$TimeoutMs = 5000,
    [switch]$ContinuousMode,
    [int]$IntervalSeconds = 300
)

# SNMP OIDs from the original C# code
$OIDs = @{
    DeviceStatus = "1.3.6.1.2.1.25.3.5.1.1.1"
    ModelName = "1.3.6.1.2.1.25.3.2.1.3.1"
    TonerNames = "1.3.6.1.2.1.43.11.1.1.6"
    TonerMaxLevels = "1.3.6.1.2.1.43.11.1.1.8"
    TonerCurrentLevels = "1.3.6.1.2.1.43.11.1.1.9"
    TrayNames = "1.3.6.1.2.1.43.8.2.1.18"
    TrayLevels = "1.3.6.1.2.1.43.8.2.1.10"
    TrayTypes = "1.3.6.1.2.1.43.8.2.1.2"
    DisplayMessages = "1.3.6.1.2.1.43.16.5"
}

# Device Status mappings
$DeviceStatusMap = @{
    1 = "Other"
    2 = "Processing"
    3 = "Idle" 
    4 = "Printing"
    5 = "Warmup"
}

# Function to test if printer is online
function Test-PrinterOnline {
    param([string]$IPAddress)
    
    try {
        # Use .NET Ping class for better compatibility
        $ping = New-Object System.Net.NetworkInformation.Ping
        $result = $ping.Send($IPAddress, 2000)  # 2 second timeout
        $ping.Dispose()
        return ($result.Status -eq "Success")
    }
    catch {
        return $false
    }
}

# Function to perform SNMP walk (simplified version)
function Invoke-SNMPWalk {
    param(
        [string]$IPAddress,
        [string]$Community = "public",
        [string]$OID
    )
    
    try {
        # Using snmpwalk command if available (requires Net-SNMP tools)
        $result = & snmpwalk -v1 -c $Community $IPAddress $OID 2>$null
        if ($result) {
            return $result | ForEach-Object {
                if ($_ -match '.*= (.*)$') {
                    $matches[1].Trim('"')
                }
            }
        }
    }
    catch {
        Write-Warning "SNMP walk failed for $IPAddress : $($_.Exception.Message)"
    }
    
    return @()
}

# Function to perform SNMP get (simplified version)
function Invoke-SNMPGet {
    param(
        [string]$IPAddress,
        [string]$Community = "public",
        [string]$OID
    )
    
    try {
        # Using snmpget command if available (requires Net-SNMP tools)
        $result = & snmpget -v1 -c $Community $IPAddress $OID 2>$null
        if ($result -and $result -match '.*= (.*)$') {
            return $matches[1].Trim('"')
        }
    }
    catch {
        Write-Warning "SNMP get failed for $IPAddress : $($_.Exception.Message)"
    }
    
    return $null
}

# Function to get printer information
function Get-PrinterInfo {
    param([string]$IPAddress)
    
    $printerInfo = @{
        IPAddress = $IPAddress
        Online = $false
        Status = "Offline"
        Model = ""
        TonerStatus = ""
        PaperStatus = ""
        DisplayMessages = ""
        LastScanned = Get-Date
    }
    
    # Test if printer is online
    if (-not (Test-PrinterOnline -IPAddress $IPAddress)) {
        return $printerInfo
    }
    
    $printerInfo.Online = $true
    
    # Get device status
    $deviceStatus = Invoke-SNMPGet -IPAddress $IPAddress -OID $OIDs.DeviceStatus
    if ($deviceStatus -and $DeviceStatusMap.ContainsKey([int]$deviceStatus)) {
        $printerInfo.Status = $DeviceStatusMap[[int]$deviceStatus]
    }
    
    # Get model name
    $modelName = Invoke-SNMPGet -IPAddress $IPAddress -OID $OIDs.ModelName
    if ($modelName) {
        $printerInfo.Model = $modelName -replace '\s+', ' '
    }
    
    # Get toner information
    $tonerNames = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.TonerNames
    $tonerMaxes = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.TonerMaxLevels
    $tonerCurrents = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.TonerCurrentLevels
    
    if ($tonerNames.Count -eq $tonerMaxes.Count -and $tonerNames.Count -eq $tonerCurrents.Count) {
        $tonerStatuses = @()
        for ($i = 0; $i -lt $tonerNames.Count; $i++) {
            try {
                $max = [int]$tonerMaxes[$i]
                $current = [int]$tonerCurrents[$i]
                
                # Handle special case where -3 means full
                if ($current -eq -3) {
                    $current = $max
                }
                
                $percentRemaining = [Math]::Round(($current / $max) * 100, 0)
                $tonerName = $tonerNames[$i]
                
                # Remove semicolon and everything after it
                $semicolonIndex = $tonerName.IndexOf(';')
                if ($semicolonIndex -gt 0) {
                    $tonerName = $tonerName.Substring(0, $semicolonIndex)
                }
                
                # Determine status level
                $statusLevel = "OK"
                if ($percentRemaining -le $TonerCriticalThreshold) {
                    $statusLevel = "CRITICAL"
                } elseif ($percentRemaining -le $TonerWarningThreshold) {
                    $statusLevel = "WARNING"
                }
                
                $tonerStatuses += "$percentRemaining% $tonerName ($statusLevel)"
            }
            catch {
                Write-Warning "Error processing toner $i for $IPAddress"
            }
        }
        $printerInfo.TonerStatus = $tonerStatuses -join "`r`n"
    }
    
    # Get paper tray information
    $trayNames = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.TrayNames
    $trayLevels = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.TrayLevels
    $trayTypes = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.TrayTypes
    
    if ($trayNames.Count -eq $trayLevels.Count -and $trayNames.Count -eq $trayTypes.Count) {
        $trayStatuses = @()
        for ($i = 0; $i -lt $trayNames.Count; $i++) {
            try {
                $level = [int]$trayLevels[$i]
                $type = [int]$trayTypes[$i]
                $name = $trayNames[$i]
                
                switch ($level) {
                    0 { 
                        if ($type -eq 4 -or $type -eq 5) {
                            $trayStatuses += "$name (Feed Tray)"
                        } else {
                            $trayStatuses += "$name (Empty)"
                        }
                    }
                    -1 { $trayStatuses += "$name (Status: Unknown)" }
                    -2 { $trayStatuses += "$name (Tray Open)" }
                    -3 { $trayStatuses += "$name (Paper Present)" }
                    default { $trayStatuses += "$name ($level Pages in Tray)" }
                }
            }
            catch {
                Write-Warning "Error processing tray $i for $IPAddress"
            }
        }
        $printerInfo.PaperStatus = $trayStatuses -join "`r`n"
    }
    
    # Get display messages
    $displayMessages = Invoke-SNMPWalk -IPAddress $IPAddress -OID $OIDs.DisplayMessages
    if ($displayMessages) {
        $cleanMessages = @()
        foreach ($message in $displayMessages) {
            # Remove hex patterns and empty lines
            $cleanMessage = $message -replace '^([0-9A-F]{2}\s){1,}([0-9A-F]{2})$', ''
            $cleanMessage = $cleanMessage -replace '^\s*$', ''
            if ($cleanMessage) {
                $cleanMessages += $cleanMessage
            }
        }
        $printerInfo.DisplayMessages = $cleanMessages -join "`r`n"
    }
    
    return $printerInfo
}

# Function to get all network printers
function Get-NetworkPrinters {
    try {
        $printers = Get-CimInstance -ClassName Win32_Printer | Where-Object { 
            $_.PortName -match '^\d+\.\d+\.\d+\.\d+$' -or 
            $_.PortName -like 'IP_*' 
        }
        
        $networkPrinters = @()
        foreach ($printer in $printers) {
            $ipAddress = ""
            
            # Try to extract IP from port name
            if ($printer.PortName -match '^\d+\.\d+\.\d+\.\d+$') {
                $ipAddress = $printer.PortName
            } else {
                # Get IP from TCP/IP printer port
                $port = Get-CimInstance -ClassName Win32_TCPIPPrinterPort | 
                        Where-Object { $_.Name -eq $printer.PortName }
                if ($port) {
                    $ipAddress = $port.HostAddress
                }
            }
            
            if ($ipAddress) {
                $networkPrinters += @{
                    Name = $printer.Name
                    IPAddress = $ipAddress
                    IsDefault = $printer.Default
                }
            }
        }
        
        return $networkPrinters
    }
    catch {
        Write-Error "Failed to get network printers: $($_.Exception.Message)"
        return @()
    }
}

# Function to export results to CSV
function Export-PrinterStatus {
    param(
        [array]$PrinterData,
        [string]$FilePath
    )
    
    $csvData = @()
    foreach ($printer in $PrinterData) {
        $csvData += [PSCustomObject]@{
            PrinterName = $printer.Name
            IPAddress = $printer.IPAddress
            Online = $printer.Info.Online
            Status = $printer.Info.Status
            Model = $printer.Info.Model
            TonerStatus = $printer.Info.TonerStatus
            PaperStatus = $printer.Info.PaperStatus
            DisplayMessages = $printer.Info.DisplayMessages
            LastScanned = $printer.Info.LastScanned
            IsDefault = $printer.IsDefault
        }
    }
    
    $csvData | Export-Csv -Path $FilePath -NoTypeInformation
    Write-Host "Results exported to: $FilePath" -ForegroundColor Green
}

# Function to display results in console
function Show-PrinterStatus {
    param([array]$PrinterData)
    
    Write-Host "`n=== Printer Supply Status ===" -ForegroundColor Cyan
    Write-Host "Scan completed at: $(Get-Date)" -ForegroundColor Gray
    Write-Host ""
    
    foreach ($printer in $PrinterData) {
        $color = if ($printer.Info.Online) { "Green" } else { "Red" }
        Write-Host "Printer: $($printer.Name)" -ForegroundColor White
        Write-Host "  IP: $($printer.IPAddress)" -ForegroundColor Gray
        Write-Host "  Status: $($printer.Info.Status)" -ForegroundColor $color
        
        if ($printer.Info.Model) {
            Write-Host "  Model: $($printer.Info.Model)" -ForegroundColor Gray
        }
        
        if ($printer.Info.TonerStatus) {
            Write-Host "  Toner Status:" -ForegroundColor Yellow
            $printer.Info.TonerStatus -split "`r`n" | ForEach-Object {
                $tonerColor = "White"
                if ($_ -match "CRITICAL") { $tonerColor = "Red" }
                elseif ($_ -match "WARNING") { $tonerColor = "Yellow" }
                Write-Host "    $_" -ForegroundColor $tonerColor
            }
        }
        
        if ($printer.Info.PaperStatus) {
            Write-Host "  Paper Status:" -ForegroundColor Cyan
            $printer.Info.PaperStatus -split "`r`n" | ForEach-Object {
                Write-Host "    $_" -ForegroundColor Gray
            }
        }
        
        if ($printer.Info.DisplayMessages) {
            Write-Host "  Display Messages:" -ForegroundColor Magenta
            $printer.Info.DisplayMessages -split "`r`n" | ForEach-Object {
                Write-Host "    $_" -ForegroundColor Gray
            }
        }
        
        Write-Host ""
    }
}

# Main scanning function
function Start-PrinterScan {
    Write-Host "Discovering network printers..." -ForegroundColor Cyan
    $networkPrinters = Get-NetworkPrinters
    
    if ($networkPrinters.Count -eq 0) {
        Write-Warning "No network printers found."
        return
    }
    
    Write-Host "Found $($networkPrinters.Count) network printer(s)" -ForegroundColor Green
    
    $printerData = @()
    $current = 0
    
    foreach ($printer in $networkPrinters) {
        $current++
        Write-Host "Scanning printer $current of $($networkPrinters.Count): $($printer.Name) ($($printer.IPAddress))" -ForegroundColor Yellow
        
        $info = Get-PrinterInfo -IPAddress $printer.IPAddress
        
        $printerData += @{
            Name = $printer.Name
            IPAddress = $printer.IPAddress
            IsDefault = $printer.IsDefault
            Info = $info
        }
    }
    
    # Display results
    Show-PrinterStatus -PrinterData $printerData
    
    # Export to CSV
    Export-PrinterStatus -PrinterData $printerData -FilePath $OutputPath
    
    return $printerData
}

# Main script execution
try {
    Write-Host "Printer Supply Monitor" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    Write-Host "Warning Threshold: $TonerWarningThreshold%" -ForegroundColor Yellow
    Write-Host "Critical Threshold: $TonerCriticalThreshold%" -ForegroundColor Red
    Write-Host ""
    
    # Check if SNMP tools are available
    $snmpAvailable = $false
    try {
        $null = & snmpget --version 2>$null
        $snmpAvailable = $true
    }
    catch {
        Write-Warning "Net-SNMP tools not found. Install from http://www.net-snmp.org/ for full functionality."
        Write-Host "Continuing with basic printer discovery only..." -ForegroundColor Yellow
    }
    
    if ($ContinuousMode) {
        Write-Host "Starting continuous monitoring mode (interval: $IntervalSeconds seconds)" -ForegroundColor Green
        Write-Host "Press Ctrl+C to stop..." -ForegroundColor Gray
        
        do {
            Start-PrinterScan
            Write-Host "Waiting $IntervalSeconds seconds before next scan..." -ForegroundColor Gray
            Start-Sleep -Seconds $IntervalSeconds
        } while ($true)
    }
    else {
        Start-PrinterScan
    }
}
catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}

# Usage examples:
# .\PrinterSupplyMonitor.ps1                                    # Single scan
# .\PrinterSupplyMonitor.ps1 -ContinuousMode -IntervalSeconds 600  # Continuous every 10 minutes  
# .\PrinterSupplyMonitor.ps1 -TonerWarningThreshold 15 -TonerCriticalThreshold 3  # Custom thresholds
# .\PrinterSupplyMonitor.ps1 -OutputPath "C:\Reports\PrinterStatus.csv"  # Custom output path

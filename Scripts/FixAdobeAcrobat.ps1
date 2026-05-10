#Requires -Version 5.1
<#
.SYNOPSIS
    Deep diagnostic tool to identify the root cause of Adobe Acrobat hangs.

.DESCRIPTION
    Monitors Adobe Acrobat processes in real-time to capture thread states, file handles,
    event logs, and operations in progress when hangs occur. Provides actionable intelligence
    on what's actually causing the freeze.

.PARAMETER Mode
    Monitor: Watch for hangs in real-time
    Analyze: Analyze existing hang/crash dumps
    PostMortem: Analyze after a hang has occurred

.EXAMPLE
    .\Diagnose-AdobeHang.ps1 -Mode Monitor
    .\Diagnose-AdobeHang.ps1 -Mode PostMortem

.NOTES
    Author: Earl's Automation Suite
    Version: 1.0
    Requires: Admin rights for full diagnostics
#>

[CmdletBinding()]
param(
    [ValidateSet('Monitor', 'Analyze', 'PostMortem')]
    [string]$Mode = 'PostMortem',
    
    [int]$MonitorDurationSeconds = 300,
    
    [string]$OutputPath = "$env:TEMP\AdobeDiagnostics"
)

$ErrorActionPreference = 'Continue'

# Create output directory
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
}

$Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$LogFile = Join-Path $OutputPath "AdobeDiag_$Timestamp.log"

#region Helper Functions

function Write-DiagLog {
    param(
        [string]$Message,
        [string]$Level = 'INFO',
        [switch]$NoConsole
    )
    
    $Color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARNING' { 'Yellow' }
        'SUCCESS' { 'Green' }
        'CRITICAL' { 'Magenta' }
        default { 'Cyan' }
    }
    
    $LogTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogMessage = "[$LogTimestamp] [$Level] $Message"
    
    # Write to file
    Add-Content -Path $LogFile -Value $LogMessage
    
    # Write to console
    if (-not $NoConsole) {
        Write-Host $LogMessage -ForegroundColor $Color
    }
}

function Test-IsAdmin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    return $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-AdobeProcessDetails {
    $AdobeProcessNames = @('Acrobat', 'AcroRd32', 'AcroCEF', 'RdrCEF', 'AdobeCollabSync')
    
    $Processes = Get-Process | Where-Object { 
        $ProcessName = $_.ProcessName
        $AdobeProcessNames | Where-Object { $ProcessName -like "*$_*" }
    } | Select-Object Id, ProcessName, Responding, 
                     @{N='CPU(s)';E={[math]::Round($_.CPU,2)}},
                     @{N='Memory(MB)';E={[math]::Round($_.WorkingSet64/1MB,2)}},
                     @{N='Threads';E={$_.Threads.Count}},
                     @{N='Handles';E={$_.HandleCount}},
                     StartTime
    
    return $Processes
}

function Get-HungThreadInfo {
    param([int]$ProcessId)
    
    Write-DiagLog "Analyzing threads for PID $ProcessId..."
    
    try {
        $Process = Get-Process -Id $ProcessId -ErrorAction Stop
        $ThreadInfo = @()
        
        foreach ($Thread in $Process.Threads) {
            $ThreadState = $Thread.ThreadState
            $WaitReason = $Thread.WaitReason
            
            $ThreadInfo += [PSCustomObject]@{
                ThreadId = $Thread.Id
                State = $ThreadState
                WaitReason = if ($ThreadState -eq 'Wait') { $WaitReason } else { 'N/A' }
                StartTime = $Thread.StartTime
                CPUTime = $Thread.TotalProcessorTime
                Priority = $Thread.PriorityLevel
            }
        }
        
        return $ThreadInfo
    }
    catch {
        Write-DiagLog "Could not analyze threads: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Get-FileHandles {
    param([int]$ProcessId)
    
    Write-DiagLog "Retrieving file handles for PID $ProcessId..."
    
    try {
        # Use handle.exe if available, otherwise use PowerShell method
        $HandleExe = Get-Command handle.exe -ErrorAction SilentlyContinue
        
        if ($HandleExe) {
            $Output = & handle.exe -p $ProcessId -nobanner 2>&1
            $Files = $Output | Where-Object { $_ -match '\\' -and $_ -notmatch 'HKEY' } | 
                     ForEach-Object { 
                         if ($_ -match ':\s+(.+)$') { 
                             $matches[1] 
                         } 
                     }
            return $Files | Select-Object -Unique
        }
        else {
            # Fallback: Get handles via WMI (limited info)
            Write-DiagLog "handle.exe not found - using limited WMI method" -Level WARNING
            $Process = Get-Process -Id $ProcessId
            return @("Handle count: $($Process.HandleCount)", "Use Sysinternals handle.exe for detailed file handles")
        }
    }
    catch {
        Write-DiagLog "Could not retrieve file handles: $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function Get-EventLogErrors {
    param([int]$HoursBack = 2)
    
    Write-DiagLog "Scanning event logs for Adobe-related errors (last $HoursBack hours)..."
    
    $StartTime = (Get-Date).AddHours(-$HoursBack)
    $Events = @()
    
    # Application errors
    try {
        $AppEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'Application'
            Level = 2,3  # Error, Warning
            StartTime = $StartTime
        } -ErrorAction SilentlyContinue | 
        Where-Object { 
            $_.Message -match 'Adobe|Acrobat|AcroRd|PDF' 
        } | Select-Object -First 20
        
        $Events += $AppEvents
    }
    catch {
        Write-DiagLog "No application errors found or cannot access event log" -Level WARNING
    }
    
    # System errors related to drivers/services
    try {
        $SysEvents = Get-WinEvent -FilterHashtable @{
            LogName = 'System'
            Level = 2,3
            StartTime = $StartTime
        } -ErrorAction SilentlyContinue | 
        Where-Object { 
            $_.Message -match 'Adobe|print|driver|timeout|hang' 
        } | Select-Object -First 10
        
        $Events += $SysEvents
    }
    catch {
        Write-DiagLog "No system errors found or cannot access event log" -Level WARNING
    }
    
    return $Events
}

function Get-PrinterDriverInfo {
    Write-DiagLog "Checking printer configuration (common hang cause)..."
    
    try {
        $Printers = Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published
        $DefaultPrinter = Get-CimInstance -ClassName Win32_Printer | Where-Object { $_.Default -eq $true }
        
        return [PSCustomObject]@{
            Printers = $Printers
            DefaultPrinter = $DefaultPrinter.Name
            TotalPrinters = $Printers.Count
        }
    }
    catch {
        Write-DiagLog "Could not retrieve printer info: $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function Get-NetworkDriveInfo {
    Write-DiagLog "Checking network drives (can cause timeouts)..."
    
    try {
        $NetworkDrives = Get-PSDrive -PSProvider FileSystem | 
                        Where-Object { $_.DisplayRoot -like '\\*' } |
                        Select-Object Name, DisplayRoot, 
                                     @{N='Used(GB)';E={[math]::Round($_.Used/1GB,2)}},
                                     @{N='Free(GB)';E={[math]::Round($_.Free/1GB,2)}}
        
        # Test connectivity
        foreach ($Drive in $NetworkDrives) {
            $TestPath = "$($Drive.Name):\"
            $Accessible = Test-Path $TestPath -ErrorAction SilentlyContinue
            $Drive | Add-Member -NotePropertyName 'Accessible' -NotePropertyValue $Accessible -Force
        }
        
        return $NetworkDrives
    }
    catch {
        Write-DiagLog "Could not retrieve network drive info: $($_.Exception.Message)" -Level WARNING
        return $null
    }
}

function Get-RecentPDFFiles {
    Write-DiagLog "Finding recently accessed PDF files..."
    
    $RecentPaths = @(
        "$env:USERPROFILE\Documents\*.pdf",
        "$env:USERPROFILE\Downloads\*.pdf",
        "$env:USERPROFILE\Desktop\*.pdf"
    )
    
    $RecentPDFs = foreach ($Path in $RecentPaths) {
        Get-ChildItem -Path $Path -ErrorAction SilentlyContinue |
        Where-Object { $_.LastAccessTime -gt (Get-Date).AddHours(-24) } |
        Select-Object FullName, 
                     @{N='Size(MB)';E={[math]::Round($_.Length/1MB,2)}},
                     LastAccessTime |
        Sort-Object LastAccessTime -Descending
    }
    
    return $RecentPDFs | Select-Object -First 20
}

function Get-AdobePlugins {
    Write-DiagLog "Checking Adobe plugins (can cause conflicts)..."
    
    $PluginPaths = @(
        "${env:ProgramFiles}\Adobe\Acrobat DC\Acrobat\plug_ins",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat DC\Acrobat\plug_ins",
        "${env:ProgramFiles}\Adobe\Acrobat Reader DC\Reader\plug_ins",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC\Reader\plug_ins",
        "$env:APPDATA\Adobe\Acrobat\DC\Plug-ins"
    )
    
    $Plugins = foreach ($Path in $PluginPaths) {
        if (Test-Path $Path) {
            Get-ChildItem -Path $Path -Filter *.api -ErrorAction SilentlyContinue |
            Select-Object Name, 
                         @{N='Size(KB)';E={[math]::Round($_.Length/1KB,2)}},
                         LastWriteTime,
                         Directory
        }
    }
    
    return $Plugins
}

function Get-AdobePreferences {
    Write-DiagLog "Checking Adobe preferences files..."
    
    $PrefPaths = @(
        "$env:APPDATA\Adobe\Acrobat\DC\Preferences",
        "$env:APPDATA\Adobe\Acrobat\DC\Security"
    )
    
    $Prefs = foreach ($Path in $PrefPaths) {
        if (Test-Path $Path) {
            Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue |
            Select-Object FullName, Length, LastWriteTime
        }
    }
    
    return $Prefs
}

function Invoke-DeepAnalysis {
    param([int]$ProcessId)
    
    Write-DiagLog "=== DEEP ANALYSIS MODE ===" -Level CRITICAL
    
    $Analysis = @{
        ProcessDetails = $null
        ThreadAnalysis = $null
        FileHandles = $null
        StackTrace = $null
        Responding = $null
    }
    
    try {
        $Process = Get-Process -Id $ProcessId -ErrorAction Stop
        
        $Analysis.ProcessDetails = [PSCustomObject]@{
            Name = $Process.ProcessName
            PID = $Process.Id
            Responding = $Process.Responding
            CPUPercent = [math]::Round((Get-Counter "\\Process($($Process.ProcessName))\\% Processor Time" -ErrorAction SilentlyContinue).CounterSamples.CookedValue, 2)
            MemoryMB = [math]::Round($Process.WorkingSet64/1MB, 2)
            Threads = $Process.Threads.Count
            Handles = $Process.HandleCount
            StartTime = $Process.StartTime
            Runtime = ((Get-Date) - $Process.StartTime).ToString()
        }
        
        $Analysis.Responding = $Process.Responding
        $Analysis.ThreadAnalysis = Get-HungThreadInfo -ProcessId $ProcessId
        $Analysis.FileHandles = Get-FileHandles -ProcessId $ProcessId
        
        # Identify hung threads
        if ($Analysis.ThreadAnalysis) {
            $HungThreads = $Analysis.ThreadAnalysis | Where-Object { 
                $_.State -eq 'Wait' -and $_.WaitReason -in @('UserRequest', 'Executive', 'Suspended')
            }
            
            if ($HungThreads) {
                Write-DiagLog "FOUND $($HungThreads.Count) POTENTIALLY HUNG THREAD(S):" -Level CRITICAL
                $HungThreads | ForEach-Object {
                    Write-DiagLog "  Thread $($_.ThreadId): State=$($_.State), WaitReason=$($_.WaitReason)" -Level WARNING
                }
            }
        }
        
        # Analyze file handles for clues
        if ($Analysis.FileHandles) {
            $PDFHandles = $Analysis.FileHandles | Where-Object { $_ -match '\.pdf$' }
            $NetworkHandles = $Analysis.FileHandles | Where-Object { $_ -match '^\\\\' }
            
            if ($PDFHandles) {
                Write-DiagLog "OPEN PDF FILES:" -Level CRITICAL
                $PDFHandles | ForEach-Object { Write-DiagLog "  $_" -Level WARNING }
            }
            
            if ($NetworkHandles) {
                Write-DiagLog "NETWORK FILE HANDLES (potential timeout):" -Level CRITICAL
                $NetworkHandles | ForEach-Object { Write-DiagLog "  $_" -Level WARNING }
            }
        }
    }
    catch {
        Write-DiagLog "Deep analysis failed: $($_.Exception.Message)" -Level ERROR
    }
    
    return $Analysis
}

#endregion

#region Main Execution

Write-DiagLog "=== ADOBE ACROBAT HANG DIAGNOSTIC TOOL ===" -Level SUCCESS
Write-DiagLog "Mode: $Mode"
Write-DiagLog "Output Directory: $OutputPath"
Write-DiagLog "Log File: $LogFile"
Write-DiagLog ""

if (-not (Test-IsAdmin)) {
    Write-DiagLog "WARNING: Not running as Administrator - some diagnostics may be limited" -Level WARNING
    Write-DiagLog ""
}

$Report = [PSCustomObject]@{
    Timestamp = Get-Date
    Mode = $Mode
    ProcessSnapshot = $null
    EventLogErrors = $null
    PrinterInfo = $null
    NetworkDrives = $null
    RecentPDFs = $null
    Plugins = $null
    DeepAnalysis = $null
    RootCauseAnalysis = @()
}

switch ($Mode) {
    'Monitor' {
        Write-DiagLog "Monitoring Adobe Acrobat for $MonitorDurationSeconds seconds..."
        Write-DiagLog "Open Adobe Acrobat and reproduce the hang issue"
        Write-DiagLog ""
        
        $MonitorStart = Get-Date
        $HangDetected = $false
        
        while (((Get-Date) - $MonitorStart).TotalSeconds -lt $MonitorDurationSeconds) {
            $Processes = Get-AdobeProcessDetails
            
            if ($Processes) {
                $NotResponding = $Processes | Where-Object { -not $_.Responding }
                
                if ($NotResponding -and -not $HangDetected) {
                    Write-DiagLog "HANG DETECTED!" -Level CRITICAL
                    Write-DiagLog "Process: $($NotResponding.ProcessName) (PID: $($NotResponding.Id))" -Level CRITICAL
                    
                    $HangDetected = $true
                    $Report.DeepAnalysis = Invoke-DeepAnalysis -ProcessId $NotResponding.Id
                    
                    Write-DiagLog "Capturing full diagnostic snapshot..." -Level CRITICAL
                    break
                }
                
                # Display current status
                Write-Host "`r[$(Get-Date -Format 'HH:mm:ss')] Monitoring... Processes: $($Processes.Count)" -NoNewline
            }
            
            Start-Sleep -Seconds 2
        }
        
        if (-not $HangDetected) {
            Write-DiagLog "`nNo hang detected during monitoring period" -Level SUCCESS
            Write-DiagLog "Consider running in PostMortem mode after a hang occurs" -Level WARNING
        }
    }
    
    'PostMortem' {
        Write-DiagLog "Analyzing system state after hang..."
        Write-DiagLog ""
        
        # Get current Adobe processes
        $Report.ProcessSnapshot = Get-AdobeProcessDetails
        
        if ($Report.ProcessSnapshot) {
            Write-DiagLog "Found Adobe Processes:" -Level SUCCESS
            $Report.ProcessSnapshot | Format-Table -AutoSize | Out-String | Write-DiagLog -NoConsole
            
            # Deep analysis on non-responding processes
            $NotResponding = $Report.ProcessSnapshot | Where-Object { -not $_.Responding }
            if ($NotResponding) {
                Write-DiagLog "NON-RESPONDING PROCESS DETECTED!" -Level CRITICAL
                $Report.DeepAnalysis = Invoke-DeepAnalysis -ProcessId $NotResponding[0].Id
            }
        }
        else {
            Write-DiagLog "No Adobe processes currently running" -Level WARNING
            Write-DiagLog "Analyzing recent history and system state..." -Level WARNING
        }
        
        # Gather environmental factors
        $Report.EventLogErrors = Get-EventLogErrors
        $Report.PrinterInfo = Get-PrinterDriverInfo
        $Report.NetworkDrives = Get-NetworkDriveInfo
        $Report.RecentPDFs = Get-RecentPDFFiles
        $Report.Plugins = Get-AdobePlugins
        
        # Display findings
        Write-DiagLog ""
        Write-DiagLog "=== EVENT LOG ANALYSIS ===" -Level SUCCESS
        if ($Report.EventLogErrors) {
            Write-DiagLog "Found $($Report.EventLogErrors.Count) relevant event(s):" -Level WARNING
            $Report.EventLogErrors | Select-Object -First 5 | ForEach-Object {
                Write-DiagLog "  [$($_.LevelDisplayName)] $($_.TimeCreated): $($_.Message.Substring(0, [Math]::Min(100, $_.Message.Length)))..." -Level WARNING
            }
        }
        else {
            Write-DiagLog "No relevant errors found in event logs" -Level SUCCESS
        }
        
        Write-DiagLog ""
        Write-DiagLog "=== PRINTER ANALYSIS ===" -Level SUCCESS
        if ($Report.PrinterInfo) {
            Write-DiagLog "Default Printer: $($Report.PrinterInfo.DefaultPrinter)" -Level SUCCESS
            Write-DiagLog "Total Printers: $($Report.PrinterInfo.TotalPrinters)" -Level SUCCESS
            
            # Check for problematic drivers
            $ProblematicDrivers = $Report.PrinterInfo.Printers | Where-Object { 
                $_.DriverName -match 'Generic|Microsoft XPS|Fax|OneNote'
            }
            if ($ProblematicDrivers) {
                Write-DiagLog "POTENTIALLY PROBLEMATIC PRINTER DRIVERS:" -Level WARNING
                $ProblematicDrivers | ForEach-Object {
                    Write-DiagLog "  $($_.Name): $($_.DriverName)" -Level WARNING
                }
            }
        }
        
        Write-DiagLog ""
        Write-DiagLog "=== NETWORK DRIVE ANALYSIS ===" -Level SUCCESS
        if ($Report.NetworkDrives) {
            $InaccessibleDrives = $Report.NetworkDrives | Where-Object { -not $_.Accessible }
            if ($InaccessibleDrives) {
                Write-DiagLog "INACCESSIBLE NETWORK DRIVES (can cause timeouts):" -Level CRITICAL
                $InaccessibleDrives | ForEach-Object {
                    Write-DiagLog "  $($_.Name): $($_.DisplayRoot)" -Level WARNING
                }
            }
            else {
                Write-DiagLog "All network drives accessible" -Level SUCCESS
            }
        }
        
        Write-DiagLog ""
        Write-DiagLog "=== RECENT PDF FILES ===" -Level SUCCESS
        if ($Report.RecentPDFs) {
            Write-DiagLog "Recently accessed PDFs (last 24h):" -Level SUCCESS
            $Report.RecentPDFs | Select-Object -First 10 | ForEach-Object {
                $SizeMB = $_.'Size(MB)'
                Write-DiagLog "  $($_.FullName) [$SizeMB MB]" -Level SUCCESS
            }
            
            # Flag large PDFs
            $LargePDFs = $Report.RecentPDFs | Where-Object { $_.'Size(MB)' -gt 50 }
            if ($LargePDFs) {
                Write-DiagLog "LARGE PDF FILES (can cause performance issues):" -Level WARNING
                $LargePDFs | ForEach-Object {
                    $SizeMB = $_.'Size(MB)'
                    Write-DiagLog "  $($_.FullName) [$SizeMB MB]" -Level WARNING
                }
            }
        }
        
        Write-DiagLog ""
        Write-DiagLog "=== PLUGIN ANALYSIS ===" -Level SUCCESS
        if ($Report.Plugins) {
            Write-DiagLog "Found $($Report.Plugins.Count) plugin(s)" -Level SUCCESS
            $Report.Plugins | ForEach-Object {
                Write-DiagLog "  $($_.Name)" -Level SUCCESS
            }
        }
    }
    
    'Analyze' {
        Write-DiagLog "Analyzing existing dumps and logs..."
        # Placeholder for dump file analysis
        Write-DiagLog "This mode requires minidump files - not yet implemented" -Level WARNING
    }
}

# Root Cause Analysis
Write-DiagLog ""
Write-DiagLog "=== ROOT CAUSE ANALYSIS ===" -Level CRITICAL

if ($Report.DeepAnalysis) {
    if (-not $Report.DeepAnalysis.Responding) {
        $Report.RootCauseAnalysis += "Process is NOT RESPONDING - confirmed hang"
    }
    
    if ($Report.DeepAnalysis.FileHandles) {
        $NetworkFiles = $Report.DeepAnalysis.FileHandles | Where-Object { $_ -match '^\\\\' }
        if ($NetworkFiles) {
            $Report.RootCauseAnalysis += "LIKELY CAUSE: Network file access - possible timeout on: $($NetworkFiles[0])"
        }
        
        $LargePDF = $Report.RecentPDFs | Where-Object { $_.'Size(MB)' -gt 100 } | Select-Object -First 1
        if ($LargePDF) {
            $PDFSize = $LargePDF.'Size(MB)'
            $Report.RootCauseAnalysis += "POSSIBLE CAUSE: Large PDF file being processed ($PDFSize MB)"
        }
    }
}

if ($Report.PrinterInfo) {
    $GenericDrivers = $Report.PrinterInfo.Printers | Where-Object { $_.DriverName -match 'Generic' }
    if ($GenericDrivers) {
        $Report.RootCauseAnalysis += "POSSIBLE CAUSE: Generic printer driver - known to cause Acrobat hangs"
    }
}

if ($Report.NetworkDrives) {
    $InaccessibleDrives = $Report.NetworkDrives | Where-Object { -not $_.Accessible }
    if ($InaccessibleDrives) {
        $Report.RootCauseAnalysis += "POSSIBLE CAUSE: Inaccessible network drive(s) causing timeout"
    }
}

if ($Report.EventLogErrors) {
    $CriticalErrors = $Report.EventLogErrors | Where-Object { $_.LevelDisplayName -eq 'Error' -and $_.Message -match 'crash|exception|fault' }
    if ($CriticalErrors) {
        $Report.RootCauseAnalysis += "CRITICAL EVENT: $($CriticalErrors[0].Message.Substring(0, [Math]::Min(150, $CriticalErrors[0].Message.Length)))"
    }
}

if ($Report.RootCauseAnalysis.Count -eq 0) {
    $Report.RootCauseAnalysis += "Unable to determine definitive root cause - may require process dump analysis"
}

foreach ($Cause in $Report.RootCauseAnalysis) {
    Write-DiagLog $Cause -Level CRITICAL
}

# Export full report
$ReportFile = Join-Path $OutputPath "DiagnosticReport_$Timestamp.json"
$Report | ConvertTo-Json -Depth 10 | Out-File -FilePath $ReportFile -Encoding UTF8

Write-DiagLog ""
Write-DiagLog "=== DIAGNOSTIC COMPLETE ===" -Level SUCCESS
Write-DiagLog "Full report saved to: $ReportFile" -Level SUCCESS
Write-DiagLog "Log file: $LogFile" -Level SUCCESS

# Recommendations
Write-DiagLog ""
Write-DiagLog "=== RECOMMENDATIONS ===" -Level SUCCESS

if ($Report.RootCauseAnalysis -match 'Network') {
    Write-DiagLog "1. Avoid opening PDFs from network drives - copy locally first" -Level WARNING
    Write-DiagLog "2. Check network connectivity and timeouts" -Level WARNING
}

if ($Report.RootCauseAnalysis -match 'printer|driver') {
    Write-DiagLog "1. Update or change default printer driver" -Level WARNING
    Write-DiagLog "2. Try disabling printer in Acrobat: Edit > Preferences > General > Enable printing" -Level WARNING
}

if ($Report.RecentPDFs | Where-Object { $_.'Size(MB)' -gt 50 }) {
    Write-DiagLog "1. Large PDFs detected - consider splitting or optimizing files" -Level WARNING
    Write-DiagLog "2. Increase virtual memory/page file size" -Level WARNING
}

Write-DiagLog "3. Run Fix-AdobeAcrobat.ps1 to optimize settings" -Level WARNING
Write-DiagLog "4. Consider disabling plugins: Edit > Preferences > Security (Enhanced)" -Level WARNING

#endregion

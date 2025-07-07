#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Comprehensive Windows System Repair Script
.DESCRIPTION
    Performs automated system repairs including DISM, SFC, volume checks, and cleanup operations
.NOTES
    Author: Earl Daniels
    Version: 2.0
    Requires: Administrator privileges
#>

[CmdletBinding()]
param(
    [string]$LogDirectory = "C:\Logs",
    [switch]$SkipVolumeCheck,
    [switch]$Verbose = $true
)

# Initialize logging
$logPath = Join-Path $LogDirectory "SystemRepair_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$null = New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction SilentlyContinue

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('INFO', 'WARNING', 'ERROR', 'SUCCESS')]
        [string]$Level = 'INFO'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Output to console with color coding
    switch ($Level) {
        'ERROR'   { Write-Host $logEntry -ForegroundColor Red }
        'WARNING' { Write-Host $logEntry -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $logEntry -ForegroundColor Green }
        default   { Write-Host $logEntry }
    }
    
    # Write to log file
    try {
        $logEntry | Out-File -FilePath $logPath -Append -Encoding UTF8
    } catch {
        Write-Warning "Failed to write to log file: $_"
    }
}

function Show-Progress {
    param(
        [string]$Activity,
        [int]$DurationMinutes = 15
    )
    
    $endTime = (Get-Date).AddMinutes($DurationMinutes)
    $counter = 0
    $spinner = @('|', '/', '-', '\')
    
    while ((Get-Date) -lt $endTime) {
        $remaining = [math]::Round((($endTime - (Get-Date)).TotalMinutes), 1)
        Write-Host "`r  $($spinner[$counter % 4]) $Activity - Est. $remaining min remaining..." -NoNewline -ForegroundColor Cyan
        $counter++
        Start-Sleep -Seconds 2
    }
    Write-Host "`r" -NoNewline  # Clear the line
}

function Invoke-CommandWithLogging {
    param(
        [Parameter(Mandatory)]
        [string]$Command,
        [string]$Description,
        [switch]$ContinueOnError,
        [int]$TimeoutMinutes = 30
    )
    
    Write-Log "Executing: $Description" -Level INFO
    
    try {
        # Parse command and arguments
        $parts = $Command -split ' ', 2
        $executable = $parts[0]
        $arguments = if ($parts.Count -gt 1) { $parts[1] } else { '' }
        
        # Create process start info
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $executable
        $psi.Arguments = $arguments
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        
        # Start process
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        
        # Create string builders for output
        $outputBuilder = New-Object System.Text.StringBuilder
        $errorBuilder = New-Object System.Text.StringBuilder
        
        # Event handlers for real-time output
        $outputHandler = {
            if ($EventArgs.Data) {
                $outputBuilder.AppendLine($EventArgs.Data)
                Write-Host "  $($EventArgs.Data)" -ForegroundColor Gray
            }
        }
        
        $errorHandler = {
            if ($EventArgs.Data) {
                $errorBuilder.AppendLine($EventArgs.Data)
                Write-Host "  ERROR: $($EventArgs.Data)" -ForegroundColor Red
            }
        }
        
        # Register event handlers
        Register-ObjectEvent -InputObject $process -EventName OutputDataReceived -Action $outputHandler | Out-Null
        Register-ObjectEvent -InputObject $process -EventName ErrorDataReceived -Action $errorHandler | Out-Null
        
        # Start the process
        $process.Start() | Out-Null
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
        
        # Wait for completion with timeout
        $timeoutMs = $TimeoutMinutes * 60 * 1000
        $completed = $process.WaitForExit($timeoutMs)
        
        # Clean up event handlers
        Get-EventSubscriber | Where-Object { $_.SourceObject -eq $process } | Unregister-Event
        
        if (-not $completed) {
            Write-Log "WARNING: $Description timed out after $TimeoutMinutes minutes" -Level WARNING
            try {
                $process.Kill()
                $process.WaitForExit(5000)
            } catch {
                Write-Log "Failed to terminate process: $_" -Level ERROR
            }
            return $false
        }
        
        # Get final output
        $output = $outputBuilder.ToString()
        $errorOutput = $errorBuilder.ToString()
        $exitCode = $process.ExitCode
        
        # Process results
        if ($exitCode -eq 0) {
            Write-Log "SUCCESS: $Description completed successfully (Exit Code: $exitCode)" -Level SUCCESS
            if ($Verbose -and $output) {
                Write-Log "Output: $output" -Level INFO
            }
            return $true
        } else {
            Write-Log "ERROR: $Description failed with exit code $exitCode" -Level ERROR
            if ($output) { Write-Log "Output: $output" -Level ERROR }
            if ($errorOutput) { Write-Log "Error Output: $errorOutput" -Level ERROR }
            
            if (-not $ContinueOnError) {
                throw "Command failed: $Command (Exit Code: $exitCode)"
            }
            return $false
        }
        
    } catch {
        Write-Log "EXCEPTION: $Description failed with error: $_" -Level ERROR
        if (-not $ContinueOnError) {
            throw
        }
        return $false
    } finally {
        # Cleanup
        if ($process) {
            try {
                if (-not $process.HasExited) {
                    $process.Kill()
                }
                $process.Dispose()
            } catch {
                # Ignore cleanup errors
            }
        }
    }
}

function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-SystemInfo {
    Write-Log "Gathering system information..." -Level INFO
    
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
    $computerInfo = Get-CimInstance -ClassName Win32_ComputerSystem
    
    Write-Log "OS: $($osInfo.Caption) Build $($osInfo.BuildNumber)" -Level INFO
    Write-Log "Computer: $($computerInfo.Name) ($($computerInfo.Manufacturer) $($computerInfo.Model))" -Level INFO
    Write-Log "Total RAM: $([math]::Round($computerInfo.TotalPhysicalMemory / 1GB, 2)) GB" -Level INFO
}

# Main execution
try {
    Write-Log "=== Windows System Repair Script Started ===" -Level INFO
    
    # Verify administrator privileges
    if (-not (Test-AdminPrivileges)) {
        Write-Log "ERROR: This script requires administrator privileges" -Level ERROR
        throw "Administrator privileges required"
    }
    
    # Gather system information
    Get-SystemInfo
    
    # Initialize repair status tracking
    $repairResults = @{}
    
    # 1. DISM Health Scan
    Write-Log "Phase 1: DISM Component Store Health Check" -Level INFO
    $repairResults['DISMScan'] = Invoke-CommandWithLogging -Command "DISM /Online /Cleanup-Image /ScanHealth /LogPath:`"$LogDirectory\DISM_Scan.log`"" -Description "DISM Health Scan" -ContinueOnError
    
    # 2. DISM Restore Health (only if scan found issues)
    if ($repairResults['DISMScan']) {
        Write-Log "Phase 2: DISM Component Store Repair" -Level INFO
        $repairResults['DISMRestore'] = Invoke-CommandWithLogging -Command "DISM /Online /Cleanup-Image /RestoreHealth /LogPath:`"$LogDirectory\DISM_Restore.log`"" -Description "DISM Restore Health" -ContinueOnError
    }
    
    # 3. System File Checker
    Write-Log "Phase 3: System File Checker" -Level INFO
    $repairResults['SFC'] = Invoke-CommandWithLogging -Command "sfc /scannow" -Description "System File Checker" -ContinueOnError
    
    # 4. Volume Health Check (if not skipped)
    if (-not $SkipVolumeCheck) {
        Write-Log "Phase 4: Volume Health Check" -Level INFO
        try {
            $volumes = Get-Volume | Where-Object { $_.DriveLetter -and $_.FileSystemType -eq 'NTFS' }
            foreach ($volume in $volumes) {
                Write-Log "Checking volume $($volume.DriveLetter):" -Level INFO
                $checkResult = Repair-Volume -DriveLetter $volume.DriveLetter -OfflineScanAndFix -ErrorAction SilentlyContinue
                if ($checkResult) {
                    Write-Log "Volume $($volume.DriveLetter): repair initiated" -Level SUCCESS
                } else {
                    Write-Log "Volume $($volume.DriveLetter): no issues found or repair not needed" -Level INFO
                }
            }
            $repairResults['VolumeCheck'] = $true
        } catch {
            Write-Log "Volume health check failed: $_" -Level ERROR
            $repairResults['VolumeCheck'] = $false
        }
    }
    
    # 5. Component Store Cleanup
    Write-Log "Phase 5: Component Store Cleanup" -Level INFO
    $repairResults['Cleanup'] = Invoke-CommandWithLogging -Command "DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase /LogPath:`"$LogDirectory\DISM_Cleanup.log`"" -Description "Component Store Cleanup" -ContinueOnError
    
    # 6. Windows Update Database Reset (if SFC failed)
    if (-not $repairResults['SFC']) {
        Write-Log "Phase 6: Windows Update Database Reset" -Level INFO
        try {
            Stop-Service -Name wuauserv, cryptsvc, bits, msiserver -Force -ErrorAction SilentlyContinue
            
            $folders = @("C:\Windows\SoftwareDistribution", "C:\Windows\System32\catroot2")
            foreach ($folder in $folders) {
                if (Test-Path $folder) {
                    $backupFolder = "$folder.old"
                    if (Test-Path $backupFolder) { Remove-Item $backupFolder -Recurse -Force }
                    Rename-Item $folder $backupFolder -ErrorAction SilentlyContinue
                    Write-Log "Renamed $folder to $backupFolder" -Level INFO
                }
            }
            
            Start-Service -Name wuauserv, cryptsvc, bits, msiserver -ErrorAction SilentlyContinue
            $repairResults['WUDatabaseReset'] = $true
            Write-Log "Windows Update database reset completed" -Level SUCCESS
        } catch {
            Write-Log "Windows Update database reset failed: $_" -Level ERROR
            $repairResults['WUDatabaseReset'] = $false
        }
    }
    
    # Generate summary report
    Write-Log "=== REPAIR SUMMARY ===" -Level INFO
    $successCount = 0
    $totalCount = 0
    
    foreach ($operation in $repairResults.GetEnumerator()) {
        $totalCount++
        $status = if ($operation.Value) { "SUCCESS"; $successCount++ } else { "FAILED" }
        $level = if ($operation.Value) { "SUCCESS" } else { "ERROR" }
        Write-Log "$($operation.Key): $status" -Level $level
    }
    
    Write-Log "Overall Success Rate: $successCount/$totalCount operations completed successfully" -Level INFO
    
    # Recommend reboot if any repairs were performed
    if ($repairResults.Values -contains $true) {
        Write-Log "RECOMMENDATION: Restart the computer to complete repairs" -Level WARNING
    }
    
    Write-Log "=== Windows System Repair Script Completed ===" -Level SUCCESS
    Write-Log "Log file saved to: $logPath" -Level INFO
    
} catch {
    Write-Log "CRITICAL ERROR: Script execution failed: $_" -Level ERROR
    Write-Log "Log file saved to: $logPath" -Level INFO
    exit 1
}

# Optional: Prompt for reboot
$rebootChoice = Read-Host "Would you like to restart the computer now? (y/N)"
if ($rebootChoice -match '^[Yy]') {
    Write-Log "Initiating system restart..." -Level INFO
    Restart-Computer -Force
}

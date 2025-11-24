#Requires -Version 5.1
<#
.SYNOPSIS
    Diagnoses and fixes Adobe Acrobat not responding issues.

.DESCRIPTION
    Automated script to force close hanging Acrobat processes, clear cache,
    repair installation, and optimize settings for better performance.

.PARAMETER Action
    Specify action: Diagnose, Fix, or Full (default)

.EXAMPLE
    .\Fix-AdobeAcrobat.ps1
    .\Fix-AdobeAcrobat.ps1 -Action Diagnose

.NOTES
    Author: Earl's Automation Suite
    Version: 1.0
#>

[CmdletBinding()]
param(
    [ValidateSet('Diagnose', 'Fix', 'Full')]
    [string]$Action = 'Full'
)

$ErrorActionPreference = 'Continue'
$VerbosePreference = 'Continue'

# Initialize results
$Results = @{
    ProcessesKilled = @()
    CacheCleared = $false
    SettingsOptimized = $false
    Errors = @()
}

#region Functions

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $Color = switch ($Level) {
        'ERROR' { 'Red' }
        'WARNING' { 'Yellow' }
        'SUCCESS' { 'Green' }
        default { 'White' }
    }
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "[$Timestamp] [$Level] $Message" -ForegroundColor $Color
}

function Get-AdobeProcesses {
    $AdobeProcessNames = @(
        'Acrobat',
        'AcroRd32',
        'AcroCEF',
        'RdrCEF',
        'AdobeCollabSync',
        'AdobeARM',
        'Adobe Desktop Service',
        'CCXProcess',
        'CCLibrary'
    )
    
    $Processes = Get-Process | Where-Object { 
        $ProcessName = $_.ProcessName
        $AdobeProcessNames | Where-Object { $ProcessName -like "*$_*" }
    }
    
    return $Processes
}

function Stop-AdobeProcesses {
    Write-Log "Checking for Adobe processes..."
    
    $Processes = Get-AdobeProcesses
    
    if ($Processes) {
        Write-Log "Found $($Processes.Count) Adobe process(es)" -Level WARNING
        
        foreach ($Process in $Processes) {
            try {
                Write-Log "Stopping $($Process.ProcessName) (PID: $($Process.Id))..."
                Stop-Process -Id $Process.Id -Force -ErrorAction Stop
                $Results.ProcessesKilled += $Process.ProcessName
                Start-Sleep -Milliseconds 500
                Write-Log "Successfully stopped $($Process.ProcessName)" -Level SUCCESS
            }
            catch {
                Write-Log "Failed to stop $($Process.ProcessName): $($_.Exception.Message)" -Level ERROR
                $Results.Errors += "Process kill failed: $($Process.ProcessName)"
            }
        }
    }
    else {
        Write-Log "No Adobe processes currently running" -Level SUCCESS
    }
}

function Get-AdobeInstallPath {
    $PossiblePaths = @(
        "${env:ProgramFiles}\Adobe\Acrobat DC\Acrobat",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat DC\Acrobat",
        "${env:ProgramFiles}\Adobe\Acrobat Reader DC\Reader",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat Reader DC\Reader",
        "${env:ProgramFiles}\Adobe\Acrobat 2020\Acrobat",
        "${env:ProgramFiles(x86)}\Adobe\Acrobat 2020\Acrobat"
    )
    
    foreach ($Path in $PossiblePaths) {
        if (Test-Path $Path) {
            return $Path
        }
    }
    
    return $null
}

function Clear-AdobeCache {
    Write-Log "Clearing Adobe Acrobat cache..."
    
    $CachePaths = @(
        "$env:LOCALAPPDATA\Adobe\Acrobat",
        "$env:LOCALAPPDATA\Adobe\AcrobatReader",
        "$env:TEMP\Adobe*",
        "$env:APPDATA\Adobe\Acrobat",
        "$env:LOCALAPPDATA\Temp\AcroRd32*"
    )
    
    $TotalCleared = 0
    
    foreach ($Path in $CachePaths) {
        if (Test-Path $Path) {
            try {
                $Items = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue
                $Size = ($Items | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
                
                if ($Size) {
                    $SizeMB = [math]::Round($Size / 1MB, 2)
                    Write-Log "Clearing $Path ($SizeMB MB)..."
                    
                    Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
                    $TotalCleared += $SizeMB
                }
            }
            catch {
                Write-Log "Could not clear $Path : $($_.Exception.Message)" -Level WARNING
            }
        }
    }
    
    if ($TotalCleared -gt 0) {
        Write-Log "Successfully cleared $([math]::Round($TotalCleared, 2)) MB of cache" -Level SUCCESS
        $Results.CacheCleared = $true
    }
    else {
        Write-Log "No cache found to clear" -Level SUCCESS
    }
}

function Optimize-AdobeSettings {
    Write-Log "Optimizing Adobe Acrobat settings..."
    
    # Registry paths for Acrobat DC
    $RegPaths = @(
        "HKCU:\Software\Adobe\Adobe Acrobat\DC\AVGeneral",
        "HKCU:\Software\Adobe\Acrobat Reader\DC\AVGeneral",
        "HKCU:\Software\Adobe\Adobe Acrobat\DC\Privileged",
        "HKCU:\Software\Adobe\Acrobat Reader\DC\Privileged"
    )
    
    foreach ($RegPath in $RegPaths) {
        if (Test-Path $RegPath) {
            try {
                # Disable Protected Mode (common cause of hangs)
                Set-ItemProperty -Path $RegPath -Name "bProtectedMode" -Value 0 -Force -ErrorAction SilentlyContinue
                
                # Disable Protected View at Startup
                Set-ItemProperty -Path $RegPath -Name "bProtectedView" -Value 0 -Force -ErrorAction SilentlyContinue
                
                Write-Log "Optimized settings in $RegPath" -Level SUCCESS
                $Results.SettingsOptimized = $true
            }
            catch {
                Write-Log "Could not modify $RegPath : $($_.Exception.Message)" -Level WARNING
            }
        }
    }
    
    # Disable Adobe services that can cause conflicts
    $ServicesToDisable = @('AdobeARMservice')
    
    foreach ($ServiceName in $ServicesToDisable) {
        $Service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($Service -and $Service.StartType -ne 'Disabled') {
            try {
                Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
                Set-Service -Name $ServiceName -StartupType Manual -ErrorAction Stop
                Write-Log "Set $ServiceName to Manual startup" -Level SUCCESS
            }
            catch {
                Write-Log "Could not modify $ServiceName : $($_.Exception.Message)" -Level WARNING
            }
        }
    }
}

function Get-DiagnosticInfo {
    Write-Log "=== ADOBE ACROBAT DIAGNOSTIC REPORT ===" -Level SUCCESS
    
    # Check for running processes
    $Processes = Get-AdobeProcesses
    if ($Processes) {
        Write-Log "Running Adobe Processes:" -Level WARNING
        $Processes | ForEach-Object { 
            Write-Log "  - $($_.ProcessName) (PID: $($_.Id), Memory: $([math]::Round($_.WorkingSet64/1MB, 2)) MB)"
        }
    }
    else {
        Write-Log "No Adobe processes currently running" -Level SUCCESS
    }
    
    # Check installation
    $InstallPath = Get-AdobeInstallPath
    if ($InstallPath) {
        Write-Log "Installation found at: $InstallPath" -Level SUCCESS
        
        $ExePath = Join-Path $InstallPath "Acrobat.exe"
        if (-not (Test-Path $ExePath)) {
            $ExePath = Join-Path $InstallPath "AcroRd32.exe"
        }
        
        if (Test-Path $ExePath) {
            $Version = (Get-Item $ExePath).VersionInfo.FileVersion
            Write-Log "Version: $Version" -Level SUCCESS
        }
    }
    else {
        Write-Log "Adobe Acrobat/Reader installation not found" -Level ERROR
    }
    
    # Check cache size
    $CachePaths = @(
        "$env:LOCALAPPDATA\Adobe\Acrobat",
        "$env:LOCALAPPDATA\Adobe\AcrobatReader"
    )
    
    $TotalCacheSize = 0
    foreach ($Path in $CachePaths) {
        if (Test-Path $Path) {
            $Size = (Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue | 
                     Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            if ($Size) {
                $TotalCacheSize += $Size
            }
        }
    }
    
    if ($TotalCacheSize -gt 0) {
        $CacheSizeMB = [math]::Round($TotalCacheSize / 1MB, 2)
        Write-Log "Total cache size: $CacheSizeMB MB" -Level WARNING
    }
    
    # Check system resources
    $OS = Get-CimInstance Win32_OperatingSystem
    $FreeMemoryGB = [math]::Round($OS.FreePhysicalMemory / 1MB, 2)
    Write-Log "Available RAM: $FreeMemoryGB GB" -Level SUCCESS
    
    Write-Log "=== END DIAGNOSTIC REPORT ===" -Level SUCCESS
}

function Repair-AdobeInstallation {
    Write-Log "Attempting to repair Adobe installation..."
    
    $InstallPath = Get-AdobeInstallPath
    if (-not $InstallPath) {
        Write-Log "Cannot repair: Installation path not found" -Level ERROR
        return
    }
    
    # Check for repair executable
    $ParentPath = Split-Path $InstallPath -Parent
    $SetupPath = Get-ChildItem -Path $ParentPath -Filter "*Setup*.exe" -Recurse -ErrorAction SilentlyContinue | 
                 Select-Object -First 1
    
    if ($SetupPath) {
        Write-Log "Found setup at: $($SetupPath.FullName)"
        Write-Log "To repair installation, run:" -Level WARNING
        Write-Log "  Start-Process -FilePath '$($SetupPath.FullName)' -ArgumentList '/sAll /rs /msi REINSTALL=ALL REINSTALLMODE=vomus'" -Level WARNING
    }
    else {
        Write-Log "Manual repair required via Control Panel > Programs > Adobe Acrobat > Repair" -Level WARNING
    }
}

#endregion

#region Main Execution

Write-Log "=== Adobe Acrobat Fix & Diagnostic Tool ===" -Level SUCCESS
Write-Log "Action: $Action"
Write-Log ""

switch ($Action) {
    'Diagnose' {
        Get-DiagnosticInfo
    }
    'Fix' {
        Stop-AdobeProcesses
        Clear-AdobeCache
        Optimize-AdobeSettings
    }
    'Full' {
        Get-DiagnosticInfo
        Write-Log ""
        Stop-AdobeProcesses
        Clear-AdobeCache
        Optimize-AdobeSettings
        Repair-AdobeInstallation
    }
}

Write-Log ""
Write-Log "=== SUMMARY ===" -Level SUCCESS

if ($Results.ProcessesKilled.Count -gt 0) {
    Write-Log "Processes stopped: $($Results.ProcessesKilled -join ', ')" -Level SUCCESS
}

if ($Results.CacheCleared) {
    Write-Log "Cache cleared successfully" -Level SUCCESS
}

if ($Results.SettingsOptimized) {
    Write-Log "Settings optimized" -Level SUCCESS
}

if ($Results.Errors.Count -gt 0) {
    Write-Log "Errors encountered: $($Results.Errors.Count)" -Level WARNING
    $Results.Errors | ForEach-Object { Write-Log "  - $_" -Level WARNING }
}

Write-Log ""
Write-Log "Adobe Acrobat should now be stable. Try reopening the application." -Level SUCCESS

#endregion

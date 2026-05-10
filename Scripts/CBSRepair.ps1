<#
.SYNOPSIS
    Automated CBS Repair Script for Windows Server 2016
.DESCRIPTION
    Fixes CBS errors by stopping Windows Update related services,
    renaming corrupted pending.xml, running DISM and SFC,
    and logging output.
    Ensures elevation and prevents Access Denied errors.
#>

# =============================
# Elevation Check (Auto-Restart as Admin)
# =============================
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Output "Script not running as Administrator. Relaunching elevated..."
    Start-Process powershell "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    Exit
}

# =============================
# Setup Logging
# =============================
$LogFile = "C:\Windows\Temp\CBS_Repair.log"
Function Write-Log {
    param([string]$Message)
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "$Time - $Message"
    Write-Output $Entry
    Add-Content -Path $LogFile -Value $Entry
}
Write-Log "= Starting CBS Repair Script ="

# =============================
# Step 0: Stop Windows Update and Dependencies
# =============================
$ServicesToStop = @("wuauserv","cryptsvc","bits","trustedinstaller")
foreach ($service in $ServicesToStop) {
    try {
        $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -ne 'Stopped') {
            Write-Log "Stopping service $service..."
            Stop-Service -Name $service -Force -ErrorAction Stop
            Write-Log "Service $service stopped."
        } else {
            Write-Log "Service $service already stopped or not found."
        }
    } catch {
        Write-Log "ERROR: Could not stop service $service - $_"
    }
}

# =============================
# Step 1: System Resource Check
# =============================
$disk = Get-PSDrive C
$mem = Get-CimInstance Win32_OperatingSystem
Write-Log "Disk Free Space on C: $([math]::Round($disk.Free/1GB,2)) GB"
Write-Log "Total Memory: $([math]::Round($mem.TotalVisibleMemorySize/1MB,2)) GB"
Write-Log "Free Memory : $([math]::Round($mem.FreePhysicalMemory/1MB,2)) GB"
if ($disk.Free -lt 5GB) {
    Write-Log "WARNING: Less than 5GB free on C: drive. Cleanup may be required."
}

# =============================
# Step 2: Handle pending.xml Files
# =============================
$PendingFiles = Get-ChildItem "C:\Windows\WinSxS\" -Filter "pending.xml*" -ErrorAction SilentlyContinue
if ($PendingFiles) {
    foreach ($file in $PendingFiles) {
        try {
            $newName = "$($file.FullName).$((Get-Date).ToString('yyyyMMddHHmmss')).old"
            Write-Log "Taking ownership of $($file.FullName)"
            & takeown.exe /f $file.FullName /A /R /D Y | Out-Null

            Write-Log "Resetting permissions for $($file.FullName)"
            & icacls.exe $file.FullName /reset /T /C | Out-Null

            Write-Log "Granting Administrators full access to $($file.FullName)"
            & icacls.exe $file.FullName /grant Administrators:F /T /C | Out-Null

            Write-Log "Renaming $($file.FullName) to $newName"
            Rename-Item -Path $file.FullName -NewName $newName -Force

            Write-Log "Successfully renamed $($file.Name) to $newName"
        } catch {
            Write-Log "ERROR: Failed to rename $($file.FullName) - $_"
        }
    }
} else {
    Write-Log "No pending.xml* files found. Skipping."
}

# =============================
# Step 3: Run DISM Commands
# =============================
$DismCmds = @(
    @("/Online","/Cleanup-Image","/CheckHealth"),
    @("/Online","/Cleanup-Image","/ScanHealth"),
    @("/Online","/Cleanup-Image","/RestoreHealth")
)
foreach ($args in $DismCmds) {
    try {
        Write-Log "Running: DISM $($args -join ' ')"
        & dism.exe @args 2>&1 | Tee-Object -FilePath $LogFile -Append
    } catch {
        Write-Log "ERROR: DISM failed with: $_"
    }
}

# =============================
# Step 4: Run SFC
# =============================
try {
    Write-Log "Running: sfc /scannow"
    & sfc.exe /scannow 2>&1 | Tee-Object -FilePath $LogFile -Append
} catch {
    Write-Log "ERROR: SFC failed with: $_"
}

# =============================
# Complete
# =============================
Write-Log "= CBS Repair Script Completed ="
Write-Output "CBS Repair finished. Log file saved at: $LogFile"

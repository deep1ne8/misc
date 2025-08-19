<#
.SYNOPSIS
    Automated CBS Repair Script for Windows Server 2016
.DESCRIPTION
    Handles common CBS errors:
    - 0x800705aa (ERROR_NO_SYSTEM_RESOURCES)
    - 0x80070002 (ERROR_FILE_NOT_FOUND)
    Renames corrupted pending.xml files, runs DISM and SFC, and logs output.
#>

$LogFile = "C:\Windows\Temp\CBS_Repair.log"

Function Write-Log {
    param([string]$Message)
    $Time = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Entry = "$Time - $Message"
    Write-Output $Entry
    Add-Content -Path $LogFile -Value $Entry
}

Write-Log "===== Starting CBS Repair Script ====="

# Step 1: Check system resources
$disk = Get-PSDrive C
$mem = Get-CimInstance Win32_OperatingSystem

Write-Log "Disk Free Space on C: $([math]::Round($disk.Free/1GB,2)) GB"
Write-Log "Total Memory: $([math]::Round($mem.TotalVisibleMemorySize/1MB,2)) GB"
Write-Log "Free Memory : $([math]::Round($mem.FreePhysicalMemory/1MB,2)) GB"

if ($disk.Free -lt 5GB) {
    Write-Log "WARNING: Less than 5GB free on C: drive. Cleanup may be required."
}

# Step 2: Rename pending.xml if exists
$PendingPath = "C:\Windows\WinSxS\pending.xml"
$PendingBad = "C:\Windows\WinSxS\pending.xml.bad"

foreach ($file in @($PendingPath, $PendingBad)) {
    if (Test-Path $file) {
        try {
            $newName = "$file.$((Get-Date).ToString('yyyyMMddHHmmss')).old"
            Takeown /f $file | Out-Null
            Icacls $file /grant Administrators:F | Out-Null
            Rename-Item -Path $file -NewName $newName -Force
            Write-Log "Renamed $file to $newName"
        } catch {
            Write-Log "ERROR: Failed to rename $file - $_"
        }
    } else {
        Write-Log "$file not found. Skipping."
    }
}

# Step 3: Run DISM
$DismCmds = @(
    "/Online /Cleanup-Image /CheckHealth",
    "/Online /Cleanup-Image /ScanHealth",
    "/Online /Cleanup-Image /RestoreHealth"
)

foreach ($cmd in $DismCmds) {
    Write-Log "Running: DISM $cmd"
    Start-Process -FilePath "dism.exe" -ArgumentList $cmd -Wait -NoNewWindow -RedirectStandardOutput $LogFile -Append
}

# Step 4: Run SFC
Write-Log "Running: sfc /scannow"
Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -Wait -NoNewWindow -RedirectStandardOutput $LogFile -Append

Write-Log "===== CBS Repair Script Completed ====="

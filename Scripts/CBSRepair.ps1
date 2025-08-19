<#
.SYNOPSIS
    Automated CBS Repair Script for Windows Server 2016
.DESCRIPTION
    Fixes CBS errors by renaming corrupted pending.xml, running DISM and SFC,
    and logging output. All runs in the same terminal session.
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
            takeown /f $file | Out-Null
            icacls $file /grant Administrators:F | Out-Null
            Rename-Item -Path $file -NewName $newName -Force
            Write-Log "Renamed $file to $newName"
        } catch {
            Write-Log "ERROR: Failed to rename $file - $_"
        }
    } else {
        Write-Log "$file not found. Skipping."
    }
}

# Step 3: Run DISM commands inline
$DismCmds = @(
    "/Online /Cleanup-Image /CheckHealth",
    "/Online /Cleanup-Image /ScanHealth",
    "/Online /Cleanup-Image /RestoreHealth"
)

foreach ($cmd in $DismCmds) {
    Write-Log "Running: DISM $cmd"
    & dism.exe $cmd 2>&1 | Tee-Object -FilePath $LogFile -Append
}

# Step 4: Run SFC inline
Write-Log "Running: sfc /scannow"
& sfc.exe /scannow 2>&1 | Tee-Object -FilePath $LogFile -Append

Write-Log "===== CBS Repair Script Completed ====="

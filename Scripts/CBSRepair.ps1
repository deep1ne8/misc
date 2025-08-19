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

# Step 2: Rename ALL pending.xml* files if they exist (with ownership fix)
$PendingFiles = Get-ChildItem "C:\Windows\WinSxS\" -Filter "pending.xml*" -ErrorAction SilentlyContinue

if ($PendingFiles) {
    foreach ($file in $PendingFiles) {
        try {
            $newName = "$($file.FullName).$((Get-Date).ToString('yyyyMMddHHmmss')).old"

            Write-Log "Taking ownership of $($file.FullName)"
            & takeown.exe /f $file.FullName /A | Out-Null

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



# Step 3: Run DISM commands inline with correct argument splitting
$DismCmds = @(
    @("/Online","/Cleanup-Image","/CheckHealth"),
    @("/Online","/Cleanup-Image","/ScanHealth"),
    @("/Online","/Cleanup-Image","/RestoreHealth")
)

foreach ($args in $DismCmds) {
    Write-Log "Running: DISM $($args -join ' ')"
    & dism.exe @args 2>&1 | Tee-Object -FilePath $LogFile -Append
}

# Step 4: Run SFC inline
Write-Log "Running: sfc /scannow"
& sfc.exe /scannow 2>&1 | Tee-Object -FilePath $LogFile -Append

Write-Log "===== CBS Repair Script Completed ====="

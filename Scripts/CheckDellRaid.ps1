# Define log file
$logFile = "C:\RAID_Troubleshoot_Log.txt"

# Function to log output
function Log-Output {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Write-Output $logEntry
    $logEntry | Out-File -Append -FilePath $logFile
}

# Clear previous log
if (Test-Path $logFile) { Remove-Item $logFile -Force }

# Check if Dell OpenManage is installed (for iDRAC and RAID control)
$omreport = Get-Command omreport -ErrorAction SilentlyContinue
if ($omreport) {
    Log-Output "Dell OpenManage detected. Checking RAID Status..."
    
    # Get RAID Controller Information
    $controllers = omreport storage controller | Select-String "ID:"
    foreach ($controller in $controllers) {
        $controllerID = if ($controller -match "ID:\s+(\d+)") { $matches[1] } else { "Unknown" }
        Log-Output "RAID Controller ID: $controllerID"
        
        # Check Virtual Disks (RAID arrays)
        $virtualDisks = omreport storage vdisk controller=$controllerID
        Log-Output "Virtual Disk Status: `n$virtualDisks"
        
        # Check Physical Disks
        $physicalDisks = omreport storage pdisk controller=$controllerID
        Log-Output "Physical Disk Status: `n$physicalDisks"
    }
} else {
    Log-Output "Dell OpenManage not found. Checking with PowerShell Cmdlets..."

    # Check Virtual Disks
    $virtualDisks = Get-VirtualDisk
    foreach ($vdisk in $virtualDisks) {
        Log-Output "Virtual Disk: $($vdisk.FriendlyName)"
        Log-Output "RAID Level: $($vdisk.ResiliencySettingName)"
        Log-Output "Health: $($vdisk.HealthStatus)"
        Log-Output "Operational Status: $($vdisk.OperationalStatus)"
        Log-Output "-----------------------------------"
    }

    # Check Physical Disks
    $physicalDisks = Get-PhysicalDisk
    foreach ($pdisk in $physicalDisks) {
        Log-Output "Physical Disk: $($pdisk.DeviceId) - $($pdisk.Model)"
        Log-Output "Health: $($pdisk.HealthStatus)"
        Log-Output "Media Type: $($pdisk.MediaType)"
        Log-Output "Failure Predicted: $($pdisk.FailurePredicted)"
        Log-Output "-----------------------------------"
    }
}

Log-Output "RAID Troubleshooting Completed."

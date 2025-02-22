# Function to check if the system requires a restart
function Check-SystemRestart {
    # Initialize restart flag
    $needsRestart = $false

    # Check for pending updates
    if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") {
        Write-Output "System restart required: Pending updates."
        $needsRestart = $true
    }

    # Check for pending file rename operations
    $pendingFileRenamePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager"
    $pendingFileRenameValue = "PendingFileRenameOperations"
    if ((Get-ItemProperty -Path $pendingFileRenamePath -ErrorAction SilentlyContinue).$pendingFileRenameValue) {
        Write-Output "System restart required: Pending file rename operations."
        $needsRestart = $true
    }

    # Check for WMI status (Windows Update or other flags)
    $wmiRestartPending = Get-WmiObject -Namespace "Root\CCM\ClientSDK" -Class CCM_ClientUtilities -ErrorAction SilentlyContinue |
                         ForEach-Object { $_.DetermineIfRebootPending().RebootPending }
    if ($wmiRestartPending -eq $true) {
        Write-Output "System restart required: WMI indicates a pending reboot."
        $needsRestart = $true
    }

    # Final output
    if (-not $needsRestart) {
        Write-Output "No restart required."
    }

    return $needsRestart
}

# Execute the function
if (Check-SystemRestart) {
    Write-Output "System requires a restart."
} else {
    Write-Output "System does not require a restart."
}

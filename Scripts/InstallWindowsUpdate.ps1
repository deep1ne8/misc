function Write-Log {
    param (
        [string]$Message
    )
    $logPath = "C:\Windows\Temp\WinUpgradelog.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "[$timestamp] VERBOSE: $Message"
    Write-Host "[$timestamp] VERBOSE: $Message" -ForegroundColor Green
}
Write-Host "`n======================================================"
Write-Log -Message "Disabling Windows Update"
Write-Host "`n"
# Remove Windows Update policies
Start-Sleep 3
Write-Host "`n"
$paths = @("HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate", "HKCU:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate")
foreach ($path in $paths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Path $path -Recurse -Force
            Write-Log -Message "Removed: $path"
        } catch {
            Write-Log -Message "Failed to remove $path"
        }
    } else {
        Write-Log -Message "Not found: $path"
    }
}
Start-Sleep 3
Write-Host "`n======================================================"
# Install PSWindowsUpdate
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
Write-Log -Message "Installing PSWindowsUpdate"
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force
    }
    Start-Sleep 1
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force
    }
    Start-Sleep 1
    Import-Module -Name PSWindowsUpdate -Force
} catch {
    Write-Log -Message "Failed to install PSWindowsUpdate"
    return
}
Start-Sleep 3
Write-Host "`n====================================================="
# Initialize Windows Update
Write-Log -Message "Cleaning up Windows update components"
try {
    Import-Module -Name PSWindowsUpdate -Force
    Reset-WUComponents
    Start-Sleep 3
    Write-Log -Message "Windows Update components cleaned up"
    Start-Sleep 3
    Write-Log -Message "Suspending Bitlocker on C:\"
    try {
        $bitLockerStatus = Get-BitLockerVolume -MountPoint C: | Select-Object -ExpandProperty VolumeStatus
        if ($bitLockerStatus -eq 'FullyEncrypted') {
            Disable-BitLocker -MountPoint C:\
        } else {
            Write-Log -Message "BitLocker is not enabled on C:\"
        }
    } catch {
        Write-Log -Message "Failed to disable BitLocker"
    }
    Start-Sleep 3
    Write-Host "`n======================================================"
    # Get and install Windows updates
    Write-Log -Message "Getting and installing Windows updates"
    Start-Sleep 3
    try {
        Get-WindowsUpdate -Verbose -Install -AcceptAll
    } catch {
        Write-Log -Message "Failed to get Windows updates"
        return
    }
    Start-Sleep 3
    
} catch {
    Write-Log -Message "Failed to initialize Windows Update"
    return
}
Write-Host "`n======================================================"
# Set Windows Update options
Write-Log -Message "Setting Windows Update options"
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name AUOptions -Value 4 -Force
Write-Host "`n"
# Enable automatic updates
Write-Log -Message "Enabling Windows Auto Updates"
Write-Host "`n"
try {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\wuauserv\Parameters" -Name AutoUpdate -Value 1 -Force
} catch {
    Write-Log -Message "Failed to enable automatic updates"
    return
}
Write-Log -Message "Installation complete"
return
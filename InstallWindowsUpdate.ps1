function Write-Log {
    param (
        [string]$Message
    )
    $logPath = "C:\Windows\Temp\WinUpgradelog.txt"
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "[$timestamp] VERBOSE: $Message"
    Write-Host "[$timestamp] VERBOSE: $Message" -ForegroundColor Green
}

# Remove Windows Update policies
Write-Log -Message "Removing Windows Update policies"
Start-Sleep 3
Write-Host "`n"
$paths = @("HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate", "HKCU:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate")
foreach ($path in $paths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Path $path -Recurse -Force
            Write-Host "Removed: $path" -ForegroundColor Green
        } catch {
            Write-Host "Failed to remove $path" -ForegroundColor Red
        }
    } else {
        Write-Host "Not found: $path"  -ForegroundColor Yellow
    }
}
Start-Sleep 3
Write-Host "`n"
# Install PSWindowsUpdate
Set-ExecutionPolicy Unrestricted -Scope CurrentUser -Force
try {
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Force
    }
    Start-Sleep 1
    Write-Host "`n"
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force
    }
    Start-Sleep 1
    Write-Host "`n"
    Import-Module -Name PSWindowsUpdate -Force
} catch {
    Write-Log -Message "Failed to install PSWindowsUpdate"
    return
}
Start-Sleep 3
Write-Host "`n"
# Initialize Windows Update
Write-Log -Message "Cleaning up Windows update components"
try {
    Import-Module -Name PSWindowsUpdate -Force
    Reset-WUComponents
    Start-Sleep 3
    try {
        $bitLockerStatus = Get-BitLockerVolume -MountPoint C: | Select-Object -ExpandProperty VolumeStatus
        if ($bitLockerStatus -eq 'FullyEncrypted') {
            Disable-BitLocker -MountPoint C:\
        } else {
            Write-Log -Message "BitLocker is not enabled on C:\"
        }
    } catch {
        Write-Host "Failed to disable BitLocker"
        return
    }
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
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" -Name AUOptions -Value 4 -Force
Write-Host "`n"
# Enable automatic updates
Write-Log -Message "Enabling Windows Auto Updates"
try {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\wuauserv\Parameters" -Name AutoUpdate -Value 1 -Force
} catch {
    Write-Log -Message "Failed to enable automatic updates"
    return
}

Write-Log -Message "Installation complete"
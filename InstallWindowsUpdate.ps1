# Remove Windows Update Get-PnPSitePolicy 
$paths = @("HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate", "HKCU:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate")
foreach ($path in $paths) {
    if (Test-Path $path) {
        try {
            Remove-Item -Path $path -Recurse -Force
            Write-Output "Removed: $path"
        } catch {
            Write-Error "Failed to remove $path"
        }
    } else {
        Write-Output "Not found: $path"
    }
}

# Install PSWindowsUpdate
Set-ExecutionPolicy Unrestricted -Scope LocalMachine -Force -Confirm:$false
try {
    Install-Module -Name PSWindowsUpdate  -Force -Verbose
    Import-Module -Name PSWindowsUpdate -Force -Verbose
} catch {
    Write-Host "Failed to install PSWindowsUpdate"  -ForeGroundColor Red
    return
}

# Initialize Windows Update
Write-Host "Cleaning up Windows update components"  -ForeGroundColor Green
try {
    Reset-WUComponents
    Start-Sleep 3
    Disable-BitLocker -MountPoint C:\
    Start-Sleep 3
    Get-WindowsUpdate -Verbose -Install -AcceptAll
    Start-Sleep 3
    
} catch {
    Write-Host "Failed to initialize Windows Update"  -ForeGroundColor Red
    return
}

# Enable automatic updates
Write-Host "Enabling Windows Auto Updates" -ForeGroundColor Green
try {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\wuauserv\Parameters" -Name AutoUpdate -Value 1 -Force
} catch {
    Write-Host "Failed to enable automatic updates"  -ForeGroundColor Red
    return
}

Write-Host "Installation complete"  -ForeGroundColor Green
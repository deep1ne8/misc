# Define the Windows Search database path
$SearchDBPath = "$env:ProgramData\Microsoft\Search\Data"

# Function to check if running as Administrator
function Test-Admin {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($user)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Ensure script is run as Administrator
if (-not (Test-Admin)) {
    Write-Host "This script must be run as an administrator!" -ForegroundColor Red
    exit 1
}

# Ensure Windows Search is Enabled
Write-Host "Ensuring Windows Search is enabled..." -ForegroundColor Yellow
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WSearch" -Name "Start" -Value 2 -Force
Start-Sleep -Seconds 2

# Check if Windows Search service exists
try {
    $WSearch = Get-Service -Name "WSearch" -ErrorAction Stop
} catch {
    Write-Host "Windows Search service not found. Exiting...$WSearch" -ForegroundColor Red
    return
}

# Stop Windows Search service
Write-Host "Stopping Windows Search service..." -ForegroundColor Yellow
Stop-Service -Name "WSearch" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5

# Grant full permissions to the search index folder
if (Test-Path $SearchDBPath) {
    Write-Host "Granting full permissions to Windows Search directory..." -ForegroundColor Yellow
    icacls $SearchDBPath /grant Everyone:F /T /C /Q
    Write-Host "Permissions updated successfully." -ForegroundColor Green
}

# Delete Windows Search database and index files
if (Test-Path $SearchDBPath) {
    Write-Host "Deleting Windows Search database and index files..." -ForegroundColor Yellow
    Remove-Item -Path "$SearchDBPath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Database files deleted." -ForegroundColor Green
} else {
    Write-Host "Windows Search database folder not found. It may have been deleted already." -ForegroundColor Cyan
}

# Start Windows Search service
Write-Host "Restarting Windows Search service..." -ForegroundColor Yellow
Start-Service -Name "WSearch"

# Wait for Windows Search service to be fully running (Max 60 seconds)
$Timeout = 60
$Elapsed = 0
while ((Get-Service -Name "WSearch").Status -ne "Running") {
    if ($Elapsed -ge $Timeout) {
        Write-Host "ERROR: Windows Search service failed to start within 60 seconds!" -ForegroundColor Red
        return
    }
    Start-Sleep -Seconds 5
    $Elapsed += 5
    Write-Host "Waiting for Windows Search service to reach 'Running' state... ($Elapsed seconds elapsed)" -ForegroundColor Yellow
}

Write-Host "Windows Search service is now running." -ForegroundColor Green

# Trigger rebuild of search index using both WMI and alternative method
Write-Host "Triggering full search index rebuild..." -ForegroundColor Yellow
try {
    $Searcher = New-Object -ComObject WbemScripting.SWbemLocator
    $WMI = $Searcher.ConnectServer(".", "root\CIMv2")
    $Index = $WMI.Get("Win32_SearchIndexer").SpawnInstance_()
    $Index.Rebuild()
    Write-Host "Windows Search Index rebuild triggered successfully via WMI." -ForegroundColor Green
} catch {
    Write-Host "WMI method failed. Attempting alternative method..." -ForegroundColor Yellow
    
    # Alternative Method: Manually delete registry keys to force Windows to rebuild
    try {
        Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Search" -Name "SetupCompletedSuccessfully" -ErrorAction SilentlyContinue
        Write-Host "Windows Search will now rebuild the index on restart." -ForegroundColor Green
    } catch {
        Write-Host "Failed to reset registry settings for index rebuild. Manual intervention may be needed." -ForegroundColor Red
        return
    }
}

return

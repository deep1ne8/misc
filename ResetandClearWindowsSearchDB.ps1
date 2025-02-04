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

# Check if Windows Search service exists
try {
    $WSearch = Get-Service -Name "WSearch" -ErrorAction Stop
} catch {
    Write-Host "Windows Search service not found. Exiting...$WSearch" -ForegroundColor Red
    exit 1
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
Start-Sleep -Seconds 5

# Wait for Windows Search service to be fully running
Write-Host "Ensuring Windows Search service is running..." -ForegroundColor Yellow
while ((Get-Service -Name "WSearch").Status -ne "Running") {
    Start-Sleep -Seconds 5
}

# Trigger rebuild of search index
Write-Host "Triggering full search index rebuild..." -ForegroundColor Yellow
try {
    $Searcher = New-Object -ComObject WbemScripting.SWbemLocator
    $WMI = $Searcher.ConnectServer(".", "root\CIMv2")
    $Index = $WMI.Get("Win32_SearchIndexer").SpawnInstance_()
    $Index.Rebuild()
    Write-Host "Windows Search Index has been cleared and will be rebuilt automatically." -ForegroundColor Green
} catch {
    Write-Host "Failed to trigger Windows Search rebuild. The service might be disabled." -ForegroundColor Red
}

return

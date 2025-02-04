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
    return
}

# Stop Windows Search service
Write-Host "Stopping Windows Search service..." -ForegroundColor Yellow
Stop-Service -Name "WSearch" -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 5

# Grant full permissions to the search index folder
Write-Host "Granting full permissions to Windows Search directory..." -ForegroundColor Yellow
$acl = Get-Acl $SearchDBPath
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
$acl.SetAccessRule($rule)
Set-Acl -Path $SearchDBPath -AclObject $acl
Write-Host "Permissions updated successfully." -ForegroundColor Green

# Delete Windows Search database and index files
if (Test-Path $SearchDBPath) {
    Write-Host "Deleting Windows Search database and index files..." -ForegroundColor Yellow
    Remove-Item -Path "$SearchDBPath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Database files deleted." -ForegroundColor Green
    Start-Sleep -Seconds 5
} else {
    Write-Host "Windows Search database folder not found. It may have been deleted already." -ForegroundColor Cyan
}

# Start Windows Search service
Write-Host "Restarting Windows Search service..." -ForegroundColor Yellow
Start-Sleep -Seconds 5
Start-Service -Name "WSearch"
Start-Sleep -Seconds 5

# Wait for Windows Search service to finish rebuilding the index
Write-Host "Waiting for Windows Search service to finish rebuilding the index..." -ForegroundColor Yellow
while ((Get-Service -Name "WSearch").Status -ne "Running") {
    Start-Sleep -Seconds 5
}
# Trigger rebuild of search index
Write-Host "Triggering full search index rebuild..." -ForegroundColor Yellow
$Searcher = New-Object -ComObject WbemScripting.SWbemLocator
$WMI = $Searcher.ConnectServer(".", "root\CIMv2")
$Index = $WMI.Get("Win32_SearchIndexer").SpawnInstance_()
$Index.Rebuild()

Write-Host "Windows Search Index has been cleared and will be rebuilt automatically." -ForegroundColor Green
return

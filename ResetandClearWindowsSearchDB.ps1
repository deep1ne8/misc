function Reset-WindowsSearch {
    param(
        [string]$ServiceName = "WSearch",
        [string]$SearchDbPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
    )

    # Verbose output for starting the function
    Write-Host "Starting the reset of the Windows Search service..." -ForegroundColor Cyan
    Write-Host "Checking the current status of the $ServiceName service..." -ForegroundColor Cyan
    
    # Check if the service exists and get its status
    $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
    if ($null -eq $service) {
        Write-Host "Error: Service '$ServiceName' does not exist." -ForegroundColor Red
        return
    }

    Write-Host "Current status of '$ServiceName': $($service.Status)" -ForegroundColor Yellow

    if ($service.Status -eq 'Running') {
        Write-Host "Stopping the service..." -ForegroundColor Yellow
        Stop-Service -Name $ServiceName -Force -Verbose
    } else {
        Write-Host "Service is not running, skipping stop." -ForegroundColor Yellow
    }

    # Remove the corrupted search database
    if (Test-Path -Path $SearchDbPath) {
        Write-Host "Deleting the existing search database at '$SearchDbPath'..." -ForegroundColor Yellow
        Remove-Item -Path $SearchDbPath -Force -Verbose
    } else {
        Write-Host "Search database file '$SearchDbPath' not found. Skipping deletion." -ForegroundColor Yellow
    }

    # Start the service again
    Write-Host "Starting the $ServiceName service..." -ForegroundColor Yellow
    Start-Service -Name $ServiceName -Verbose

    # Check if the service is running after attempting to start it
    $service = Get-Service -Name $ServiceName
    if ($service.Status -eq 'Running') {
        Write-Host "The '$ServiceName' service has started successfully." -ForegroundColor Green
    } else {
        Write-Host "Error: The '$ServiceName' service could not start." -ForegroundColor Red
    }

    # Optional: Trigger a rebuild of the search index
    Write-Host "Opening Indexing Options to trigger a rebuild..." -ForegroundColor Cyan
    Start-Process "control.exe" -ArgumentList "/name Microsoft.IndexingOptions"
}

# Call the function
Reset-WindowsSearch

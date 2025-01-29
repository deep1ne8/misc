function Reset-WindowsSearch {
    param(
        [string]$ServiceName = "WSearch",
        [string]$SearchDbPath = "$env:ProgramData\Microsoft\Search\Data\Applications\Windows\Windows.edb"
    )

    # Suppress all output except errors
    $ErrorActionPreference = "Stop"

    try {
        # Check if the service exists and get its status
        $service = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($null -eq $service) { return }

        # Stop the service if it's running
        if ($service.Status -eq 'Running') {
            Stop-Service -Name $ServiceName -Force -ErrorAction SilentlyContinue
        }

        # Remove the corrupted search database
        if (Test-Path -Path $SearchDbPath) {
            Remove-Item -Path $SearchDbPath -Force -ErrorAction SilentlyContinue
        }

        # Start the service
        Start-Service -Name $ServiceName -ErrorAction SilentlyContinue

        # Trigger a rebuild of the search index quietly (no UI)
        Start-Process "control.exe" -ArgumentList "/name Microsoft.IndexingOptions" -WindowStyle Hidden
    } catch {
        # Error handling if needed (optional)
        # Write-Host "Error: $_" -ForegroundColor Red
    }
}

# Call the function
Reset-WindowsSearch

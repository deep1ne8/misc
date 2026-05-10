# PowerShell Verbose Logging Template

# Define log file path
$LogFile = "C:\Windows\Temp\ScriptLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to Write Log
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to console
    if ($Level -eq "ERROR") {
        Write-Host $LogEntry -ForegroundColor Red
    } elseif ($Level -eq "WARNING") {
        Write-Host $LogEntry -ForegroundColor Yellow
    } elseif ($Level -eq "VERBOSE") {
        Write-Verbose $Message
    } else {
        Write-Host $LogEntry
    }
    
    # Write to log file
    Add-Content -Path $LogFile -Value $LogEntry
}

# Enable Verbose Logging
$VerbosePreference = "Continue"

# Start Logging
Write-Log "Script execution started." "INFO"

try {
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
        Write-Log "This script must be run as an administrator!" "ERROR"
        exit 1
    }

    # Ensure Windows Search is Enabled
    Write-Log "Ensuring Windows Search is enabled..." "VERBOSE"
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\WSearch" -Name "Start" -Value 2 -Force
    Start-Sleep -Seconds 2

    # Check if Windows Search service exists
    try {
        $WSearch = Get-Service -Name "WSearch" -ErrorAction Stop
    } catch {
        Write-Log "Windows Search service not found. Exiting...$WSearch" "ERROR"
        return
    }

    # Stop Windows Search service
    Write-Log "Stopping Windows Search service..." "VERBOSE"
    Stop-Service -Name "WSearch" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 5

    # Grant full permissions to the search index folder
    if (Test-Path $SearchDBPath) {
        Write-Log "Granting full permissions to Windows Search directory..." "VERBOSE"
        icacls $SearchDBPath /grant Everyone:F /T /C /Q
        Write-Log "Permissions updated successfully." "INFO"
    }

    # Delete Windows Search database and index files
    if (Test-Path $SearchDBPath) {
        Write-Log "Deleting Windows Search database and index files..." "VERBOSE"
        Remove-Item -Path "$SearchDBPath\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Database files deleted." "INFO"
    } else {
        Write-Log "Windows Search database folder not found. It may have been deleted already." "WARNING"
    }

    # Start Windows Search service
    Write-Log "Restarting Windows Search service..." "VERBOSE"
    Start-Service -Name "WSearch"

    # Wait for Windows Search service to be fully running (Max 60 seconds)
    $Timeout = 60
    $Elapsed = 0
    while ((Get-Service -Name "WSearch").Status -ne "Running") {
        if ($Elapsed -ge $Timeout) {
            Write-Log "ERROR: Windows Search service failed to start within 60 seconds!" "ERROR"
            return
        }
        Start-Sleep -Seconds 5
        $Elapsed += 5
        Write-Log "Waiting for Windows Search service to reach 'Running' state... ($Elapsed seconds elapsed)" "VERBOSE"
    }

    Write-Log "Windows Search service is now running." "INFO"

    # Trigger rebuild of search index using both WMI and alternative method
    Write-Log "Triggering full search index rebuild..." "VERBOSE"
    try {
        $Searcher = New-Object -ComObject WbemScripting.SWbemLocator
        $WMI = $Searcher.ConnectServer(".", "root\CIMv2")
        $Index = $WMI.Get("Win32_SearchIndexer").SpawnInstance_()
        $Index.Rebuild()
        Write-Log "Windows Search Index rebuild triggered successfully via WMI." "INFO"
    } catch {
        Write-Log "WMI method failed. Attempting alternative method..." "WARNING"
        
        # Alternative Method: Manually delete registry keys to force Windows to rebuild
        try {
            Remove-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows Search" -Name "SetupCompletedSuccessfully" -ErrorAction SilentlyContinue
            Write-Log "Windows Search will now rebuild the index on restart." "INFO"
        } catch {
            Write-Log "Failed to reset registry settings for index rebuild. Manual intervention may be needed." "ERROR"
            return
        }
    }
} catch {
    Write-Log "An error occurred: $_" "ERROR"
} finally {
    Write-Log "Script execution completed." "INFO"
}


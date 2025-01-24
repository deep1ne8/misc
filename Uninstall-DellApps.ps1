# Ensure script runs with elevated privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script must be run as an administrator." -ForegroundColor Red
    exit
}

# Define application names related to Dell SupportAssist
$appNames = @(
    "Dell SupportAssist",
    "Dell SupportAssist OS Recovery Plugin for Dell Update",
    "Dell SupportAssist Remediation"
)

# Log file for recording removal progress
$logFile = "$env:Windows\Temp\DellSupportAssist_Uninstall.log"
if (-not(Test-Path $logFile)){New-Item -ItemType File -Path $logFile -Force}

# Function to log messages
function MessageLogger {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -Append -FilePath $logFile
    Write-Host $Message
}

MessageLogger "Starting Dell SupportAssist removal process..."

# Loop through each application name
foreach ($appName in $appNames) {
    MessageLogger "Checking for application: $appName"
    $apps = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" |
    Where-Object { $_.DisplayName -match $appName }

    if ($null -eq $apps) {
        MessageLogger "No applications found matching the name '$appName'."
        continue
    }

    foreach ($app in $apps) {
        if ($null -eq $app) {
            MessageLogger "Null application reference encountered. Skipping."
            continue
        }

        $uninstallString = $app.UninstallString
        if ($null -eq $uninstallString) {
            MessageLogger "No uninstall string found for $($app.DisplayName)."
            continue
        }

        try {
            MessageLogger "Uninstalling $($app.DisplayName)..."
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallString /quiet" -Wait -NoNewWindow
            MessageLogger "$($app.DisplayName) uninstalled successfully."
        } catch {
            MessageLogger "Failed to uninstall $($app.DisplayName). Error: $_"
            Write-Host "Exception: $_" -ForegroundColor Red
        }
    }
}

# Additional forced uninstallation (if needed)
# Remove registry entries related to SupportAssist
MessageLogger "Removing leftover registry entries..."
$regKeys = @(
    "HKLM:\SOFTWARE\Dell\SupportAssist",
    "HKLM:\SOFTWARE\WOW6432Node\Dell\SupportAssist"
)
foreach ($key in $regKeys) {
    if ($null -eq $key) {
        MessageLogger "Null registry key encountered. Skipping."
        continue
    }

    if (Test-Path $key) {
        try {
            Remove-Item -Path $key -Recurse -Force -ErrorAction Stop -Verbose
            MessageLogger "Registry key $key removed successfully."
        } catch {
            MessageLogger "Failed to remove registry key $key. Error: $_"
            Write-Host "Exception: $_" -ForegroundColor Red
        }
    } else {
        MessageLogger "Registry key $key not found."
    }
}

MessageLogger "Dell SupportAssist removal process completed."
Write-Host "Removal process completed. Check the log file at $logFile for details." -ForegroundColor Green
$LogContents = Get-Content -Path $LogFile -Tail 100
Write-Output $LogContents

# End of script execution
return


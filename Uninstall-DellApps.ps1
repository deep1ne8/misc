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
function Log-Message {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $Message" | Out-File -Append -FilePath $logFile
    Write-Host $Message
}

Log-Message "Starting Dell SupportAssist removal process..."

# Loop through each application name
foreach ($appName in $appNames) {
    Log-Message "Checking for application: $appName"
    $apps = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object {$_.DisplayName -match "$appName"})
    if ($apps) {
        foreach ($app in $apps) {
            Log-Message "Uninstalling $($app.Name)..."
            try {
                $app.UninstallString()
                Log-Message "$($app.Name) uninstalled successfully."
            } catch {
                Log-Message "Failed to uninstall $($app.Name). Error: $_"
            }
        }
    } else {
        Log-Message "No installed applications found matching $appName."
    }
}

# Additional forced uninstallation (if needed)
# Remove registry entries related to SupportAssist
Log-Message "Removing leftover registry entries..."
$regKeys = @(
    "HKLM:\SOFTWARE\Dell\SupportAssist",
    "HKLM:\SOFTWARE\WOW6432Node\Dell\SupportAssist"
)
foreach ($key in $regKeys) {
    if (Test-Path $key) {
        try {
            Remove-Item -Path $key -Recurse -Force -Verbose
            Log-Message "Registry key $key removed successfully."
        } catch {
            Log-Message "Failed to remove registry key $key. Error: $_"
        }
    } else {
        Log-Message "Registry key $key not found."
    }
}

Log-Message "Dell SupportAssist removal process completed."
Write-Host "Removal process completed. Check the log file at $logFile for details." -ForegroundColor Green

Get-Content -Path $logFile -Wait

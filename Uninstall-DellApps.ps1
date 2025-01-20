# Define the log file path
$logFile = "C:\Temp\DellSoftware_Uninstall.log"

# Function to log messages
function Log-Message {
    param (
        [string]$message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
}

# List of Dell applications to check and uninstall
$appNames = @(
    "Dell SupportAssist",
    "Dell OS Recovery Plugin for Dell Update",
    "Dell SupportAssist Remediation"
)

# Function to uninstall an application using CIM
function Uninstall-UsingCim {
    param (
        [string]$appName
    )
    $installedApp = Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -eq $appName }
    if ($installedApp) {
        Log-Message "$appName is installed (CIM). Attempting to uninstall."
        $uninstallResult = $installedApp.Uninstall()
        if ($uninstallResult.ReturnValue -eq 0) {
            Log-Message "$appName uninstalled successfully using CIM."
        } else {
            Log-Message "Failed to uninstall $appName using CIM. Return value: $($uninstallResult.ReturnValue)"
        }
    } else {
        Log-Message "$appName is not installed (CIM)."
    }
}

# Function to uninstall an application using WMI
function Uninstall-UsingWmi {
    param (
        [string]$appName
    )
    $installedApp = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -eq $appName }
    if ($installedApp) {
        Log-Message "$appName is installed (WMI). Attempting to uninstall."
        $uninstallResult = $installedApp.Uninstall()
        if ($uninstallResult.ReturnValue -eq 0) {
            Log-Message "$appName uninstalled successfully using WMI."
        } else {
            Log-Message "Failed to uninstall $appName using WMI. Return value: $($uninstallResult.ReturnValue)"
        }
    } else {
        Log-Message "$appName is not installed (WMI)."
    }
}

# Iterate through each application and attempt uninstallation
foreach ($appName in $appNames) {
    Log-Message "Checking for $appName..."
    Uninstall-UsingCim -appName $appName
    Uninstall-UsingWmi -appName $appName
}

Log-Message "Script execution completed."

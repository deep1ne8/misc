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
    "Dell OS Recovery Plugin",
    "Dell SupportAssist Remediation"
)

# Function to uninstall applications using registry uninstall strings
function Uninstall-UsingRegistry {
    param (
        [string]$appName
    )
    $uninstallKeyPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $appFound = $false

    foreach ($path in $uninstallKeyPaths) {
        $uninstallEntries = Get-ChildItem -Path $path | ForEach-Object {
            Get-ItemProperty -Path $_.PSPath
        }

        foreach ($entry in $uninstallEntries) {
            if ($entry.DisplayName -match $appName) {
                $appFound = $true
                Log-Message "$appName found in registry. Attempting to uninstall."
                
                if ($entry.UninstallString) {
                    Start-Process -FilePath "cmd.exe" -ArgumentList "/c $($entry.UninstallString)" -Wait -NoNewWindow
                    Log-Message "$appName uninstalled successfully using registry."
                } else {
                    Log-Message "Uninstall string not found for $appName in registry."
                }
                break
            }
        }
    }

    if (-not $appFound) {
        Log-Message "$appName not found in registry."
    }
}

# Function to uninstall applications using WMI
function Uninstall-UsingWmi {
    param (
        [string]$appName
    )
    $installedApp = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match $appName }
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
    Uninstall-UsingWmi -appName $appName
    Uninstall-UsingRegistry -appName $appName
}

Log-Message "Script execution completed."
Get-Content -Path $LogFile -Tail 100 -Wait

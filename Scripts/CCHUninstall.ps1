# Define application details
$appName = "CCH Axcess Engagement Workpaper Monitor"
$appVersion = "1.1.0.289"
$installFilePath = "CCH.Axcess.Engagement.ClientApp.application"

# Function to uninstall the application
function Uninstall-App {
    Write-Host "Attempting to uninstall $appName version $appVersion..."
    $app = Get-WmiObject -Class Win32_Product | Where-Object {
        $_.Name -eq $appName -and $_.Version -eq $appVersion
    }

    if ($app) {
        $app.Uninstall() | Out-Null
        Write-Host "$appName version $appVersion uninstalled successfully."
    } else {
        Write-Host "$appName version $appVersion not found or already uninstalled."
    }
}

# Function to reinstall the application
function Reinstall-App {
    Write-Host "Installing $appName using $installFilePath..."
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$installFilePath`" /quiet /norestart" -Wait -NoNewWindow
    if ($?) {
        Write-Host "$appName installed successfully."
    } else {
        Write-Host "Failed to install $appName. Please check the installation file and permissions."
    }
}

# Execution starts here
Uninstall-App
#Reinstall-App

# List of applications to uninstall
$appsToRemove = @(
    "Dell SupportAssist",
    "Dell SupportAssist OS Recovery Plugin for Dell Update",
    "Dell SupportAssist Remediation"
)

# Loop through each application
foreach ($app in $appsToRemove) {
    # Get the application using its display name
    $installedApp = Get-CimInstance -ClassName Win32_Product | Where-Object { $_.Name -eq $app }
    
    if ($installedApp) {
        Write-Host "Uninstalling $($installedApp.Name)..." -ForegroundColor Yellow
        try {
            $installedApp.Uninstall() | Out-Null
            Write-Host "$($installedApp.Name) uninstalled successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to uninstall $($installedApp.Name): $_" -ForegroundColor Red
        }
    } else {
        Write-Host "$app is not installed on this system." -ForegroundColor Cyan
    }
}

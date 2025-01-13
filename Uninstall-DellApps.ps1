# List of applications to uninstall
$appsToRemove = @(
    "Dell SupportAssist",
    "Dell SupportAssist OS Recovery Plugin for Dell Update",
    "Dell SupportAssist Remediation"
)

# Registry paths for installed applications
$registryPaths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

foreach ($app in $appsToRemove) {
    $appFound = $false
    foreach ($path in $registryPaths) {
        # Search for the application in the registry
        $uninstallKey = Get-ChildItem -Path $path | ForEach-Object {
            Get-ItemProperty -Path $_.PSPath | Where-Object { $_.DisplayName -eq $app }
        }

        if ($uninstallKey) {
            $appFound = $true
            Write-Host "Uninstalling $($uninstallKey.DisplayName)..." -ForegroundColor Yellow
            try {
                Start-Process -FilePath $uninstallKey.UninstallString -ArgumentList "/quiet" -Wait
                Write-Host "$($uninstallKey.DisplayName) uninstalled successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to uninstall $($uninstallKey.DisplayName): $_" -ForegroundColor Red
            }
        }
    }

    if (-not $appFound) {
        Write-Host "$app not found in the registry." -ForegroundColor Cyan
    }
}


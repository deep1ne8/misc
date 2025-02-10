function Uninstall-DellBloatware {
    $appNames = @(
        "Dell SupportAssist",
        "Dell SupportAssist OS Recovery Plugin for Dell Update",
        "Dell SupportAssist Remediation",
        "Dell Optimizer",
        "Dell Display Manager",
        "Dell Peripheral Manager",
        "Dell Pair",
        "Dell Core Services",
        "Dell Trusted Device"
    )

    Write-Host "Uninstalling Dell bloatware..." -ForegroundColor Yellow
    foreach ($appName in $appNames) {
        $app = Get-Package -Name $appName -ErrorAction SilentlyContinue
        if ($app) {
            Write-Host "Uninstalling $appName..." -ForegroundColor Cyan
            $app | Uninstall-Package -Confirm:$false
        } else {
            Write-Host "$appName not found." -ForegroundColor Magenta
        }
    }
}
function Uninstall-OfficeLanguagePacks {
    param (
        [switch]$Force
    )

    Write-Host "`nStarting uninstallation of Microsoft Office Language Packs and OneNote..." -ForegroundColor Yellow

    # Retrieve installed applications from the registry
    $installed = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue

    if (!$installed) {
        Write-Host "No Microsoft Office or OneNote applications found." -ForegroundColor Cyan
        return
    } else {
        Write-Host "Looking for Microsoft Office/OneNote applications to uninstall..." -ForegroundColor Cyan

        # Filter applications based on DisplayName
        $toUninstall = $installed | Where-Object { 
            ($_."DisplayName" -like "Microsoft 365*" -and $_."DisplayName" -notlike "Microsoft 365 - en-us") -or
            ($_."DisplayName" -like "Microsoft OneNote*" -and $_."DisplayName" -notlike "Microsoft OneNote - en-us")
        }

        if ($toUninstall) {
            Write-Host "`nThe following Microsoft Office/OneNote applications will be uninstalled:" -ForegroundColor Cyan
            $toUninstall | ForEach-Object { Write-Host " - $($_.DisplayName)" -ForegroundColor Magenta }

            # Uninstall applications
            foreach ($app in $toUninstall) {
                Write-Host "`nUninstalling: $($app.DisplayName)..." -ForegroundColor Red
                try {
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($app.PSChildName) /qn /norestart" -NoNewWindow -Wait -ErrorAction Stop
                    Write-Host "Successfully uninstalled: $($app.DisplayName)" -ForegroundColor Green
                } catch {
                    Write-Host "Failed to uninstall: $($app.DisplayName). Error: $_" -ForegroundColor DarkRed
                }
            }

            Write-Host "`nUninstallation process completed." -ForegroundColor Green
        } else {
            Write-Host "No matching applications found for uninstallation." -ForegroundColor Cyan
        }
    }
}

Uninstall-DellBloatware
Write-Host "`n"
Start-Sleep -Seconds 3
Uninstall-OfficeLanguagePacks
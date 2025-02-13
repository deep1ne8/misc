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
    Write-Host "`nStarting uninstallation of Microsoft Office Language Packs and OneNote..." -ForegroundColor Yellow

    # Retrieve installed applications from the registry
    $installed = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue

    if (!$installed) {
        Write-Host "No Microsoft Office or OneNote applications found." -ForegroundColor Cyan
        return
    }

    Write-Host "Looking for Microsoft Office/OneNote applications to uninstall..." -ForegroundColor Cyan

    # Filter applications based on DisplayName
    $toUninstall = $installed | Where-Object { 
        ($_."DisplayName" -like "Microsoft 365*" -and $_."DisplayName" -notlike "Microsoft 365 - en-us" -and $_."DisplayName" -notlike "Microsoft 365 - Apps for Business - en-us" -and $_."DisplayName" -notlike "Microsoft 365 - Apps for Enterprise - en-us") -or
        ($_."DisplayName" -like "Microsoft OneNote*" -and $_."DisplayName" -notlike "Microsoft OneNote - en-us")
    }

    if ($toUninstall) {
        Write-Host "Listing all Microsoft Office/OneNote applications to be uninstalled:"
        $toUninstall | ForEach-Object { Write-Host " - $($_.DisplayName)" -ForegroundColor Magenta }

        Write-Host "`n"
        Start-Sleep -Seconds 3

        $result = Read-Host "Do you want to proceed with the uninstallation? (Y/N)"
        if ($result -eq "Y") {
            foreach ($app in $toUninstall) {
                Write-Host "`nUninstalling: $($app.DisplayName)..." -ForegroundColor Red

                # Use UninstallString if available (Click-to-Run or MSI)
                if ($app.UninstallString) {
                    $uninstallCommand = $app.UninstallString

                    # Check if the UninstallString needs to be executed with cmd.exe
                    if ($uninstallCommand -match "MsiExec") {
                        $arguments = $uninstallCommand -replace "MsiExec.exe ", ""  # Strip "MsiExec.exe"
                        Start-Process -FilePath "MsiExec.exe" -ArgumentList $arguments -NoNewWindow -Wait
                    } else {
                        Start-Process -FilePath "cmd.exe" -ArgumentList "/c $uninstallCommand" -NoNewWindow -Wait
                    }
                } elseif ($app.PSChildName) {
                    # Fallback: Use MSI-based uninstallation with GUID
                    Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $($app.PSChildName) /qn /norestart" -NoNewWindow -Wait
                } else {
                    Write-Host "No uninstall method found for: $($app.DisplayName)" -ForegroundColor DarkRed
                    continue
                }

                # Confirm uninstallation
                Start-Sleep -Seconds 2  # Give time for the uninstall process
                $checkAgain = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
                              Where-Object { $_.DisplayName -eq $app.DisplayName }

                if ($checkAgain) {
                    Write-Host "❌ Failed to uninstall: $($app.DisplayName)" -ForegroundColor DarkRed
                } else {
                    Write-Host "✅ Successfully uninstalled: $($app.DisplayName)" -ForegroundColor Green
                }
            }
            Write-Host "`nUninstallation process completed." -ForegroundColor Green
        } else {
            Write-Host "Uninstallation cancelled." -ForegroundColor Yellow
            return
        }
    } else {
        Write-Host "No matching applications found for uninstallation." -ForegroundColor Cyan
    }
}
# Run the function
Uninstall-DellBloatware
Write-Host "`n"
Start-Sleep -Seconds 3
Uninstall-OfficeLanguagePacks
return
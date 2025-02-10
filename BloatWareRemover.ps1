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
Write-Host "Uninstalling Microsoft Office Language Packs/OneNote..." -ForegroundColor Yellow
$installed = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
if (!$installed) {
    Write-Host "No Microsoft Office/OneNote applications found." -ForegroundColor Cyan
    return
}else {
    Write-Host "Looking for Microsoft Office/OneNote applications to uninstall..." -ForegroundColor Cyan
    $toUninstall = $installed | Where-Object { $_.DisplayName -like "Microsoft 365*" -and $_.DisplayName -notlike "Microsoft 365 - en-us" -or $_.DisplayName -like "Microsoft OneNote*" -and $_.DisplayNAme -notlike "Microsoft OneNote - en-us" }
    Write-Host "The following Microsoft Office/OneNote applications have been found and will be uninstalled:" -ForegroundColor Cyan
    $toUninstall | ForEach-Object { Write-Host $_.DisplayName }
    $toUninstall | ForEach-Object { Start-Process -Wait "msiexec.exe" -ArgumentList @("/x", $_.PSChildName, "/qn", "/norestart") -Verbose }
    Write-Host "Uninstallation completed." -ForegroundColor Green
    }
}


Uninstall-DellBloatware
Uninstall-OfficeLanguagePacks
function Uninstall-DellBloatware {
    $appNames = @(
        "Dell SupportAssist",
        "Dell SupportAssist OS Recovery Plugin for Dell Update",
        "Dell SupportAssist Remediation"
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

@{
    Id          = 'core-services'
    Name        = 'Core Windows services'
    Kind        = 'Check'
    Category    = 'System'
    Description = 'Core services that are set to start automatically but are not running.'
    Parameters  = @{}
    Test        = {
        param($Ctx, $P)
        $names = 'Dnscache', 'Dhcp', 'EventLog', 'LanmanWorkstation', 'mpssvc', 'Winmgmt', 'Schedule', 'CryptSvc', 'Spooler', 'AudioSrv'
        $stopped = @(Get-AbService -Name $names | Where-Object { "$($_.StartType)" -eq 'Automatic' -and "$($_.Status)" -ne 'Running' })
        if ($stopped) {
            New-AbFinding -Code 'services.stopped' -Severity Warning -Title "$($stopped.Count) core service(s) not running" -Detail (($stopped | ForEach-Object { $_.DisplayName }) -join ', ') -FixTask 'start-core-services' -Data @{ Services = @($stopped | ForEach-Object { $_.Name }) }
        }
    }
}

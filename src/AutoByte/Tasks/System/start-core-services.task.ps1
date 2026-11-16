@{
    Id            = 'start-core-services'
    Name          = 'Start stopped core services'
    Kind          = 'Fix'
    Category      = 'System'
    Description   = 'Starts the core services (fixed allow-list) that are set to Automatic but stopped. Never changes startup types.'
    RequiresAdmin = $true
    Test          = {
        param($Ctx, $P)
        $names = 'Dnscache', 'Dhcp', 'EventLog', 'LanmanWorkstation', 'mpssvc', 'Winmgmt', 'Schedule', 'CryptSvc', 'Spooler', 'AudioSrv'
        $stopped = @(Get-AbService -Name $names | Where-Object { "$($_.StartType)" -eq 'Automatic' -and "$($_.Status)" -ne 'Running' })
        $Ctx.Stopped = $stopped
        if ($stopped) { @{ Needed = $true; Reason = 'Stopped: ' + (($stopped | ForEach-Object Name) -join ', ') } }
        else { @{ Needed = $false; Reason = 'All core services are running.' } }
    }
    Repair        = {
        param($Ctx, $P)
        foreach ($s in @($Ctx.Stopped)) {
            try { Set-AbServiceState -Name $s.Name -State Running; Add-AbMessage -Result $Ctx.Result -Message "Started $($s.Name)." }
            catch { Add-AbMessage -Result $Ctx.Result -Message "Could not start $($s.Name): $($_.Exception.Message)" }
        }
        @{}
    }
}

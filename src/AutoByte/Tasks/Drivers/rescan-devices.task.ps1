@{
    Id            = 'rescan-devices'
    Name          = 'Rescan hardware'
    Kind          = 'Fix'
    Category      = 'Drivers'
    Description   = 'Asks Plug and Play to re-detect devices (pnputil /scan-devices), which re-binds drivers for devices that failed to start. Changes no drivers.'
    RequiresAdmin = $true
    Test          = {
        param($Ctx, $P)
        $bad = @(Get-AbCim -ClassName Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0' | Where-Object { [int]$_.ConfigManagerErrorCode -notin 22, 45 })
        if ($bad) { return @{ Needed = $true; Reason = "$($bad.Count) device(s) report a problem: " + ((@($bad | Select-Object -First 3 | ForEach-Object Name)) -join ', ') } }
        @{ Needed = $false; Reason = 'No device reports a problem.' }
    }
    Repair        = {
        param($Ctx, $P)
        $r = Invoke-AbNative -FilePath (Get-AbSystemPath -Name PnpUtil) -ArgumentList '/scan-devices' -TimeoutSeconds 300
        if ($r.ExitCode -ne 0) { throw "pnputil /scan-devices failed ($($r.ExitCode)): $($r.Output)" }
        Add-AbMessage -Result $Ctx.Result -Message 'Hardware rescan complete. If a device still fails, run install-driver-updates.'
    }
}

@{
    Id            = 'resync-time'
    Name          = 'Resynchronise the clock'
    Kind          = 'Fix'
    Category      = 'Network'
    Description   = 'Starts the Windows Time service if needed and forces a resync when the clock is off.'
    RequiresAdmin = $true
    Test          = {
        param($Ctx, $P)
        $offset = Get-AbTimeOffset
        if ($null -eq $offset) { return @{ Needed = $false; Reason = 'Offset could not be measured (NTP blocked?); not changing anything.' } }
        if ([math]::Abs($offset) -ge [double]$Ctx.Config.TimeSkewWarnSeconds) { return @{ Needed = $true; Reason = ("Clock is off by {0:N0} s." -f $offset) } }
        @{ Needed = $false; Reason = ("Clock is within {0:N1} s." -f $offset) }
    }
    Repair        = {
        param($Ctx, $P)
        $svc = Get-AbService -Name 'W32Time'
        if ($svc -and "$($svc.Status)" -ne 'Running') { Set-AbServiceState -Name 'W32Time' -State Running }
        $w32 = Get-AbSystemPath -Name W32tm
        $null = Invoke-AbNative -FilePath $w32 -ArgumentList '/config', '/update' -TimeoutSeconds 60
        $r = Invoke-AbNative -FilePath $w32 -ArgumentList '/resync', '/force' -TimeoutSeconds 120
        if ($r.ExitCode -ne 0) { throw "w32tm /resync failed with exit code $($r.ExitCode)." }
        Add-AbMessage -Result $Ctx.Result -Message 'Clock resynchronised.'
        @{}
    }
}

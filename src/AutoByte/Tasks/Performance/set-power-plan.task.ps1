@{
    Id          = 'set-power-plan'
    Name        = 'Switch to the Balanced power plan'
    Kind        = 'Fix'
    Category    = 'Performance'
    Description = 'Activates the built-in Balanced plan when Power saver is active.'
    EndUserSafe = $true
    Test        = {
        param($Ctx, $P)
        $plan = @(Get-AbCim -Namespace 'root\cimv2\power' -ClassName Win32_PowerPlan | Where-Object { $_.IsActive })
        if ($plan -and "$($plan[0].InstanceID)" -match 'a1841308-3541-4fab-bc81-f71556f20b4a') { return @{ Needed = $true; Reason = 'Power saver is active.' } }
        @{ Needed = $false; Reason = 'Power saver is not active.' }
    }
    Repair      = {
        param($Ctx, $P)
        $r = Invoke-AbNative -FilePath (Get-AbSystemPath -Name PowerCfg) -ArgumentList '/setactive', '381b4222-f694-41f0-9685-ff5bb260df2e'
        if ($r.ExitCode -ne 0) { throw "powercfg failed ($($r.ExitCode)): $($r.Output)" }
        Add-AbMessage -Result $Ctx.Result -Message 'Balanced power plan activated.'
    }
}

@{
    Id                  = 'repair-system-files'
    Name                = 'Repair Windows system files'
    Kind                = 'Fix'
    Category            = 'System'
    Description         = 'Checks the Windows component store; if it is damaged (or -Force), runs DISM RestoreHealth then SFC. Can take 15-60 minutes.'
    RequiresAdmin       = $true
    RebootMayBeRequired = $true
    Parameters          = @{ Force = $false }
    Test                = {
        param($Ctx, $P)
        $state = Get-AbImageHealth
        $Ctx.ImageState = $state
        if ($state -eq 'NonRepairable') { throw 'The Windows image is not repairable in place; an in-place upgrade or reinstall is needed.' }
        if ($state -ne 'Healthy') { return @{ Needed = $true; Reason = "Component store is $state." } }
        if ($P.Force -and -not $Ctx.AfterRepair) { return @{ Needed = $true; Reason = 'Forced run (store reports Healthy).' } }
        @{ Needed = $false; Reason = 'Component store is Healthy.' }
    }
    Repair              = {
        param($Ctx, $P)
        $restart = $false
        if ($Ctx.ImageState -ne 'Healthy') {
            $r = Repair-AbImage
            Add-AbMessage -Result $Ctx.Result -Message "DISM RestoreHealth finished: $($r.State)."
            $restart = $r.RestartNeeded
        }
        $sfc = Invoke-AbNative -FilePath (Get-AbSystemPath -Name Sfc) -ArgumentList '/scannow' -TimeoutSeconds 3600
        Add-AbMessage -Result $Ctx.Result -Message "SFC exit code $($sfc.ExitCode)."
        if ($sfc.ExitCode -ne 0) { $restart = $true }  # SFC often finishes repairs on the next restart
        @{ RebootRequired = $restart }
    }
}

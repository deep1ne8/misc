@{
    Id            = 'update-defender-signatures'
    Name          = 'Update virus definitions'
    Kind          = 'Fix'
    Category      = 'Security'
    Description   = 'Downloads the latest Microsoft Defender definitions when they are older than the configured limit.'
    RequiresAdmin = $true
    Parameters    = @{ Force = $false }
    Test          = {
        param($Ctx, $P)
        $s = Get-AbDefenderStatus
        $age = [int]$s.AntivirusSignatureAge
        if ($age -gt [int]$Ctx.Config.DefenderSigMaxDays) { return @{ Needed = $true; Reason = "Definitions are $age days old." } }
        if ($P.Force -and -not $Ctx.AfterRepair) { return @{ Needed = $true; Reason = 'Forced update.' } }
        @{ Needed = $false; Reason = "Definitions are $age day(s) old." }
    }
    Repair        = {
        param($Ctx, $P)
        $r = Invoke-AbNative -FilePath (Get-AbSystemPath -Name MpCmdRun) -ArgumentList '-SignatureUpdate' -TimeoutSeconds 900
        if ($r.ExitCode -ne 0) { throw "MpCmdRun -SignatureUpdate failed with exit code $($r.ExitCode)." }
        Add-AbMessage -Result $Ctx.Result -Message 'Definitions updated.'
        @{}
    }
}

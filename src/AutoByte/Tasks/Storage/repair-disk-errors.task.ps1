@{
    Id                  = 'repair-disk-errors'
    Name                = 'Repair file system errors'
    Kind                = 'Fix'
    Category            = 'Storage'
    Description         = 'Scans the Windows volume online (Repair-Volume -Scan). Spot fixes are applied online; if a full repair is needed it is scheduled for the next restart (chkdsk). Does not fix a failing disk: back up and replace it.'
    RequiresAdmin       = $true
    RebootMayBeRequired = $true
    Parameters          = @{ Drive = '' }
    Test                = {
        param($Ctx, $P)
        if ($Ctx.ContainsKey('Scheduled') -and $Ctx.Scheduled) { return @{ Needed = $false; Reason = 'Repair scheduled for the next restart.' } }
        $drive = if ($P.Drive) { "$($P.Drive)".TrimEnd(':') } else { "$env:SystemDrive".TrimEnd(':') }
        if (-not $drive) { $drive = 'C' }
        $scan = Get-AbVolumeScan -DriveLetter $drive
        if ($scan -eq 'NoErrorsFound') { return @{ Needed = $false; Reason = "Drive ${drive}: no file system errors." } }
        @{ Needed = $true; Reason = "Drive ${drive}: $scan." }
    }
    Repair              = {
        param($Ctx, $P)
        $drive = if ($P.Drive) { "$($P.Drive)".TrimEnd(':') } else { "$env:SystemDrive".TrimEnd(':') }
        if (-not $drive) { $drive = 'C' }
        $scan = Get-AbVolumeScan -DriveLetter $drive
        if ($scan -eq 'SpotFixNeeded') {
            $r = Repair-AbVolume -DriveLetter $drive -Mode SpotFix
            Add-AbMessage -Result $Ctx.Result -Message "Spot fix on ${drive}: $r"
            if ($r -eq 'NoErrorsFound') { return @{ RebootRequired = $false } }
        }
        $r = Repair-AbVolume -DriveLetter $drive -Mode OfflineScanAndFix
        Add-AbMessage -Result $Ctx.Result -Message "Full repair of ${drive} scheduled for the next restart ($r)."
        $Ctx.Scheduled = $true
        @{ RebootRequired = $true }
    }
}

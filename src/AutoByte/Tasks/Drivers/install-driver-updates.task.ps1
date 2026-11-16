@{
    Id                  = 'install-driver-updates'
    Name                = 'Install driver updates'
    Kind                = 'Fix'
    Category            = 'Drivers'
    Description         = 'Installs the driver updates Windows Update (or the WSUS/Intune policy in force) offers for this PC. Only Microsoft-published, signed drivers are installed.'
    RequiresAdmin       = $true
    RebootMayBeRequired = $true
    Test                = {
        param($Ctx, $P)
        $drv = @(Get-AbPendingUpdate -Criteria "IsInstalled=0 and IsHidden=0 and Type='Driver'")
        if ($drv) { return @{ Needed = $true; Reason = "$($drv.Count) driver update(s): " + ((@($drv | Select-Object -First 3 | ForEach-Object Title)) -join '; '); Ids = @($drv.UpdateId) } }
        @{ Needed = $false; Reason = 'Windows Update offers no driver updates. For vendor-only drivers use the Dell/HP tools (abyte legacy).' }
    }
    Repair              = {
        param($Ctx, $P)
        $ids = @(Get-AbPendingUpdate -Criteria "IsInstalled=0 and IsHidden=0 and Type='Driver'" | ForEach-Object UpdateId)
        $r = Install-AbUpdate -UpdateId $ids -Criteria "IsInstalled=0 and IsHidden=0 and Type='Driver'"
        foreach ($row in $r.Results) { Add-AbMessage -Result $Ctx.Result -Message "$($row.Title): result $($row.ResultCode) ($($row.HResult))" }
        if ($r.Failed -and -not $r.Installed) { throw "All $($r.Failed) driver update(s) failed." }
        @{ RebootRequired = [bool]$r.RebootRequired }
    }
}

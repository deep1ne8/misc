@{
    Id                  = 'install-security-updates'
    Name                = 'Install security updates'
    Kind                = 'Fix'
    Category            = 'Updates'
    Description         = 'Downloads and installs pending security and critical updates through Windows Update (honours WSUS/Intune policy). This is how CVE fixes are applied. Use -All to install every pending software update. Usually needs a restart.'
    RequiresAdmin       = $true
    RebootMayBeRequired = $true
    Parameters          = @{ All = $false }
    Test                = {
        param($Ctx, $P)
        $security = '0fa1201d-4330-4fa8-8ae9-b877473b6441', 'e6cf1350-c01b-414d-a61f-263d14d133b4'
        $pending = @(Get-AbPendingUpdate | Where-Object { $P.All -or @($_.CategoryIds | Where-Object { $security -contains $_ }).Count -or $_.Severity })
        if ($pending) { return @{ Needed = $true; Reason = "$($pending.Count) update(s): " + ((@($pending | Select-Object -First 3 | ForEach-Object Title)) -join '; ') } }
        @{ Needed = $false; Reason = 'No security updates pending.' }
    }
    Repair              = {
        param($Ctx, $P)
        $security = '0fa1201d-4330-4fa8-8ae9-b877473b6441', 'e6cf1350-c01b-414d-a61f-263d14d133b4'
        $ids = @(Get-AbPendingUpdate | Where-Object { $P.All -or @($_.CategoryIds | Where-Object { $security -contains $_ }).Count -or $_.Severity } | ForEach-Object UpdateId)
        $r = Install-AbUpdate -UpdateId $ids
        foreach ($row in $r.Results) { Add-AbMessage -Result $Ctx.Result -Message "$($row.Title): result $($row.ResultCode) ($($row.HResult))" }
        if ($r.Failed) {
            Add-AbMessage -Result $Ctx.Result -Message "$($r.Failed) update(s) failed. Run 'abyte check event-logs' for the error code and the matching fix."
            if (-not $r.Installed) { throw "All $($r.Failed) update(s) failed to install." }
        }
        @{ RebootRequired = [bool]$r.RebootRequired }
    }
}

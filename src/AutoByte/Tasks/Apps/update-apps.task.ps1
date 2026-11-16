@{
    Id                  = 'update-apps'
    Name                = 'Update applications'
    Kind                = 'Fix'
    Category            = 'Apps'
    Description         = 'Upgrades applications with winget (silent, from the winget source). -Ids limits it to a comma-separated list of package ids; -Exclude skips ids. Apps that are open may be closed by their installer.'
    RequiresAdmin       = $true
    RebootMayBeRequired = $true
    Parameters          = @{ Ids = ''; Exclude = '' }
    Test                = {
        param($Ctx, $P)
        $apps = Get-AbAppUpgrade
        if ($null -eq $apps) { throw 'winget is not available on this PC.' }
        $want = @("$($P.Ids)" -split ',' | ForEach-Object Trim | Where-Object { $_ })
        $skip = @("$($P.Exclude)" -split ',' | ForEach-Object Trim | Where-Object { $_ })
        $todo = @(@($apps) | Where-Object { (-not $want -or $want -contains $_.Id) -and $skip -notcontains $_.Id })
        if ($todo) { return @{ Needed = $true; Reason = "$($todo.Count) app(s): " + ((@($todo | Select-Object -First 5 | ForEach-Object Id)) -join ', '); Ids = @($todo.Id) } }
        @{ Needed = $false; Reason = 'All selected applications are up to date.' }
    }
    Repair              = {
        param($Ctx, $P)
        $winget = Get-AbWingetPath
        $want = @("$($P.Ids)" -split ',' | ForEach-Object Trim | Where-Object { $_ })
        $skip = @("$($P.Exclude)" -split ',' | ForEach-Object Trim | Where-Object { $_ })
        $todo = @(@(Get-AbAppUpgrade) | Where-Object { (-not $want -or $want -contains $_.Id) -and $skip -notcontains $_.Id })
        $failed = 0; $reboot = $false
        foreach ($a in $todo) {
            if ($a.Id -notmatch '^[A-Za-z0-9][A-Za-z0-9._+\-]*$') { continue }   # never pass odd text to winget
            $r = Invoke-AbNative -FilePath $winget -ArgumentList 'upgrade', '--id', $a.Id, '--exact', '--silent', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity' -TimeoutSeconds 1800
            # 0 ok; 0x8A15002B (-1978335189) no applicable upgrade; 3010/1641 reboot.
            switch ($r.ExitCode) {
                { $_ -in 0, -1978335189 } { Add-AbMessage -Result $Ctx.Result -Message "Updated $($a.Id) to $($a.Available)." }
                { $_ -in 3010, 1641 } { $reboot = $true; Add-AbMessage -Result $Ctx.Result -Message "Updated $($a.Id) (restart needed)." }
                default { $failed++; Add-AbMessage -Result $Ctx.Result -Message ("{0} failed: exit 0x{1:X8}" -f $a.Id, $r.ExitCode) }
            }
        }
        if ($failed -and $failed -eq $todo.Count) { throw "All $failed app update(s) failed." }
        @{ RebootRequired = $reboot }
    }
}

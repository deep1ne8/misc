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
        if ($Ctx.AfterRepair -and $Ctx.ContainsKey('Attempted')) {
            # Apps already tried: failures are reported by Repair; successes that winget still lists are
            # a known detection quirk (Store apps, old VC++ runtimes) and usually clear after a restart.
            $still = @($todo | Where-Object { $Ctx.Succeeded -contains $_.Id })
            if ($still -and $Ctx.Result.Data -is [hashtable]) {
                $Ctx.Result.Data['StillListed'] = @($still | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Name = $_.Name } })
                Add-AbMessage -Result $Ctx.Result -Message ("winget still lists $($still.Count) updated app(s) (usually clears after a restart): " + ((@($still | ForEach-Object Name)) -join ', '))
            }
            $todo = @($todo | Where-Object { $Ctx.Attempted -notcontains $_.Id })
        }
        if ($todo) { return @{ Needed = $true; Reason = "$($todo.Count) app(s): " + ((@($todo | Select-Object -First 5 | ForEach-Object Id)) -join ', '); Ids = @($todo.Id) } }
        @{ Needed = $false; Reason = 'All selected applications are up to date.' }
    }
    Repair              = {
        param($Ctx, $P)
        $winget = Get-AbWingetPath
        $want = @("$($P.Ids)" -split ',' | ForEach-Object Trim | Where-Object { $_ })
        $skip = @("$($P.Exclude)" -split ',' | ForEach-Object Trim | Where-Object { $_ })
        $todo = @(@(Get-AbAppUpgrade) | Where-Object { (-not $want -or $want -contains $_.Id) -and $skip -notcontains $_.Id -and $_.Id -match '^[A-Za-z0-9][A-Za-z0-9._+\-]*$' })
        $updated = New-Object System.Collections.Generic.List[object]
        $failed = New-Object System.Collections.Generic.List[object]
        $Ctx.Attempted = @(); $Ctx.Succeeded = @()
        $reboot = $false
        $n = 0
        foreach ($a in $todo) {
            $n++
            Write-Progress -Id 7 -Activity 'Updating applications' -Status "$n of $($todo.Count): $($a.Name) $($a.Version) -> $($a.Available)" -PercentComplete ([int](100 * ($n - 1) / [Math]::Max(1, $todo.Count)))
            $Ctx.Attempted += $a.Id
            $r = Invoke-AbNative -FilePath $winget -ArgumentList 'upgrade', '--id', $a.Id, '--exact', '--silent', '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity' -TimeoutSeconds 1800
            $o = Get-AbWingetOutcome -ExitCode $r.ExitCode
            if ($o.Outcome -in 'Ok', 'Reboot', 'Current') {
                $Ctx.Succeeded += $a.Id
                if ($o.Outcome -eq 'Reboot') { $reboot = $true }
                $updated.Add([pscustomobject]@{ Id = $a.Id; Name = $a.Name; From = $a.Version; To = $a.Available; Note = $o.Reason })
                Add-AbMessage -Result $Ctx.Result -Message ("Updated {0} ({1} -> {2}){3}" -f $a.Name, $a.Version, $a.Available, $(if ($o.Reason) { ". $($o.Reason)" } else { '' }))
            }
            else {
                $failed.Add([pscustomobject]@{ Id = $a.Id; Name = $a.Name; From = $a.Version; To = $a.Available; Code = $o.Code; Reason = $o.Reason; Hint = $o.Hint })
                Add-AbMessage -Result $Ctx.Result -Message ("Not updated: {0} - {1} {2} [{3}]" -f $a.Name, $o.Reason, $o.Hint, $o.Code)
            }
        }
        Write-Progress -Id 7 -Activity 'Updating applications' -Completed
        @{
            RebootRequired = $reboot
            Partial        = ($failed.Count -gt 0 -and $updated.Count -gt 0)
            Failed         = $(if ($todo.Count -and $failed.Count -eq $todo.Count) { "None of the $($failed.Count) app update(s) succeeded." } else { '' })
            Data           = @{ Updated = $updated.ToArray(); Failed = $failed.ToArray(); StillListed = @() }
        }
    }
}

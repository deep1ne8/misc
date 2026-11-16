@{
    Id            = 'security-updates'
    Name          = 'Missing security updates'
    Kind          = 'Check'
    Category      = 'Updates'
    Description   = 'Asks Windows Update (or the WSUS/Intune source in force) which security and critical updates are not installed.'
    RequiresAdmin = $true
    Test          = {
        param($Ctx, $P)
        $security = '0fa1201d-4330-4fa8-8ae9-b877473b6441', 'e6cf1350-c01b-414d-a61f-263d14d133b4'   # Security Updates, Critical Updates
        try { $all = @(Get-AbPendingUpdate) }
        catch { return New-AbFinding -Code 'updates.search-failed' -Severity Warning -Title 'Could not search for updates' -Detail $_.Exception.Message -FixTask 'reset-windows-update' }
        $sec = @($all | Where-Object { @($_.CategoryIds | Where-Object { $security -contains $_ }).Count -or $_.Severity })
        if ($sec) {
            $crit = @($sec | Where-Object { $_.Severity -eq 'Critical' }).Count
            $sev = if ($crit) { 'Critical' } else { 'Warning' }
            New-AbFinding -Code 'updates.security-missing' -Severity $sev -Title "$($sec.Count) security update(s) not installed" -Detail ((@($sec | Select-Object -First 5 | ForEach-Object { if ($_.Severity) { "$($_.Title) [$($_.Severity)]" } else { $_.Title } })) -join '; ') `
                -FixTask 'install-security-updates' -Data @{ Count = $sec.Count; Critical = $crit; KBs = (@($sec | ForEach-Object KB | Where-Object { $_ }) -join ',') }
        }
        $other = @($all).Count - $sec.Count
        if ($other -gt 0) { New-AbFinding -Code 'updates.other-available' -Severity Info -Title "$other other update(s) available" -Detail 'Feature, quality or optional updates.' }
    }
}

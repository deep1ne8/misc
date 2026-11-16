@{
    Id          = 'time-sync'
    Name        = 'Clock accuracy'
    Kind        = 'Check'
    Category    = 'Network'
    Description = 'Compares the PC clock with time.windows.com. A wrong clock breaks sign-ins (Microsoft 365, Kerberos) and certificates.'
    Test        = {
        param($Ctx, $P)
        $offset = Get-AbTimeOffset
        if ($null -eq $offset) {
            return (New-AbFinding -Code 'time.unknown' -Severity Info -Title 'Clock offset could not be measured' -Detail 'time.windows.com was not reachable over NTP (UDP 123).')
        }
        $abs = [math]::Abs($offset)
        if ($abs -ge [double]$Ctx.Config.TimeSkewWarnSeconds) {
            $sev = if ($abs -ge 300) { 'Critical' } else { 'Warning' }
            New-AbFinding -Code 'time.skew' -Severity $sev -Title ("Clock is off by {0:N0} seconds" -f $abs) -Detail 'Sign-ins and secure websites can fail when the clock is more than 5 minutes off.' -FixTask 'resync-time' -Data @{ OffsetSeconds = $offset }
        }
    }
}

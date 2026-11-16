@{
    Id          = 'event-errors'
    Name        = 'System errors (last 24 h)'
    Kind        = 'Check'
    Category    = 'System'
    Description = 'Counts critical/error events in the System log and spots unexpected shutdowns and crashes in the last 7 days.'
    Test        = {
        param($Ctx, $P)
        $now = Get-AbDate
        $events = @(Get-AbWinEvent -FilterHashtable @{ LogName = 'System'; Level = 1, 2; StartTime = $now.AddHours(-24) })
        if ($events.Count -ge [int]$Ctx.Config.EventErrorWarn) {
            $top = ($events | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
            New-AbFinding -Code 'events.many-errors' -Severity Warning -Title "$($events.Count) system errors in the last 24 hours" -Detail "Top sources: $top" -Data @{ Count = $events.Count }
        }
        $week = @(Get-AbWinEvent -FilterHashtable @{ LogName = 'System'; Id = 41, 1001, 6008; StartTime = $now.AddDays(-7) })
        $power = @($week | Where-Object { $_.Id -in 41, 6008 }).Count
        $crash = @($week | Where-Object { $_.Id -eq 1001 -and "$($_.ProviderName)" -match 'BugCheck|WER-SystemErrorReporting' }).Count
        if ($power) { New-AbFinding -Code 'events.unexpected-shutdown' -Severity Warning -Title "$power unexpected shutdown(s) this week" -Detail 'Power loss, a forced power-off or a crash. Check power, battery and recent driver changes.' -Data @{ Count = $power } }
        if ($crash) { New-AbFinding -Code 'events.bugcheck' -Severity Critical -Title "$crash blue-screen crash(es) this week" -Detail 'Often a driver or hardware fault. Update drivers/BIOS and check memory.' -Data @{ Count = $crash } }
    }
}

@{
    Id          = 'event-logs'
    Name        = 'Event Viewer analysis'
    Kind        = 'Check'
    Category    = 'System'
    Description = 'Reads the System, Application and Setup logs plus the Defender, Windows Update and boot-performance logs, matches known problem events (crashes, blue screens, disk, driver, service, update, domain and malware events) and maps each to a fix.'
    Parameters  = @{ Days = 0 }
    Test        = {
        param($Ctx, $P)
        $days = if ([int]$P.Days -gt 0) { [int]$P.Days } else { [int]$Ctx.Config.EventLookbackDays }
        Invoke-AbEventAnalysis -Days $days

        # Anything else: a high volume of errors with no specific rule still deserves a look.
        $since = (Get-AbDate).AddHours(-24)
        foreach ($log in 'System', 'Application') {
            $errors = @(Get-AbWinEvent -FilterHashtable @{ LogName = $log; Level = 1, 2; StartTime = $since } |
                    Where-Object { -not (Test-AbEventNoise ([pscustomobject]@{ Provider = [string]$_.ProviderName; Id = [int]$_.Id })) })
            if ($errors.Count -ge [int]$Ctx.Config.EventErrorWarn) {
                $top = ($errors | Group-Object ProviderName | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '
                New-AbFinding -Code 'events.many-errors' -Severity Warning -Title "$($errors.Count) $log log errors in the last 24 hours" -Detail "Top sources: $top" -Data @{ Count = $errors.Count; Log = $log }
            }
        }
    }
}

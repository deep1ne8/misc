@{
    Id          = 'uptime'
    Name        = 'Time since last restart'
    Kind        = 'Check'
    Category    = 'System'
    Description = 'Flags PCs that have not restarted for a long time (updates and fixes often need a restart).'
    Test        = {
        param($Ctx, $P)
        $os = Get-AbCim -ClassName Win32_OperatingSystem
        $days = [math]::Round(((Get-AbDate) - [datetime]$os.LastBootUpTime).TotalDays, 1)
        if ($days -ge [double]$Ctx.Config.UptimeWarnDays) {
            New-AbFinding -Code 'system.uptime-long' -Severity Warning -Title "Not restarted for $days days" -Detail 'Restart the PC to apply updates and clear memory.' -Data @{ Days = $days }
        }
    }
}

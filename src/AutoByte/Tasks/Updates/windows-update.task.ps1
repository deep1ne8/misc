@{
    Id          = 'windows-update'
    Name        = 'Windows Update'
    Kind        = 'Check'
    Category    = 'Updates'
    Description = 'Windows Update services are usable and updates have installed recently.'
    Test        = {
        param($Ctx, $P)
        foreach ($s in @(Get-AbService -Name 'wuauserv', 'BITS')) {
            if ("$($s.StartType)" -eq 'Disabled') {
                New-AbFinding -Code 'updates.service-disabled' -Severity Critical -Title "$($s.DisplayName) is disabled" -Detail 'Windows cannot install security updates.' -FixTask 'reset-windows-update' -Data @{ Service = $s.Name }
            }
        }
        $hotfixes = @(Get-AbCim -ClassName Win32_QuickFixEngineering | Where-Object { $_.InstalledOn })
        if ($hotfixes) {
            $last = ($hotfixes | ForEach-Object { [datetime]$_.InstalledOn } | Sort-Object -Descending | Select-Object -First 1)
            $days = [int]((Get-AbDate) - $last).TotalDays
            if ($days -gt 45) {
                New-AbFinding -Code 'updates.stale' -Severity Warning -Title "No updates installed for $days days" -Detail ('Last update installed on ' + $last.ToString('yyyy-MM-dd') + '. Resetting Windows Update components often helps.') -FixTask 'reset-windows-update' -Data @{ Days = $days }
            }
        }
    }
}

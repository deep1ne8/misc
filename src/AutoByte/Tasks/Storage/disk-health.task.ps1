@{
    Id          = 'disk-health'
    Name        = 'Disk health'
    Kind        = 'Check'
    Category    = 'Storage'
    Description = 'Reports physical disks that Windows marks as unhealthy or degraded (possible hardware failure).'
    Test        = {
        param($Ctx, $P)
        try { $disks = @(Get-AbPhysicalDisk) }
        catch { return (New-AbFinding -Code 'disk.health-unknown' -Severity Info -Title 'Disk health could not be read' -Detail $_.Exception.Message) }
        foreach ($d in $disks) {
            if ("$($d.HealthStatus)" -ne 'Healthy') {
                New-AbFinding -Code 'disk.unhealthy' -Severity Critical -Title "Disk '$($d.FriendlyName)' reports $($d.HealthStatus)" -Detail 'Back up data now and plan a disk replacement.' -Data @{ Disk = "$($d.FriendlyName)"; Health = "$($d.HealthStatus)"; Status = "$($d.OperationalStatus)" }
            }
        }
    }
}

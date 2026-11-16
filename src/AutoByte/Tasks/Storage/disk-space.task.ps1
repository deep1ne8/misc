@{
    Id          = 'disk-space'
    Name        = 'Disk space'
    Kind        = 'Check'
    Category    = 'Storage'
    Description = 'Free space on every fixed drive; the system drive is held to stricter limits.'
    Test        = {
        param($Ctx, $P)
        $warn = [double]$Ctx.Config.DiskWarnPercent
        $crit = [double]$Ctx.Config.DiskCriticalPercent
        $sysDrive = if ($env:SystemDrive) { $env:SystemDrive } else { 'C:' }
        foreach ($d in @(Get-AbCim -ClassName Win32_LogicalDisk -Filter 'DriveType=3')) {
            if (-not $d.Size) { continue }
            $pct = [math]::Round(100 * [double]$d.FreeSpace / [double]$d.Size, 1)
            $freeGB = [math]::Round([double]$d.FreeSpace / 1GB, 1)
            $isSys = $d.DeviceID -eq $sysDrive
            $data = @{ Drive = $d.DeviceID; FreePercent = $pct; FreeGB = $freeGB }
            $fix = if ($isSys) { 'clean-temp-files' } else { '' }
            if ($pct -lt $crit) {
                New-AbFinding -Code 'disk.critical' -Severity Critical -Title "Drive $($d.DeviceID) is almost full ($pct% free)" -Detail "$freeGB GB free. Windows and apps can fail when the drive fills up." -FixTask $fix -Data $data
            }
            elseif ($pct -lt $warn) {
                New-AbFinding -Code 'disk.low' -Severity Warning -Title "Drive $($d.DeviceID) is low on space ($pct% free)" -Detail "$freeGB GB free." -FixTask $fix -Data $data
            }
        }
    }
}

@{
    Id          = 'drivers'
    Name        = 'Device drivers'
    Kind        = 'Check'
    Category    = 'Drivers'
    Description = 'Finds devices that Device Manager flags with a problem (missing, failed or disabled drivers) and driver updates offered by Windows Update.'
    Parameters  = @{ CheckUpdates = $true }
    Test        = {
        param($Ctx, $P)
        # ConfigManagerErrorCode meanings (Device Manager "Code NN").
        $codes = @{
            1 = 'not configured correctly'; 3 = 'driver may be corrupted'; 10 = 'cannot start'; 12 = 'resource conflict'
            14 = 'needs a restart'; 18 = 'driver needs reinstalling'; 19 = 'registry entry damaged'; 21 = 'being removed'
            22 = 'disabled'; 24 = 'not present or driver missing'; 28 = 'no driver installed'; 31 = 'driver could not load'
            32 = 'driver service disabled'; 37 = 'driver failed to initialise'; 39 = 'driver missing or corrupted'
            43 = 'stopped after reporting a problem'; 45 = 'not connected'; 48 = 'driver blocked (known issue)'; 52 = 'driver signature not verified'
        }
        $bad = @(Get-AbCim -ClassName Win32_PnPEntity -Filter 'ConfigManagerErrorCode <> 0' |
                Where-Object { [int]$_.ConfigManagerErrorCode -notin 22, 45 })   # disabled on purpose / unplugged
        foreach ($d in $bad | Select-Object -First 10) {
            $code = [int]$d.ConfigManagerErrorCode
            $why = if ($codes.ContainsKey($code)) { $codes[$code] } else { 'has a problem' }
            $fix = if ($code -in 1, 3, 18, 19, 24, 28, 31, 37, 39, 43, 52) { 'install-driver-updates' } elseif ($code -eq 14) { '' } else { 'rescan-devices' }
            New-AbFinding -Code 'drivers.device-problem' -Severity Warning -Title "$($d.Name): $why (Code $code)" -Detail "Device class: $($d.PNPClass). Hardware id: $(@($d.HardwareID)[0])" -FixTask $fix `
                -Data @{ Device = [string]$d.Name; ErrorCode = $code; DeviceId = [string]$d.PNPDeviceID }
        }
        if ($P.CheckUpdates) {
            try {
                $drv = @(Get-AbPendingUpdate -Criteria "IsInstalled=0 and IsHidden=0 and Type='Driver'")
                if ($drv) {
                    New-AbFinding -Code 'drivers.updates-available' -Severity Info -Title "$($drv.Count) driver update(s) available" -Detail ((@($drv | Select-Object -First 5 | ForEach-Object Title)) -join '; ') -FixTask 'install-driver-updates' -Data @{ Count = $drv.Count }
                }
            }
            catch { New-AbFinding -Code 'drivers.update-search-failed' -Severity Info -Title 'Could not search Windows Update for drivers' -Detail $_.Exception.Message }
        }
    }
}

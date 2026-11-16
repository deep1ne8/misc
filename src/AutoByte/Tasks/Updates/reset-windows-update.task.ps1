@{
    Id                  = 'reset-windows-update'
    Name                = 'Reset Windows Update components'
    Kind                = 'Fix'
    Category            = 'Updates'
    Description         = 'Re-enables disabled update services (set to Manual), stops them, renames the SoftwareDistribution and catroot2 caches (kept as .bak), and restarts the services. Runs when a service is disabled, or with -Force.'
    RequiresAdmin       = $true
    RebootMayBeRequired = $true
    Parameters          = @{ Force = $false }
    Test                = {
        param($Ctx, $P)
        $disabled = @(Get-AbService -Name 'wuauserv', 'BITS' | Where-Object { "$($_.StartType)" -eq 'Disabled' })
        if ($disabled) { return @{ Needed = $true; Reason = 'Disabled: ' + (($disabled | ForEach-Object Name) -join ', ') } }
        if ($P.Force -and -not $Ctx.AfterRepair) { return @{ Needed = $true; Reason = 'Forced reset.' } }
        @{ Needed = $false; Reason = 'Update services are enabled. Use -Force to reset the caches anyway.' }
    }
    Repair              = {
        param($Ctx, $P)
        $services = 'wuauserv', 'BITS', 'CryptSvc', 'msiserver'
        foreach ($s in @(Get-AbService -Name 'wuauserv', 'BITS' | Where-Object { "$($_.StartType)" -eq 'Disabled' })) {
            Set-AbServiceStartType -Name $s.Name -StartupType Manual
            Add-AbMessage -Result $Ctx.Result -Message "Set $($s.Name) to Manual (was Disabled)."
        }
        foreach ($s in $services) {
            $svc = Get-AbService -Name $s
            if ($svc -and "$($svc.Status)" -eq 'Running') { Set-AbServiceState -Name $s -State Stopped }
        }
        try {
            $stamp = (Get-AbDate).ToString('yyyyMMddHHmmss')
            foreach ($name in 'SoftwareDistribution', 'Catroot2') {
                $path = Get-AbSystemPath -Name $name
                if (Test-AbPath -Path $path) {
                    $leaf = Split-Path -Leaf $path
                    Rename-AbItem -Path $path -NewName "$leaf.bak-$stamp"
                    Add-AbMessage -Result $Ctx.Result -Message "Renamed $leaf to $leaf.bak-$stamp."
                }
            }
        }
        finally {
            foreach ($s in 'CryptSvc', 'BITS', 'wuauserv') {
                try { Set-AbServiceState -Name $s -State Running } catch { Add-AbMessage -Result $Ctx.Result -Message "Could not start ${s}: $($_.Exception.Message)" }
            }
        }
        @{ RebootRequired = $true }
    }
}

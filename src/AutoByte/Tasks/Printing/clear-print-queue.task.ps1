@{
    Id            = 'clear-print-queue'
    Name          = 'Clear stuck print jobs'
    Kind          = 'Fix'
    Category      = 'Printing'
    Description   = 'Stops the Print Spooler, deletes queued job files older than 15 minutes, and starts the spooler again.'
    RequiresAdmin = $true
    EndUserSafe   = $true
    Test          = {
        param($Ctx, $P)
        $svc = Get-AbService -Name 'Spooler'
        if (-not $svc -or "$($svc.StartType)" -eq 'Disabled') { return @{ Needed = $false; Reason = 'Print Spooler is not installed or is disabled.' } }
        $old = (Get-AbDate).AddMinutes(-15)
        $stuck = @(Get-AbChildItem -Path (Get-AbSystemPath -Name Spool) | Where-Object { $_.LastWriteTime -lt $old })
        $Ctx.Stuck = $stuck
        if ("$($svc.Status)" -ne 'Running') { return @{ Needed = $true; Reason = 'Print Spooler is stopped.' } }
        if ($stuck) { return @{ Needed = $true; Reason = "$($stuck.Count) old spool file(s)." } }
        @{ Needed = $false; Reason = 'Spooler running, no stuck jobs.' }
    }
    Repair        = {
        param($Ctx, $P)
        $svc = Get-AbService -Name 'Spooler'
        if ("$($svc.Status)" -eq 'Running') { Set-AbServiceState -Name 'Spooler' -State Stopped }
        try {
            $n = 0
            foreach ($f in @($Ctx.Stuck)) { if (Remove-AbFile -Path $f.FullName) { $n++ } }
            Add-AbMessage -Result $Ctx.Result -Message "Removed $n stuck spool file(s)."
        }
        finally { Set-AbServiceState -Name 'Spooler' -State Running }
        @{}
    }
}

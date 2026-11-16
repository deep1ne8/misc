@{
    Id          = 'print-queue'
    Name        = 'Printing'
    Kind        = 'Check'
    Category    = 'Printing'
    Description = 'Print Spooler is running and no print jobs are stuck.'
    Test        = {
        param($Ctx, $P)
        $svc = Get-AbService -Name 'Spooler'
        if (-not $svc) { return }
        if ("$($svc.StartType)" -eq 'Disabled') { return }
        if ("$($svc.Status)" -ne 'Running') {
            return (New-AbFinding -Code 'print.spooler-stopped' -Severity Warning -Title 'Print Spooler is not running' -Detail 'Nothing can print until it is started.' -FixTask 'clear-print-queue')
        }
        $old = (Get-AbDate).AddMinutes(-15)
        $stuck = @(Get-AbChildItem -Path (Get-AbSystemPath -Name Spool) | Where-Object { $_.LastWriteTime -lt $old })
        if ($stuck) {
            New-AbFinding -Code 'print.stuck-jobs' -Severity Warning -Title "$([math]::Ceiling($stuck.Count / 2)) print job(s) appear stuck" -Detail 'Jobs older than 15 minutes are still in the queue.' -FixTask 'clear-print-queue' -Data @{ Files = $stuck.Count }
        }
    }
}

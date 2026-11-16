@{
    Id          = 'performance'
    Name        = 'Why is it slow?'
    Kind        = 'Check'
    Category    = 'Performance'
    Description = 'Samples CPU, memory commit and disk activity, names the processes using the most, and checks the power plan, the system disk type, start-up apps and the last boot time.'
    Parameters  = @{ Samples = 3 }
    Test        = {
        param($Ctx, $P)
        $c = $Ctx.Config
        $cpu = New-Object System.Collections.Generic.List[double]
        $busy = New-Object System.Collections.Generic.List[double]
        for ($i = 0; $i -lt [Math]::Max(1, [int]$P.Samples); $i++) {
            if ($i) { Wait-AbInterval -Milliseconds 1000 }
            $cpuNow = @(Get-AbCim -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'")
            if ($cpuNow) { $cpu.Add([double]$cpuNow[0].PercentProcessorTime) }
            $d = @(Get-AbCim -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'")
            if ($d) { $busy.Add(100 - [double]$d[0].PercentIdleTime) }
        }
        $procs = @(Get-AbCim -ClassName Win32_PerfFormattedData_PerfProc_Process | Where-Object { $_.Name -notin '_Total', 'Idle' })
        $cs = @(Get-AbCim -ClassName Win32_ComputerSystem)
        $cores = if ($cs) { [Math]::Max(1, [int]$cs[0].NumberOfLogicalProcessors) } else { 1 }
        $fmtTop = { param($list, $prop, $unit) ($list | Sort-Object $prop -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name -replace '#\d+$') ($(& $unit $_))" }) -join ', ' }

        if ($cpu.Count) {
            $avg = [Math]::Round(($cpu | Measure-Object -Average).Average)
            if ($avg -ge [int]$c.CpuWarnPercent) {
                $top = & $fmtTop $procs 'PercentProcessorTime' { param($x) "$([Math]::Round([double]$x.PercentProcessorTime / $cores))%" }
                New-AbFinding -Code 'performance.cpu-high' -Severity Warning -Title "CPU is $avg% busy" -Detail "Heaviest processes: $top. Close or update them; a scan or update may also be running." -Data @{ Percent = $avg; Top = $top }
            }
        }
        $mem = @(Get-AbCim -ClassName Win32_PerfFormattedData_PerfOS_Memory)
        if ($mem) {
            $commit = [int]$mem[0].PercentCommittedBytesInUse
            if ($commit -ge [int]$c.CommitWarnPercent) {
                $top = & $fmtTop $procs 'WorkingSetPrivate' { param($x) "$([Math]::Round([double]$x.WorkingSetPrivate / 1MB)) MB" }
                New-AbFinding -Code 'performance.memory-pressure' -Severity Warning -Title "Memory is $commit% committed" -Detail "Windows is paging to disk, which makes everything slow. Largest: $top. Close apps or add RAM." -Data @{ Percent = $commit; Top = $top }
            }
        }
        if ($busy.Count) {
            $avg = [Math]::Round(($busy | Measure-Object -Average).Average)
            if ($avg -ge [int]$c.DiskBusyWarnPercent) {
                $top = & $fmtTop $procs 'IODataBytesPersec' { param($x) "$([Math]::Round([double]$x.IODataBytesPersec / 1MB, 1)) MB/s" }
                New-AbFinding -Code 'performance.disk-busy' -Severity Warning -Title "Disk is $avg% busy" -Detail "The disk is the bottleneck. Busiest: $top." -Data @{ Percent = $avg; Top = $top }
            }
        }
        if ((Get-AbSystemDiskMediaType) -eq 'HDD') {
            New-AbFinding -Code 'performance.hdd-system-disk' -Severity Warning -Title 'Windows runs from a spinning hard disk' -Detail 'This is the most common cause of a slow PC. Replacing it with an SSD is the biggest improvement available.'
        }
        $plan = @(try { Get-AbCim -Namespace 'root\cimv2\power' -ClassName Win32_PowerPlan | Where-Object { $_.IsActive } } catch { Write-Verbose "Power plan unavailable: $_" })
        if ($plan -and "$($plan[0].InstanceID)" -match 'a1841308-3541-4fab-bc81-f71556f20b4a') {
            New-AbFinding -Code 'performance.power-saver' -Severity Warning -Title 'Power saver plan is active' -Detail 'The CPU is held at low speed.' -FixTask 'set-power-plan'
        }
        $startup = @(try { Get-AbCim -ClassName Win32_StartupCommand } catch { Write-Verbose "Startup list unavailable: $_" })
        if ($startup.Count -ge [int]$c.StartupItemsWarn) {
            New-AbFinding -Code 'performance.startup-items' -Severity Info -Title "$($startup.Count) apps start with Windows" -Detail 'Disable the ones you do not need in Task Manager > Startup apps.' -Data @{ Count = $startup.Count; Items = (@($startup | Select-Object -First 15 | ForEach-Object Name) -join ', ') }
        }
        $boot = @(Get-AbEventRecord -LogName 'Microsoft-Windows-Diagnostics-Performance/Operational' -Id 100 -StartTime (Get-AbDate).AddDays(-30) -MaxEvents 1)
        if ($boot -and $boot[0].Data['BootTime']) {
            $sec = [Math]::Round([double]$boot[0].Data['BootTime'] / 1000)
            if ($sec -ge [int]$c.BootWarnSeconds) {
                New-AbFinding -Code 'performance.slow-boot' -Severity Warning -Title "Last start-up took $sec seconds" -Detail "Run 'abyte run event-logs' to see which apps, drivers or services slowed it." -Data @{ Seconds = $sec }
            }
        }
    }
}

@{
    Id          = 'memory'
    Name        = 'Available memory'
    Kind        = 'Check'
    Category    = 'System'
    Description = 'Low free RAM makes the PC slow.'
    Test        = {
        param($Ctx, $P)
        $os = Get-AbCim -ClassName Win32_OperatingSystem
        if (-not $os.TotalVisibleMemorySize) { return }
        $pct = [math]::Round(100 * [double]$os.FreePhysicalMemory / [double]$os.TotalVisibleMemorySize, 1)
        if ($pct -lt [double]$Ctx.Config.MemoryWarnPercent) {
            New-AbFinding -Code 'system.memory-low' -Severity Warning -Title "Only $pct% of memory is free" -Detail 'Close unused programs or restart. If this is constant, the PC may need more RAM.' -Data @{ FreePercent = $pct }
        }
    }
}

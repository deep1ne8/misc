@{
    Id          = 'reset-network-cache'
    Name        = 'Flush DNS cache'
    Kind        = 'Fix'
    Category    = 'Network'
    Description = 'Clears the DNS resolver cache when name lookups fail (or -Force). Does not reset adapters or Winsock.'
    EndUserSafe = $true
    Parameters  = @{ Force = $false }
    Test        = {
        param($Ctx, $P)
        $hosts = @("$($Ctx.Config.ConnectivityHosts)" -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $failing = @($hosts | Where-Object { -not (Resolve-AbDns -Name $_) })
        if ($Ctx.AfterRepair) {
            # Flushing cannot fix a cable or Wi-Fi problem; report it, but the flush itself succeeded.
            if ($failing) { Add-AbMessage -Result $Ctx.Result -Message "Still failing after flush: $($failing -join ', '). Check the connection." }
            return @{ Needed = $false; Reason = 'Cache flushed.' }
        }
        if ($failing) { return @{ Needed = $true; Reason = "Lookups failing: $($failing -join ', ')." } }
        if ($P.Force) { return @{ Needed = $true; Reason = 'Forced flush.' } }
        @{ Needed = $false; Reason = 'DNS lookups are working.' }
    }
    Repair      = {
        param($Ctx, $P)
        $r = Invoke-AbNative -FilePath (Get-AbSystemPath -Name IpConfig) -ArgumentList '/flushdns' -TimeoutSeconds 60
        if ($r.ExitCode -ne 0) { throw "ipconfig /flushdns failed with exit code $($r.ExitCode)." }
        Add-AbMessage -Result $Ctx.Result -Message 'DNS cache flushed.'
        @{}
    }
}

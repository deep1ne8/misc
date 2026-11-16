@{
    Id          = 'connectivity'
    Name        = 'Internet and DNS'
    Kind        = 'Check'
    Category    = 'Network'
    Description = 'Resolves and connects (HTTPS port 443) to well-known hosts.'
    Test        = {
        param($Ctx, $P)
        $hosts = @("$($Ctx.Config.ConnectivityHosts)" -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $dnsFail = @($hosts | Where-Object { -not (Resolve-AbDns -Name $_) })
        $tcpFail = @($hosts | Where-Object { $_ -notin $dnsFail -and -not (Test-AbTcp -HostName $_ -Port 443) })
        if ($dnsFail.Count -eq $hosts.Count) {
            New-AbFinding -Code 'network.dns-down' -Severity Critical -Title 'Cannot look up internet addresses (DNS)' -Detail 'Check the network cable/Wi-Fi. Flushing the DNS cache may help.' -FixTask 'reset-network-cache' -Data @{ Hosts = $dnsFail }
        }
        elseif ($dnsFail) {
            New-AbFinding -Code 'network.dns-partial' -Severity Warning -Title "Some addresses do not resolve: $($dnsFail -join ', ')" -FixTask 'reset-network-cache' -Data @{ Hosts = $dnsFail }
        }
        if ($tcpFail.Count -and ($tcpFail.Count + $dnsFail.Count) -eq $hosts.Count) {
            New-AbFinding -Code 'network.offline' -Severity Critical -Title 'No secure internet connection' -Detail 'Names resolve but HTTPS connections fail: firewall, proxy or captive portal (hotel/guest Wi-Fi)?' -Data @{ Hosts = $tcpFail }
        }
        elseif ($tcpFail) {
            New-AbFinding -Code 'network.blocked' -Severity Warning -Title "Cannot connect to: $($tcpFail -join ', ')" -Detail 'A firewall or proxy may be blocking these services.' -Data @{ Hosts = $tcpFail }
        }
    }
}

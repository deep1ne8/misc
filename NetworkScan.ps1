Get-NetNeighbor -AddressFamily IPv4 `
    | where {$_.State -ne 'Unreachable' -and $_.State -ne 'Permanent'} `
    | sort LinkLayerAddress -Unique `
    | Format-Table -AutoSize IPAddress, LinkLayerAddress, State

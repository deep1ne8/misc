#requires -Version 5.1

# NativeNetworkScanner 1.6
<##
.SYNOPSIS
    Interactive native PowerShell IPv4 network scanner.

.DESCRIPTION
    Advanced IP Scanner-style network discovery written as one self-contained
    PowerShell script. Uses runspaces for concurrent host scanning and built-in
    Windows/.NET APIs only.

    Features include:
      - Automatic IPv4 subnet detection
      - CIDR and IPv4 range scanning
      - Concurrent ICMP ping sweep
      - Reverse DNS lookup
      - ARP / neighbor-table MAC discovery
      - Embedded and optional refreshed IEEE OUI lookup
      - Configurable TCP port scanning
      - TTL-based OS heuristics
      - Device fingerprinting for EZVIZ/Hikvision cameras, Android devices,
        Apple devices, Android TV / Google TV boxes, laptops, desktops,
        servers, printers, routers and network appliances
      - Optional lightweight HTTP fingerprinting on open HTTP ports
      - Live progress and color-coded results
      - Interactive filtering, sorting, host details and exports
      - CSV and JSON export

.PARAMETER MaxConcurrency
    Maximum number of concurrent runspace workers. Default 64.

.PARAMETER PingTimeoutMs
    ICMP timeout per host. Default 700 ms.

.PARAMETER PortTimeoutMs
    TCP connection timeout per port. Default 250 ms.

.PARAMETER Ports
    TCP ports to probe. An array is supported. Default includes common
    infrastructure, Windows, camera, mobile and TV discovery ports.

.PARAMETER MaxTargets
    Maximum addresses processed by one scan. Default 4096.

.PARAMETER RefreshOui
    Download the current IEEE OUI text file before starting.

.PARAMETER NoDns
    Skip reverse DNS lookups.

.EXAMPLE
    .\NativeNetworkScanner.ps1

.EXAMPLE
    .\NativeNetworkScanner.ps1 -Ports 22,80,443,445,3389,554,5555,8000,8008,8009,62078

.EXAMPLE
    .\NativeNetworkScanner.ps1 -MaxConcurrency 96 -PingTimeoutMs 500 -PortTimeoutMs 200

.NOTES
    Standard ICMP, TCP and local ARP/neighbor-table access normally do not
    require elevation. Some Windows security policies can restrict neighbor
    information or adapter data.

    Internet access is used only when -RefreshOui is specified.

    Device classification is heuristic. It should be treated as an identification
    aid, not a definitive hardware inventory source. Randomized MAC addresses,
    disabled ICMP, NAT, VLANs, VPNs and firewalls can reduce identification quality.
#>

[CmdletBinding()]
param(
    [ValidateRange(4,256)]
    [int]$MaxConcurrency = 64,

    [ValidateRange(100,5000)]
    [int]$PingTimeoutMs = 700,

    [ValidateRange(50,3000)]
    [int]$PortTimeoutMs = 250,

    [int[]]$Ports = @(
        21,22,23,25,53,80,110,135,139,143,
        443,445,515,548,554,5900,631,2049,32400,
        3389,5000,5001,5060,5061,5555,5985,5986,62078,
        6466,6690,7000,7100,8000,8001,8002,8006,8008,8009,
        8080,8200,8443,8444,8899,3000,3001,9100
    ),

    [ValidateRange(1,65535)]
    [int]$MaxTargets = 4096,

    [switch]$RefreshOui,
    [switch]$NoDns
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ($null -eq $Ports -or @($Ports).Count -eq 0) {
    throw 'Ports must contain at least one TCP port.'
}

foreach ($port in @($Ports)) {
    if ($port -lt 1 -or $port -gt 65535) {
        throw "Invalid TCP port '$port'. Valid ports are 1 through 65535."
    }
}

$script:AppName = 'NativeNetworkScanner'
$script:OuiCachePath = Join-Path (
    [Environment]::GetFolderPath('LocalApplicationData')
) "$($script:AppName)\oui.txt"
$script:Ports = @($Ports | Sort-Object -Unique)
$script:NoDns = [bool]$NoDns
$script:CurrentResults = @()
$script:LastTargetSpec = $null
$script:LastDescription = $null
$script:LastStats = $null
$script:LastGatewayIPs = @()

$script:PortNames = @{
    21    = 'FTP'
    22    = 'SSH'
    23    = 'Telnet'
    25    = 'SMTP'
    53    = 'DNS'
    80    = 'HTTP'
    110   = 'POP3'
    135   = 'RPC'
    139   = 'NetBIOS'
    143   = 'IMAP'
    443   = 'HTTPS'
    445   = 'SMB'
    515   = 'LPD'
    554   = 'RTSP'
    5900  = 'VNC'
    631   = 'IPP'
    3389  = 'RDP'
    5555  = 'ADB'
    5985  = 'WinRM-HTTP'
    5986  = 'WinRM-HTTPS'
    62078 = 'iOS'
    6466  = 'Google-TV'
    7000  = 'AirPlay'
    7100  = 'AirPlay'
    8000  = 'Hikvision-SDK'
    8001  = 'Samsung-TV'
    8002  = 'Samsung-TV-TLS'
    8006  = 'Proxmox'
    8008  = 'Google-Cast'
    8009  = 'Chromecast'
    8080  = 'HTTP-Alt'
    8200  = 'Camera-Service'
    8443  = 'HTTPS-Alt'
    8899  = 'Camera/IoT'
    9100  = 'JetDirect'
    548   = 'AFP'
    5000  = 'NAS-HTTP'
    5001  = 'NAS-HTTPS'
    5060  = 'SIP'
    5061  = 'SIP-TLS'
    2049  = 'NFS'
    3000  = 'LG-webOS'
    3001  = 'LG-webOS-TLS'
    32400 = 'Plex'
    6690  = 'Qsync'
    8444  = 'IoT-HTTPS'
}

function Test-IPv4 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )

    try {
        $ip = [System.Net.IPAddress]::Parse($Address.Trim())
        return ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork)
    }
    catch {
        return $false
    }
}

function Convert-IPv4ToUInt32 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Address
    )

    $bytes = ([System.Net.IPAddress]::Parse($Address)).GetAddressBytes()
    if ($bytes.Count -ne 4) {
        throw "Not an IPv4 address: $Address"
    }

    return ([uint32]$bytes[0] * 16777216) +
        ([uint32]$bytes[1] * 65536) +
        ([uint32]$bytes[2] * 256) +
        [uint32]$bytes[3]
}

function Convert-UInt32ToIPv4 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [uint32]$Value
    )

    return '{0}.{1}.{2}.{3}' -f `
        ([byte](($Value -shr 24) -band 255)), `
        ([byte](($Value -shr 16) -band 255)), `
        ([byte](($Value -shr 8) -band 255)), `
        ([byte]($Value -band 255))
}

function Convert-PrefixToMask {
    [CmdletBinding()]
    param(
        [ValidateRange(0,32)]
        [int]$Prefix
    )

    if ($Prefix -eq 0) {
        return [uint32]0
    }

    return [uint32](
        [uint64]4294967295 -
        ([uint64]([math]::Pow(2, 32 - $Prefix)) - 1)
    )
}

function Convert-MaskToPrefix {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Mask
    )

    $value = Convert-IPv4ToUInt32 $Mask
    $zeroSeen = $false
    $prefix = 0

    for ($i = 31; $i -ge 0; $i--) {
        $bit = ($value -shr $i) -band 1

        if ($bit) {
            if ($zeroSeen) {
                throw "Non-contiguous subnet mask: $Mask"
            }
            $prefix++
        }
        else {
            $zeroSeen = $true
        }
    }

    return $prefix
}

function Get-NetworkAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IPAddress,

        [Parameter(Mandatory)]
        [ValidateRange(0,32)]
        [int]$PrefixLength
    )

    $mask = Convert-PrefixToMask $PrefixLength
    $ip = Convert-IPv4ToUInt32 $IPAddress
    return Convert-UInt32ToIPv4 ([uint32]($ip -band $mask))
}

function Get-IPv4Targets {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Spec,

        [int]$MaxCount = 4096
    )

    $target = $Spec.Trim()
    if ([string]::IsNullOrWhiteSpace($target)) {
        throw 'Target is empty.'
    }

    $list = New-Object 'System.Collections.Generic.List[string]'

    if ($target -match '^\s*([^/\s]+)/\s*(\d{1,2})\s*$') {
        $base = $matches[1]
        $prefix = [int]$matches[2]

        if (!(Test-IPv4 $base) -or $prefix -lt 0 -or $prefix -gt 32) {
            throw "Invalid CIDR: $Spec"
        }

        $network = [uint32](Convert-IPv4ToUInt32 (Get-NetworkAddress $base $prefix))

        if ($prefix -eq 32) {
            $first = [uint64]$network
            $last = [uint64]$network
        }
        else {
            $count = [uint64][math]::Pow(2, 32 - $prefix)
            $first = [uint64]$network
            $last = $first + $count - 1

            if ($prefix -le 30) {
                $first++
                $last--
            }
        }

        $hostCount = $last - $first + 1
        if ($hostCount -gt $MaxCount) {
            throw "Target count $hostCount exceeds MaxTargets=$MaxCount."
        }

        for ($n = $first; $n -le $last; $n++) {
            [void]$list.Add((Convert-UInt32ToIPv4 ([uint32]$n)))
        }

        return @($list)
    }

    if ($target -match '^\s*(\S+)\s*-\s*(\S+)\s*$') {
        $left = $matches[1]
        $right = $matches[2]

        if (!(Test-IPv4 $left)) {
            throw "Invalid range start: $left"
        }

        if ($right -match '^\d{1,3}$') {
            $octets = $left.Split('.')
            $right = "$($octets[0]).$($octets[1]).$($octets[2]).$right"
        }

        if (!(Test-IPv4 $right)) {
            throw "Invalid range end: $right"
        }

        $first = [uint64](Convert-IPv4ToUInt32 $left)
        $last = [uint64](Convert-IPv4ToUInt32 $right)

        if ($last -lt $first) {
            throw 'Range end must be greater than or equal to range start.'
        }

        $count = $last - $first + 1
        if ($count -gt $MaxCount) {
            throw "Target count $count exceeds MaxTargets=$MaxCount."
        }

        for ($n = $first; $n -le $last; $n++) {
            [void]$list.Add((Convert-UInt32ToIPv4 ([uint32]$n)))
        }

        return @($list)
    }

    if (Test-IPv4 $target) {
        return @($target)
    }

    throw "Unsupported target: $Spec. Use 192.168.1.0/24 or 192.168.1.1-254."
}

function Get-LocalIPv4Networks {
    [CmdletBinding()]
    param()

    $output = New-Object 'System.Collections.Generic.List[object]'
    $seen = @{}

    try {
        foreach ($config in @(Get-NetIPConfiguration -ErrorAction Stop | Where-Object {
            $_.NetAdapter -and $_.NetAdapter.Status -eq 'Up'
        })) {
            foreach ($address in @($config.IPv4Address)) {
                if (!$address) {
                    continue
                }

                $ip = [string]$address.IPAddress
                $prefix = [int]$address.PrefixLength

                if (!(Test-IPv4 $ip) -or $ip -like '127.*' -or $ip -like '169.254.*') {
                    continue
                }

                $network = Get-NetworkAddress $ip $prefix
                $key = "$network/$prefix"

                if ($seen.ContainsKey($key)) {
                    continue
                }

                $seen[$key] = $true

                $gateway = $null
                if ($config.IPv4DefaultGateway) {
                    $gateway = [string]$config.IPv4DefaultGateway.NextHop
                }

                [void]$output.Add([pscustomobject]@{
                    InterfaceAlias = [string]$config.InterfaceAlias
                    IPAddress       = $ip
                    PrefixLength    = $prefix
                    Network         = $network
                    CIDR            = $key
                    Gateway         = $gateway
                })
            }
        }
    }
    catch {
    }

    if (!$output.Count) {
        try {
            foreach ($config in @(Get-CimInstance Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = TRUE' -ErrorAction Stop)) {
                $ips = @($config.IPAddress)
                $masks = @($config.IPSubnet)

                for ($i = 0; $i -lt $ips.Count; $i++) {
                    $ip = [string]$ips[$i]

                    if (!(Test-IPv4 $ip) -or $ip -like '127.*' -or $ip -like '169.254.*' -or $i -ge $masks.Count) {
                        continue
                    }

                    $prefix = Convert-MaskToPrefix ([string]$masks[$i])
                    $network = Get-NetworkAddress $ip $prefix
                    $key = "$network/$prefix"

                    if ($seen.ContainsKey($key)) {
                        continue
                    }

                    $seen[$key] = $true
                    $gateway = @(
                        $config.DefaultIPGateway |
                        Where-Object { Test-IPv4 ([string]$_) } |
                        Select-Object -First 1
                    )

                    [void]$output.Add([pscustomobject]@{
                        InterfaceAlias = [string]$config.Description
                        IPAddress       = $ip
                        PrefixLength    = $prefix
                        Network         = $network
                        CIDR            = $key
                        Gateway         = $(if ($gateway.Count) { [string]$gateway[0] } else { $null })
                    })
                }
            }
        }
        catch {
        }
    }

    if (!$output.Count) {
        try {
            foreach ($adapter in [System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()) {
                if ($adapter.OperationalStatus -ne 'Up' -or
                    $adapter.NetworkInterfaceType -eq [System.Net.NetworkInformation.NetworkInterfaceType]::Loopback) {
                    continue
                }

                $ipProperties = $adapter.GetIPProperties()

                foreach ($unicast in @($ipProperties.UnicastAddresses)) {
                    if (!$unicast -or $unicast.Address.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
                        continue
                    }

                    $ip = $unicast.Address.ToString()

                    if ($ip -like '127.*' -or $ip -like '169.254.*') {
                        continue
                    }

                    $prefix = Convert-MaskToPrefix $unicast.IPv4Mask.ToString()
                    $network = Get-NetworkAddress $ip $prefix
                    $key = "$network/$prefix"

                    if ($seen.ContainsKey($key)) {
                        continue
                    }

                    $seen[$key] = $true

                    $gateway = @(
                        $ipProperties.GatewayAddresses |
                        Where-Object { $_.Address.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
                        ForEach-Object { $_.Address.ToString() } |
                        Select-Object -First 1
                    )

                    [void]$output.Add([pscustomobject]@{
                        InterfaceAlias = [string]$adapter.Name
                        IPAddress       = $ip
                        PrefixLength    = $prefix
                        Network         = $network
                        CIDR            = $key
                        Gateway         = $(if ($gateway.Count) { [string]$gateway[0] } else { $null })
                    })
                }
            }
        }
        catch {
        }
    }

    return @($output | Sort-Object InterfaceAlias, Network)
}

function Import-OuiDatabase {
    [CmdletBinding()]
    param()

    $database = @{}

    $embedded = @{
        '00000C' = 'Cisco Systems, Inc.'
        '00037F' = 'Atheros Communications, Inc.'
        '000C29' = 'VMware, Inc.'
        '0015E9' = 'Dell Inc.'
        '00155D' = 'Microsoft Corporation'
        '001B21' = 'Intel Corporate'
        '001B63' = 'Apple, Inc.'
        '001C23' = 'Hewlett-Packard Company'
        '001DD8' = 'Microsoft Corporation'
        '0025AE' = 'Microsoft Corporation'
        '002608' = 'Apple, Inc.'
        '003048' = 'Super Micro Computer, Inc.'
        '005056' = 'VMware, Inc.'
        '080027' = 'PCS Systemtechnik GmbH'
        '18E829' = 'Ubiquiti Inc.'
        '3C0754' = 'Apple, Inc.'
        '3C5A37' = 'TP-Link Technologies Co., Ltd.'
        '3CA82A' = 'Hewlett-Packard Company'
        '3C970E' = 'Intel Corporate'
        '50C7BF' = 'TP-Link Technologies Co., Ltd.'
        '525400' = 'QEMU'
        '78E3B5' = 'Dell Inc.'
        '80CE62' = 'Lenovo Group Limited'
        '8C8590' = 'Samsung Electronics Co.,Ltd'
        '98DAC4' = 'TP-Link Technologies Co., Ltd.'
        'A4BADB' = 'Intel Corporate'
        'AC84C6' = 'TP-Link Technologies Co., Ltd.'
        'B4D5BD' = 'Intel Corporate'
        'B827EB' = 'Raspberry Pi Trading Ltd'
        'C83A35' = 'ASUSTek COMPUTER INC.'
        'D850E6' = 'Samsung Electronics Co.,Ltd'
        'DC4A3E' = 'Raspberry Pi Trading Ltd'
        'E063DA' = 'Ubiquiti Inc.'
        'F09FC2' = 'Ubiquiti Inc.'
        'F4F5D8' = 'Apple, Inc.'
        '001A11' = 'Google, Inc.'
        '7C49EB' = 'Google, Inc.'
        'AC3743' = 'Google, Inc.'
        '001451' = 'Cisco Systems, Inc.'
        '001C42' = 'Parallels, Inc.'
        '001E67' = 'Samsung Electronics Co.,Ltd'
        '0023AE' = 'Dell Inc.'
        '0024E8' = 'Samsung Electronics Co.,Ltd'
        '0023D4' = 'Dell Inc.'
        '3C84A6' = 'Xiaomi Communications Co Ltd'
        '644BFC' = 'Xiaomi Communications Co Ltd'
        'AC233F' = 'Xiaomi Communications Co Ltd'
        '28FF3C' = 'Xiaomi Communications Co Ltd'
        '001E10' = 'Hewlett-Packard Company'
        '0026B9' = 'Hewlett-Packard Company'
        '3C4A92' = 'Lenovo Group Limited'
        'E8B1FC' = 'Lenovo Group Limited'
        'D8CB8A' = 'ASUSTek COMPUTER INC.'
        'BC5FF4' = 'ASUSTek COMPUTER INC.'
        'B06EBF' = 'Acer, Inc.'
        '00E18C' = 'Acer, Inc.'
        '001C25' = 'Acer, Inc.'
        '001C26' = 'Acer, Inc.'
        '001A4B' = 'Sony Corporation'
        '0024BE' = 'Sony Corporation'
        '001D0D' = 'LG Electronics'
        '001E75' = 'LG Electronics'
        '001C62' = 'Huawei Technologies Co., Ltd'
        '00259E' = 'Huawei Technologies Co., Ltd'
        '001A79' = 'HTC Corporation'
        '0017D5' = 'Motorola, Inc'
        '001D25' = 'Motorola, Inc'
        '001F00' = 'Samsung Electronics Co.,Ltd'
        '0015B9' = 'Samsung Electronics Co.,Ltd'
        '0018AF' = 'Samsung Electronics Co.,Ltd'
        '002339' = 'Samsung Electronics Co.,Ltd'
        '001D43' = 'Nintendo Co., Ltd.'
        '00155A' = 'Nintendo Co., Ltd.'
        '001C2B' = 'Amazon Technologies Inc.'
        '00FC8B' = 'Amazon Technologies Inc.'
        '0020A6' = 'Amlogic, Inc.'
        'A0A8CD' = 'Amlogic, Inc.'
        'CC08FA' = 'Roku, Inc.'
        'B83861' = 'Roku, Inc.'
        '0019C5' = 'Hikvision Digital Technology Co., Ltd.'
        '4419B6' = 'Hikvision Digital Technology Co., Ltd.'
        'C0E3FB' = 'Hikvision Digital Technology Co., Ltd.'
        'B0C745' = 'Hikvision Digital Technology Co., Ltd.'
        '0023B2' = 'D-Link Corporation'
        'C0A0BB' = 'D-Link Corporation'
        '001F33' = 'Belkin International, Inc.'
        'EC1A59' = 'Belkin International, Inc.'
    }

    foreach ($key in $embedded.Keys) {
        $database[$key] = $embedded[$key]
    }

    if (Test-Path -LiteralPath $script:OuiCachePath) {
        try {
            foreach ($line in [IO.File]::ReadLines($script:OuiCachePath)) {
                if ($line -match '^\s*([0-9A-Fa-f]{2})[:-]([0-9A-Fa-f]{2})[:-]([0-9A-Fa-f]{2})\s+\(hex\)\s+(.+?)\s*$') {
                    $key = ($matches[1] + $matches[2] + $matches[3]).ToUpperInvariant()
                    $database[$key] = $matches[4].Trim()
                }
                elseif ($line -match '^\s*([0-9A-Fa-f]{6})\s+(.+?)\s*$') {
                    $database[$matches[1].ToUpperInvariant()] = $matches[2].Trim()
                }
            }
        }
        catch {
        }
    }

    return $database
}

$script:OuiDatabase = Import-OuiDatabase

function Format-MacAddress {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Mac
    )

    if ([string]::IsNullOrWhiteSpace($Mac)) {
        return $null
    }

    $hex = $Mac -replace '[^0-9A-Fa-f]', ''
    if ($hex.Length -ne 12) {
        return $null
    }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt 12; $i += 2) {
        [void]$parts.Add($hex.Substring($i, 2))
    }

    return ($parts -join ':').ToUpperInvariant()
}

function Resolve-MacVendor {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Mac
    )

    $formatted = Format-MacAddress $Mac
    if (!$formatted) {
        return 'Unknown'
    }

    $key = ($formatted -replace ':','').Substring(0,6).ToUpperInvariant()
    if ($script:OuiDatabase.ContainsKey($key)) {
        return [string]$script:OuiDatabase[$key]
    }

    return 'Unknown'
}

function Get-ArpTable {
    [CmdletBinding()]
    param()

    $entries = New-Object 'System.Collections.Generic.List[object]'

    if (Get-Command Get-NetNeighbor -ErrorAction SilentlyContinue) {
        try {
            foreach ($neighbor in @(Get-NetNeighbor -AddressFamily IPv4 -ErrorAction Stop)) {
                $mac = Format-MacAddress $neighbor.LinkLayerAddress
                if ($mac -and $mac -notmatch '^00:00:00:00:00:00$') {
                    [void]$entries.Add([pscustomobject]@{
                        IP = [string]$neighbor.IPAddress
                        MAC = $mac
                    })
                }
            }
        }
        catch {
        }
    }

    if (Get-Command arp.exe -ErrorAction SilentlyContinue) {
        try {
            foreach ($line in @(& arp.exe -a 2>$null)) {
                if ([string]$line -match '^\s*(\d{1,3}(?:\.\d{1,3}){3})\s+([0-9A-Fa-f:-]{17})\s+\S+\s*$') {
                    $mac = Format-MacAddress $matches[2]
                    if ($mac -and $mac -notmatch '^00:00:00:00:00:00$') {
                        [void]$entries.Add([pscustomobject]@{
                            IP = $matches[1]
                            MAC = $mac
                        })
                    }
                }
            }
        }
        catch {
        }
    }

    $hash = @{}
    foreach ($entry in $entries) {
        if (!$hash.ContainsKey($entry.IP)) {
            $hash[$entry.IP] = $entry.MAC
        }
    }

    return $hash
}

function Resolve-HostMacAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IPAddress,

        [hashtable]$ArpTable
    )

    if ($ArpTable -and $ArpTable.ContainsKey($IPAddress)) {
        return $ArpTable[$IPAddress]
    }

    if (Get-Command Get-NetNeighbor -ErrorAction SilentlyContinue) {
        try {
            $neighbor = @(
                Get-NetNeighbor -IPAddress $IPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.LinkLayerAddress } |
                Select-Object -First 1
            )

            if ($neighbor.Count) {
                $mac = Format-MacAddress $neighbor[0].LinkLayerAddress
                if ($mac -and $mac -notmatch '^00:00:00:00:00:00$') {
                    return $mac
                }
            }
        }
        catch {
        }
    }

    if (Get-Command arp.exe -ErrorAction SilentlyContinue) {
        try {
            foreach ($line in @(& arp.exe -a $IPAddress 2>$null)) {
                if ([string]$line -match "^\s*$([regex]::Escape($IPAddress))\s+([0-9A-Fa-f:-]{17})\s+\S+\s*$") {
                    $mac = Format-MacAddress $matches[1]
                    if ($mac -and $mac -notmatch '^00:00:00:00:00:00$') {
                        return $mac
                    }
                }
            }
        }
        catch {
        }
    }

    return $null
}

function Get-OsGuess {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [int]$Ttl
    )

    if ($null -eq $Ttl -or $Ttl -le 0) {
        return 'Unknown'
    }

    if ($Ttl -le 68) {
        return 'Linux / Unix-like'
    }

    if ($Ttl -le 128) {
        return 'Windows-like'
    }

    return 'Network device / Other'
}

function Get-MacAddressType {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Mac
    )

    $formatted = Format-MacAddress $Mac
    if (!$formatted) {
        return 'Unknown'
    }

    $firstOctet = [convert]::ToInt32($formatted.Substring(0,2), 16)
    $isMulticast = (($firstOctet -band 1) -eq 1)
    $isLocallyAdministered = (($firstOctet -band 2) -eq 2)

    if ($isMulticast) {
        return 'Multicast'
    }

    if ($isLocallyAdministered) {
        return 'Randomized / Local'
    }

    return 'Global'
}

function Get-ComputerManufacturerHint {
    [CmdletBinding()]
    param(
        [string]$Vendor,
        [string]$HostName
    )

    $text = "$Vendor $HostName".ToLowerInvariant()

    $patterns = @(
        @{ Regex = 'dell|precision|latitude|optiplex|inspiron|alienware|xps|vostro'; Name = 'Dell' },
        @{ Regex = 'lenovo|thinkpad|ideapad|legion|yoga'; Name = 'Lenovo' },
        @{ Regex = 'hewlett|\bhp\b|hp-|hpnbk|probook|elitebook|zbook|pavilion|compaq|prodesk|elitedesk'; Name = 'HP' },
        @{ Regex = 'asus|vivobook|zenbook|rog |tuf'; Name = 'ASUS' },
        @{ Regex = 'acer|aspire|swift|travelmate'; Name = 'Acer' },
        @{ Regex = 'msi|micro-star|prestige|stealth'; Name = 'MSI' },
        @{ Regex = 'apple|macbook|imac|mac-mini|macmini|macpro'; Name = 'Apple' },
        @{ Regex = 'microsoft|surface'; Name = 'Microsoft' },
        @{ Regex = 'samsung'; Name = 'Samsung' },
        @{ Regex = 'huawei|matebook'; Name = 'Huawei' },
        @{ Regex = 'lg electronics'; Name = 'LG' },
        @{ Regex = 'framework'; Name = 'Framework' },
        @{ Regex = 'system76'; Name = 'System76' }
    )

    foreach ($pattern in $patterns) {
        if ($text -match $pattern.Regex) {
            return $pattern.Name
        }
    }

    return $null
}

function Resolve-InferredVendor {
    [CmdletBinding()]
    param(
        [string]$MacVendor,
        [string]$HostName,
        [string]$HttpServer,
        [string]$HttpTitle,
        [int[]]$OpenPorts = @(),
        [string]$OSType
    )

    $macVendor = if ($MacVendor) { [string]$MacVendor } else { 'Unknown' }
    if ($macVendor -ne 'Unknown') {
        return [pscustomobject]@{
            Vendor = $macVendor
            Source = 'MAC OUI'
        }
    }

    $text = "$HostName $HttpServer $HttpTitle".ToLowerInvariant()
    $computerVendor = Get-ComputerManufacturerHint -Vendor '' -HostName $HostName
    if ($computerVendor) {
        return [pscustomobject]@{
            Vendor = "$computerVendor (inferred)"
            Source = 'Hostname / OEM fingerprint'
        }
    }

    $vendorRules = @(
        @{ Regex = 'hikvision|ezviz|ds-2cd|hikvisionweb'; Vendor = 'Hikvision / EZVIZ'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'dahua|imou'; Vendor = 'Dahua / IMOU'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'axis'; Vendor = 'Axis Communications'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'reolink'; Vendor = 'Reolink'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'amcrest'; Vendor = 'Amcrest'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'ubnt|unifi|ubiquiti'; Vendor = 'Ubiquiti'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'sonicwall'; Vendor = 'SonicWall'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'fortigate|fortinet'; Vendor = 'Fortinet'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'mikrotik|routeros'; Vendor = 'MikroTik'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'cisco'; Vendor = 'Cisco'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'netgear'; Vendor = 'NETGEAR'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'tp-link|tplink'; Vendor = 'TP-Link'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'd-link|dlink'; Vendor = 'D-Link'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'synology|diskstation'; Vendor = 'Synology'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'qnap|qts'; Vendor = 'QNAP'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'sonos'; Vendor = 'Sonos'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'ring'; Vendor = 'Ring'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'wyze'; Vendor = 'Wyze'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'nest'; Vendor = 'Google Nest'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'ecobee'; Vendor = 'ecobee'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'roku'; Vendor = 'Roku'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'tizen|samsung.?tv'; Vendor = 'Samsung'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'webos|lg.?tv'; Vendor = 'LG'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'bravia|sony.?tv'; Vendor = 'Sony'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'chromecast|google.?tv|google cast'; Vendor = 'Google'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'fire.?tv|amazon'; Vendor = 'Amazon'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'nvidia.?shield|shield'; Vendor = 'NVIDIA'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'xiaomi|mibox|mi box|redmi'; Vendor = 'Xiaomi'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'samsung|galaxy'; Vendor = 'Samsung'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'apple|iphone|ipad|macbook|imac|appletv'; Vendor = 'Apple'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'google|pixel|android'; Vendor = 'Google / Android OEM'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'oneplus'; Vendor = 'OnePlus'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'oppo'; Vendor = 'OPPO'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'vivo'; Vendor = 'vivo'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'motorola|moto'; Vendor = 'Motorola'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'huawei|honor'; Vendor = 'Huawei / HONOR'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'lg electronics|webos|bravia'; Vendor = 'LG / Sony'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'playstation|ps4|ps5'; Vendor = 'Sony'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'xbox'; Vendor = 'Microsoft'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'nintendo|switch'; Vendor = 'Nintendo'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'yealink'; Vendor = 'Yealink'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'grandstream'; Vendor = 'Grandstream'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'polycom|poly '; Vendor = 'Poly'; Source = 'Hostname / HTTP fingerprint' },
        @{ Regex = 'avaya'; Vendor = 'Avaya'; Source = 'Hostname / HTTP fingerprint' }
    )

    foreach ($rule in $vendorRules) {
        if ($text -match $rule.Regex) {
            return [pscustomobject]@{
                Vendor = "$($rule.Vendor) (inferred)"
                Source = $rule.Source
            }
        }
    }

    # Service-specific inference. These are deliberately conservative and are
    # marked as inferred because services can be exposed by other products.
    if (($OpenPorts -contains 8000) -and (($OpenPorts -contains 554) -or ($OpenPorts -contains 8200) -or ($text -match 'camera|cctv|nvr|dvr'))) {
        return [pscustomobject]@{
            Vendor = 'Hikvision / EZVIZ (inferred)'
            Source = 'TCP service fingerprint'
        }
    }

    if (($OpenPorts -contains 554) -and ($text -match 'camera|cam-|ipc|cctv|nvr|dvr')) {
        return [pscustomobject]@{
            Vendor = 'IP camera vendor unknown'
            Source = 'Camera service fingerprint'
        }
    }

    if ($OpenPorts -contains 62078) {
        return [pscustomobject]@{
            Vendor = 'Apple (inferred)'
            Source = 'iOS service fingerprint'
        }
    }

    if (($OpenPorts -contains 8001) -or ($OpenPorts -contains 8002)) {
        return [pscustomobject]@{
            Vendor = 'Samsung (inferred)'
            Source = 'Smart TV service fingerprint'
        }
    }

    if (($OpenPorts -contains 3000) -or ($OpenPorts -contains 3001)) {
        return [pscustomobject]@{
            Vendor = 'LG (inferred)'
            Source = 'webOS service fingerprint'
        }
    }

    if (($OpenPorts -contains 8008) -or ($OpenPorts -contains 8009) -or ($OpenPorts -contains 6466)) {
        return [pscustomobject]@{
            Vendor = 'Google / Cast ecosystem (inferred)'
            Source = 'Cast service fingerprint'
        }
    }

    if ($OpenPorts -contains 5555) {
        return [pscustomobject]@{
            Vendor = 'Android / OEM unknown'
            Source = 'ADB service fingerprint'
        }
    }

    if (($OpenPorts -contains 9100) -or ($OpenPorts -contains 631) -or ($OpenPorts -contains 515)) {
        return [pscustomobject]@{
            Vendor = 'Printer vendor unknown'
            Source = 'Printer service fingerprint'
        }
    }

    if (($OpenPorts -contains 5000) -or ($OpenPorts -contains 5001) -or ($OpenPorts -contains 2049) -or ($OpenPorts -contains 6690)) {
        return [pscustomobject]@{
            Vendor = 'NAS / Storage vendor unknown'
            Source = 'Storage service fingerprint'
        }
    }

    return [pscustomobject]@{
        Vendor = 'Unknown'
        Source = 'No vendor evidence available'
    }
}

function New-DeviceFingerprintResult {
    [CmdletBinding()]
    param(
        [string]$DeviceType = 'Unknown',
        [string]$DeviceFamily = 'Unknown',
        [string]$Manufacturer = 'Unknown',
        [AllowNull()]
        [string]$ModelHint = $null,
        [ValidateSet('Very High','High','Medium','Low')]
        [string]$Confidence = 'Low',
        [AllowNull()]
        [string]$Basis = $null
    )

    if ([string]::IsNullOrWhiteSpace($DeviceType)) {
        $DeviceType = 'Unknown'
    }

    if ([string]::IsNullOrWhiteSpace($DeviceFamily)) {
        $DeviceFamily = 'Unknown'
    }

    if ([string]::IsNullOrWhiteSpace($Manufacturer)) {
        $Manufacturer = 'Unknown'
    }

    [pscustomobject]@{
        DeviceType     = $DeviceType
        DeviceFamily   = $DeviceFamily
        Manufacturer   = $Manufacturer
        ModelHint      = $ModelHint
        Confidence     = $Confidence
        Basis          = $Basis
    }
}

function Get-DeviceFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [psobject]$Result,

        [string[]]$GatewayIPs = @()
    )

    $ip = [string]$Result.IPAddress
    $hostName = [string]$Result.HostName
    $vendor = [string]$Result.Vendor
    $osType = [string]$Result.OSType
    $httpServer = [string]$Result.HttpServer
    $httpTitle = [string]$Result.HttpTitle
    $openPorts = @($Result.OpenPortNumbers)

    $hostText = "$hostName $vendor $httpServer $httpTitle".ToLowerInvariant()
    $vendorText = $vendor.ToLowerInvariant()

    $laptopSignal = $hostText -match 'laptop|notebook|portable|thinkpad|latitude|inspiron|xps|envy|spectre|probook|elitebook|pavilion|ideapad|yoga|vivobook|zenbook|swift|aspire|travelmate|macbook|surface.laptop|hpnbk'
    $desktopSignal = $hostText -match 'desktop|optiplex|prodesk|elitedesk|workstation|tower|precision.tower|imac|mac.?mini|macmini|macpro|nuc|mini.?pc'
    $serverSignal = $hostText -match 'server|srv[-_]?\d|domain.?controller|dc[-_]?\d|fileserver|file.?server|sql[-_]?server|hyper.?v|esxi|proxmox'
    $computerServiceSignal = (
        ($openPorts -contains 135) -or
        ($openPorts -contains 139) -or
        ($openPorts -contains 445) -or
        ($openPorts -contains 3389) -or
        ($openPorts -contains 5985) -or
        ($openPorts -contains 5986) -or
        ($openPorts -contains 22)
    )

    if ($GatewayIPs -contains $ip) {
        $manufacturer = if ($vendor -and $vendor -ne 'Unknown') { $vendor } else { 'Gateway vendor unknown' }
        return New-DeviceFingerprintResult `
            -DeviceType 'Router / Gateway' `
            -DeviceFamily 'Network Infrastructure' `
            -Manufacturer $manufacturer `
            -ModelHint 'Local default gateway' `
            -Confidence 'High' `
            -Basis 'Host IP matches the local default gateway.'
    }

    # Cameras. Hikvision's common TCP/8000 management service is a strong
    # fingerprint when paired with RTSP/camera services. This check must run
    # before generic computer detection, which was the cause of the prior
    # 192.168.0.3 / 192.168.0.6 misclassification.
    $hikvisionSignal = (
        $hostText -match 'hikvision|ezviz|ds-2cd|hikvisionweb|ezcam' -or
        $vendorText -match 'hikvision|ezviz' -or
        $openPorts -contains 8000
    )
    $cameraRtspSignal = ($openPorts -contains 554)
    $cameraServiceSignal = (
        $cameraRtspSignal -or
        ($openPorts -contains 8000) -or
        ($openPorts -contains 8200) -or
        ($openPorts -contains 8899)
    )

    if ($hikvisionSignal -and $cameraServiceSignal) {
        $isNvr = $hostText -match 'nvr|dvr|recorder'
        $type = if ($isNvr) { 'Hikvision / EZVIZ NVR / DVR' } else { 'EZVIZ / Hikvision Camera' }
        $family = if ($isNvr) { 'Video Recorder' } else { 'Camera / CCTV' }
        $modelHint = if ($hostText -match 'ds-2cd') {
            'Hikvision DS-2CD family'
        }
        elseif ($cameraRtspSignal -and ($openPorts -contains 8000)) {
            'Hikvision camera, RTSP + SDK services'
        }
        else {
            'Hikvision / EZVIZ camera family'
        }

        $manufacturer = if ($vendorText -match 'ezviz' -or $hostText -match 'ezviz') {
            'EZVIZ'
        }
        elseif ($vendorText -match 'hikvision' -or $hostText -match 'hikvision') {
            'Hikvision'
        }
        else {
            'Hikvision / EZVIZ'
        }

        return New-DeviceFingerprintResult `
            -DeviceType $type `
            -DeviceFamily $family `
            -Manufacturer $manufacturer `
            -ModelHint $modelHint `
            -Confidence $(if ($openPorts -contains 8000 -and $cameraRtspSignal) { 'Very High' } else { 'High' }) `
            -Basis 'Hikvision/EZVIZ identity or TCP/8000 combined with camera service fingerprint detected.'
    }

    if ($cameraRtspSignal -and ($hostText -match 'camera|cam-|ipc|cctv|nvr|dvr|reolink|dahua|axis|amcrest')) {
        $cameraVendor = if ($vendor -and $vendor -ne 'Unknown') { $vendor } else { 'Camera vendor unknown' }
        return New-DeviceFingerprintResult `
            -DeviceType 'IP Camera / CCTV' `
            -DeviceFamily 'Camera / CCTV' `
            -Manufacturer $cameraVendor `
            -ModelHint 'RTSP network camera' `
            -Confidence 'High' `
            -Basis 'Camera hostname/vendor plus RTSP service detected.'
    }

    # Printers and multifunction devices.
    if ($hostText -match 'printer|print|epson|brother|canon|ricoh|xerox|lexmark|mfp|laserjet|deskjet' -or
        $openPorts -contains 9100 -or $openPorts -contains 631 -or $openPorts -contains 515) {
        $printerManufacturer = $null
        if ($hostText -match 'epson') { $printerManufacturer = 'Epson' }
        elseif ($hostText -match 'brother') { $printerManufacturer = 'Brother' }
        elseif ($hostText -match 'canon') { $printerManufacturer = 'Canon' }
        elseif ($hostText -match 'ricoh') { $printerManufacturer = 'Ricoh' }
        elseif ($hostText -match 'xerox') { $printerManufacturer = 'Xerox' }
        elseif ($hostText -match 'lexmark') { $printerManufacturer = 'Lexmark' }
        elseif ($vendor -and $vendor -ne 'Unknown') { $printerManufacturer = $vendor }
        else { $printerManufacturer = 'Printer vendor unknown' }

        return New-DeviceFingerprintResult `
            -DeviceType 'Printer / MFP' `
            -DeviceFamily 'Printer' `
            -Manufacturer $printerManufacturer `
            -ModelHint 'Network printer service detected' `
            -Confidence 'High' `
            -Basis 'Printer hostname or IPP/LPD/JetDirect service detected.'
    }

    # NAS / storage.
    $nasSignal = (
        $hostText -match 'nas|storage|synology|diskstation|qnap|qts|truenas|freenas|unraid' -or
        (($openPorts -contains 5000 -or $openPorts -contains 5001) -and $openPorts -contains 445) -or
        $openPorts -contains 2049 -or
        $openPorts -contains 6690
    )
    if ($nasSignal) {
        $nasManufacturer = if ($hostText -match 'synology|diskstation') { 'Synology' }
            elseif ($hostText -match 'qnap|qts') { 'QNAP' }
            elseif ($hostText -match 'truenas|freenas') { 'iXsystems / TrueNAS' }
            elseif ($hostText -match 'unraid') { 'Unraid' }
            elseif ($vendor -and $vendor -ne 'Unknown') { $vendor }
            else { 'NAS / Storage vendor unknown' }

        return New-DeviceFingerprintResult `
            -DeviceType 'NAS / Storage' `
            -DeviceFamily 'Network Storage' `
            -Manufacturer $nasManufacturer `
            -ModelHint 'Network storage service detected' `
            -Confidence $(if ($hostText -match 'synology|qnap|truenas|freenas|unraid') { 'High' } else { 'Medium' }) `
            -Basis 'NAS hostname or SMB/NFS/storage service combination detected.'
    }

    # Apple TV before generic Apple device handling.
    $appleTvSignal = (
        $hostText -match 'apple.?tv|appletv' -or
        ($vendorText -match 'apple' -and ($openPorts -contains 7000 -or $openPorts -contains 7100))
    )
    if ($appleTvSignal) {
        return New-DeviceFingerprintResult `
            -DeviceType 'Apple TV' `
            -DeviceFamily 'Media Player' `
            -Manufacturer 'Apple' `
            -ModelHint 'Apple TV / AirPlay endpoint' `
            -Confidence 'High' `
            -Basis 'Apple identity plus Apple TV hostname or AirPlay service port detected.'
    }

    # Smart TVs. Samsung documents the TV service on TCP/8001. LG webOS commonly
    # exposes its websocket services on TCP/3000 and TCP/3001. These are supporting
    # fingerprints and should not be treated as absolute hardware identification.
    $smartTvSignal = (
        $hostText -match 'tizen|samsung.?tv|webos|lg.?tv|bravia|sony.?tv' -or
        $openPorts -contains 8001 -or
        $openPorts -contains 8002 -or
        $openPorts -contains 3000 -or
        $openPorts -contains 3001
    )

    if ($smartTvSignal) {
        $tvManufacturer = 'Smart TV vendor unknown'
        $tvModel = 'Smart TV'

        if ($vendorText -match 'samsung' -or $hostText -match 'tizen|samsung.?tv' -or $openPorts -contains 8001 -or $openPorts -contains 8002) {
            $tvManufacturer = 'Samsung'
            $tvModel = 'Samsung Smart TV / Tizen family'
        }
        elseif ($vendorText -match 'lg' -or $hostText -match 'webos|lg.?tv' -or $openPorts -contains 3000 -or $openPorts -contains 3001) {
            $tvManufacturer = 'LG'
            $tvModel = 'LG Smart TV / webOS family'
        }
        elseif ($vendorText -match 'sony' -or $hostText -match 'bravia|sony.?tv') {
            $tvManufacturer = 'Sony'
            $tvModel = 'Sony BRAVIA / Smart TV family'
        }

        return New-DeviceFingerprintResult `
            -DeviceType 'Smart TV' `
            -DeviceFamily 'Television' `
            -Manufacturer $tvManufacturer `
            -ModelHint $tvModel `
            -Confidence $(if ($hostText -match 'tizen|webos|bravia|samsung.?tv|lg.?tv|sony.?tv') { 'High' } else { 'Medium' }) `
            -Basis 'Smart TV hostname/platform fingerprint or documented TV service port detected.'
    }

    # Android TV, Google TV, Chromecast, Fire TV and NVIDIA SHIELD.
    $androidTvHostnameSignal = $hostText -match 'android.?tv|google.?tv|mibox|mi.?box|redmi.?tv|tv.?box|tvbox|chromecast|fire.?tv|nvidia.?shield|shield|bravia|smart.?tv|roku'
    $androidTvServiceSignal = (
        $openPorts -contains 8008 -or
        $openPorts -contains 8009 -or
        $openPorts -contains 6466 -or
        (($openPorts -contains 5555) -and ($vendorText -match 'amlogic|rockchip|xiaomi|google|nvidia|amazon'))
    )

    if ($hostText -match 'roku') {
        return New-DeviceFingerprintResult `
            -DeviceType 'Streaming Device' `
            -DeviceFamily 'Media Player' `
            -Manufacturer 'Roku' `
            -ModelHint 'Roku streaming device family' `
            -Confidence 'High' `
            -Basis 'Roku hostname signature detected.'
    }

    if ($hostText -match 'chromecast' -or $openPorts -contains 8008 -or $openPorts -contains 8009) {
        return New-DeviceFingerprintResult `
            -DeviceType 'Streaming Device' `
            -DeviceFamily 'Media Player' `
            -Manufacturer 'Google' `
            -ModelHint 'Chromecast / Google TV family' `
            -Confidence $(if ($hostText -match 'chromecast|google.?tv') { 'High' } else { 'Medium' }) `
            -Basis 'Google Cast hostname or TCP 8008/8009 service detected.'
    }

    if ($hostText -match 'fire.?tv') {
        return New-DeviceFingerprintResult `
            -DeviceType 'Streaming Device' `
            -DeviceFamily 'Media Player' `
            -Manufacturer 'Amazon' `
            -ModelHint 'Amazon Fire TV family' `
            -Confidence 'High' `
            -Basis 'Fire TV hostname signature detected.'
    }

    if ($androidTvHostnameSignal -or $androidTvServiceSignal) {
        $manufacturer = 'Android TV / Google TV ecosystem'
        $modelHint = 'Android TV / Google TV box'

        if ($vendorText -match 'xiaomi' -or $hostText -match 'mibox|mi.?box|redmi.?tv') {
            $manufacturer = 'Xiaomi'
            $modelHint = 'Xiaomi Mi Box / Android TV family'
        }
        elseif ($vendorText -match 'nvidia' -or $hostText -match 'nvidia.?shield|shield') {
            $manufacturer = 'NVIDIA'
            $modelHint = 'NVIDIA SHIELD family'
        }
        elseif ($vendorText -match 'amazon' -or $hostText -match 'fire.?tv') {
            $manufacturer = 'Amazon'
            $modelHint = 'Amazon Fire TV family'
        }
        elseif ($vendorText -match 'google|cast' -or $hostText -match 'chromecast|google.?tv') {
            $manufacturer = 'Google'
            $modelHint = 'Chromecast / Google TV family'
        }
        elseif ($vendorText -match 'roku' -or $hostText -match 'roku') {
            $manufacturer = 'Roku'
            $modelHint = 'Roku streaming device'
        }
        elseif ($vendorText -match 'amlogic') {
            $manufacturer = 'Amlogic / OEM unknown'
        }

        return New-DeviceFingerprintResult `
            -DeviceType 'Android TV / Streaming Device' `
            -DeviceFamily 'Media Player / Smart TV' `
            -Manufacturer $manufacturer `
            -ModelHint $modelHint `
            -Confidence $(if ($androidTvHostnameSignal) { 'High' } else { 'Medium' }) `
            -Basis 'Android TV/Google Cast/streaming hostname or service fingerprint detected.'
    }

    # VoIP phones.
    if ($hostText -match 'yealink|grandstream|polycom|poly |avaya|voip|sip|ip.?phone' -or
        $openPorts -contains 5060 -or $openPorts -contains 5061) {
        $phoneManufacturer = if ($hostText -match 'yealink') { 'Yealink' }
            elseif ($hostText -match 'grandstream') { 'Grandstream' }
            elseif ($hostText -match 'polycom|poly ') { 'Poly' }
            elseif ($hostText -match 'avaya') { 'Avaya' }
            elseif ($vendor -and $vendor -ne 'Unknown') { $vendor }
            else { 'VoIP vendor unknown' }

        return New-DeviceFingerprintResult `
            -DeviceType 'VoIP / IP Phone' `
            -DeviceFamily 'Telephony' `
            -Manufacturer $phoneManufacturer `
            -ModelHint 'SIP / IP phone service detected' `
            -Confidence 'Medium' `
            -Basis 'VoIP hostname or SIP service detected.'
    }

    # Game consoles.
    if ($hostText -match 'playstation|ps4|ps5') {
        return New-DeviceFingerprintResult `
            -DeviceType 'Game Console' `
            -DeviceFamily 'Gaming' `
            -Manufacturer 'Sony' `
            -ModelHint 'PlayStation family' `
            -Confidence 'High' `
            -Basis 'PlayStation hostname/service signature detected.'
    }
    if ($hostText -match 'xbox') {
        return New-DeviceFingerprintResult `
            -DeviceType 'Game Console' `
            -DeviceFamily 'Gaming' `
            -Manufacturer 'Microsoft' `
            -ModelHint 'Xbox family' `
            -Confidence 'High' `
            -Basis 'Xbox hostname signature detected.'
    }
    if ($hostText -match 'nintendo|switch') {
        return New-DeviceFingerprintResult `
            -DeviceType 'Game Console' `
            -DeviceFamily 'Gaming' `
            -Manufacturer 'Nintendo' `
            -ModelHint 'Nintendo Switch / Nintendo family' `
            -Confidence 'High' `
            -Basis 'Nintendo hostname signature detected.'
    }

    # Generic smart-home / IoT.
    if ($hostText -match 'sonos|ring|wyze|nest|ecobee|smart.?bulb|hue|philips.?hue|tuya|tasmota|shelly|homeassistant|home.?assistant') {
        $iotManufacturer = $null
        if ($hostText -match 'sonos') { $iotManufacturer = 'Sonos' }
        elseif ($hostText -match 'ring') { $iotManufacturer = 'Ring' }
        elseif ($hostText -match 'wyze') { $iotManufacturer = 'Wyze' }
        elseif ($hostText -match 'nest') { $iotManufacturer = 'Google Nest' }
        elseif ($hostText -match 'ecobee') { $iotManufacturer = 'ecobee' }
        elseif ($hostText -match 'hue|philips') { $iotManufacturer = 'Philips Hue' }
        elseif ($hostText -match 'shelly') { $iotManufacturer = 'Shelly' }
        elseif ($hostText -match 'tasmota|tuya') { $iotManufacturer = 'Tuya / Tasmota ecosystem' }
        else { $iotManufacturer = 'IoT vendor unknown' }

        return New-DeviceFingerprintResult `
            -DeviceType 'Smart Home / IoT' `
            -DeviceFamily 'IoT' `
            -Manufacturer $iotManufacturer `
            -ModelHint 'Smart-home / IoT endpoint' `
            -Confidence 'Medium' `
            -Basis 'Known smart-home hostname/service signature detected.'
    }

    # Apple computers before Apple mobile. Apple devices with randomized MACs
    # can still be identified from hostnames or iOS/macOS services.
    $isAppleComputer = (
        $hostText -match 'macbook|imac|mac.?mini|macmini|macpro' -or
        ($vendorText -match 'apple' -and ($laptopSignal -or $desktopSignal -or $computerServiceSignal))
    )
    if ($isAppleComputer) {
        $modelHint = 'Mac computer'
        if ($hostText -match 'macbook') { $modelHint = 'MacBook family' }
        elseif ($hostText -match 'imac') { $modelHint = 'iMac family' }
        elseif ($hostText -match 'mac.?mini|macmini') { $modelHint = 'Mac mini family' }
        elseif ($hostText -match 'macpro') { $modelHint = 'Mac Pro family' }

        return New-DeviceFingerprintResult `
            -DeviceType $(if ($laptopSignal) { 'Laptop' } elseif ($desktopSignal) { 'Desktop' } else { 'Laptop / Desktop' }) `
            -DeviceFamily 'Computer' `
            -Manufacturer 'Apple' `
            -ModelHint $modelHint `
            -Confidence 'High' `
            -Basis 'Apple OUI/hostname or macOS computer fingerprint detected.'
    }

    $isAppleMobile = (
        $hostText -match 'iphone|ipad|ipod|apple.?watch' -or
        ($vendorText -match 'apple' -and $openPorts -contains 62078)
    )
    if ($isAppleMobile -or ($vendorText -match 'apple' -and !$isAppleComputer)) {
        $modelHint = 'Apple mobile / tablet device'
        if ($hostText -match 'ipad') { $modelHint = 'iPad family' }
        elseif ($hostText -match 'iphone') { $modelHint = 'iPhone family' }
        elseif ($hostText -match 'watch') { $modelHint = 'Apple Watch family' }
        elseif ($openPorts -contains 62078) { $modelHint = 'iOS/iPadOS device' }

        return New-DeviceFingerprintResult `
            -DeviceType 'Apple Device' `
            -DeviceFamily 'Mobile / Tablet / Wearable' `
            -Manufacturer 'Apple' `
            -ModelHint $modelHint `
            -Confidence $(if ($openPorts -contains 62078 -or $hostText -match 'iphone|ipad|watch') { 'High' } else { 'Medium' }) `
            -Basis 'Apple hostname/OUI or iOS service fingerprint detected.'
    }

    # Android phones/tablets and OEMs. ADB TCP/5555 is a useful secondary
    # signal, while vendor/hostname information can identify common OEMs.
    $mobileAndroidHostname = $hostText -match 'android|galaxy|pixel|oneplus|redmi|poco|mi.?phone|xiaomi|oppo|vivo|realme|moto.?g|motorola|huawei|honor|tablet|phone'
    $mobileAndroidService = ($openPorts -contains 5555 -and $hostText -notmatch 'windows|server|camera|ezviz|hikvision|tv|box|nvr')
    $androidVendorSignal = $vendorText -match 'samsung|xiaomi|huawei|google|motorola|lg electronics|htc|oneplus|oppo|vivo|realme'

    if ($mobileAndroidHostname -or $mobileAndroidService -or ($androidVendorSignal -and $osType -eq 'Linux / Unix-like')) {
        $modelHint = 'Android device'
        $androidManufacturer = $null
        if ($hostText -match 'galaxy' -or $vendorText -match 'samsung') {
            $androidManufacturer = 'Samsung'
            $modelHint = 'Samsung Galaxy / Android family'
        }
        elseif ($hostText -match 'pixel' -or $vendorText -match 'google') {
            $androidManufacturer = 'Google'
            $modelHint = 'Google Pixel / Android family'
        }
        elseif ($hostText -match 'redmi|poco|xiaomi' -or $vendorText -match 'xiaomi') {
            $androidManufacturer = 'Xiaomi'
            $modelHint = 'Xiaomi / Redmi / POCO family'
        }
        elseif ($hostText -match 'oneplus' -or $vendorText -match 'oneplus') {
            $androidManufacturer = 'OnePlus'
            $modelHint = 'OnePlus Android family'
        }
        elseif ($hostText -match 'oppo' -or $vendorText -match 'oppo') {
            $androidManufacturer = 'OPPO'
            $modelHint = 'OPPO Android family'
        }
        elseif ($hostText -match 'vivo' -or $vendorText -match 'vivo') {
            $androidManufacturer = 'vivo'
            $modelHint = 'vivo Android family'
        }
        elseif ($hostText -match 'huawei|honor' -or $vendorText -match 'huawei') {
            $androidManufacturer = 'Huawei / HONOR'
            $modelHint = 'Huawei / HONOR Android family'
        }
        elseif ($hostText -match 'motorola|moto' -or $vendorText -match 'motorola') {
            $androidManufacturer = 'Motorola'
            $modelHint = 'Motorola Android family'
        }
        else {
            $androidManufacturer = 'Android / OEM unknown'
        }

        return New-DeviceFingerprintResult `
            -DeviceType 'Android Device' `
            -DeviceFamily 'Mobile / Tablet / Embedded' `
            -Manufacturer $androidManufacturer `
            -ModelHint $modelHint `
            -Confidence $(if ($mobileAndroidHostname -or $androidVendorSignal) { 'High' } else { 'Medium' }) `
            -Basis 'Android/mobile hostname, OEM identity or ADB service detected.'
    }

    # Windows endpoint form factor. Hostname evidence takes precedence over a
    # generic TTL classification so laptops are not incorrectly shown as servers.
    if ($osType -eq 'Windows-like' -and $serverSignal) {
        $manufacturer = Get-ComputerManufacturerHint -Vendor $vendor -HostName $hostName
        if (!$manufacturer) { $manufacturer = if ($vendor -ne 'Unknown') { $vendor } else { 'PC / OEM unknown' } }

        return New-DeviceFingerprintResult `
            -DeviceType 'Windows Server' `
            -DeviceFamily 'Server / Computer' `
            -Manufacturer $manufacturer `
            -ModelHint 'Windows server endpoint' `
            -Confidence 'High' `
            -Basis 'Windows TTL plus server hostname/service fingerprint detected.'
    }

    if ($osType -eq 'Windows-like' -and ($laptopSignal -or $desktopSignal -or $computerServiceSignal)) {
        $manufacturer = Get-ComputerManufacturerHint -Vendor $vendor -HostName $hostName
        if (!$manufacturer) { $manufacturer = if ($vendor -ne 'Unknown') { $vendor } else { 'PC / OEM unknown' } }
        $deviceType = if ($laptopSignal) { 'Laptop' } elseif ($desktopSignal) { 'Desktop' } else { 'Windows Workstation' }

        return New-DeviceFingerprintResult `
            -DeviceType $deviceType `
            -DeviceFamily 'Computer' `
            -Manufacturer $manufacturer `
            -ModelHint 'Windows endpoint' `
            -Confidence $(if ($laptopSignal -or $desktopSignal) { 'High' } else { 'Medium' }) `
            -Basis 'Windows TTL plus computer hostname/OEM/service fingerprint detected.'
    }

    if ($osType -eq 'Linux / Unix-like' -and $serverSignal) {
        $manufacturer = if ($vendor -ne 'Unknown') { $vendor } else { 'Linux / OEM unknown' }
        return New-DeviceFingerprintResult `
            -DeviceType 'Linux / Unix Server' `
            -DeviceFamily 'Server / Computer' `
            -Manufacturer $manufacturer `
            -ModelHint 'Linux/Unix server endpoint' `
            -Confidence 'High' `
            -Basis 'Linux/Unix TTL plus server hostname fingerprint detected.'
    }

    if ($osType -eq 'Linux / Unix-like' -and $openPorts -contains 22) {
        $manufacturer = if ($vendor -ne 'Unknown') { $vendor } else { 'Linux / OEM unknown' }
        return New-DeviceFingerprintResult `
            -DeviceType 'Linux / Unix Computer' `
            -DeviceFamily 'Computer / Embedded' `
            -Manufacturer $manufacturer `
            -ModelHint 'Linux/Unix SSH endpoint' `
            -Confidence 'Medium' `
            -Basis 'Linux/Unix TTL plus SSH service detected.'
    }

    # Remaining infrastructure devices. Vendor identity is also considered so
    # a managed switch/AP/firewall with no discoverable HTTP/SSH service can
    # still be categorized from its OUI/vendor evidence.
    $networkVendorSignal = $vendorText -match 'ubiquiti|unifi|cisco|sonicwall|fortinet|fortigate|mikrotik|tp-link|d-link|netgear|juniper|aruba|ruckus|pfsense|opnsense|huawei'
    if ($networkVendorSignal -or
        $hostText -match 'switch|router|firewall|gateway|accesspoint|access-point|unifi|ubnt|sonicwall|forti|mikrotik|pfsense|opnsense|cisco' -or
        (($openPorts -contains 22) -and
         (($openPorts -contains 80) -or ($openPorts -contains 443)) -and
         $openPorts -notcontains 445 -and
         $openPorts -notcontains 3389)) {
        $manufacturer = if ($vendor -and $vendor -ne 'Unknown') { $vendor } else { 'Network vendor unknown' }
        return New-DeviceFingerprintResult `
            -DeviceType 'Network Device' `
            -DeviceFamily 'Network Infrastructure' `
            -Manufacturer $manufacturer `
            -ModelHint 'Network appliance / infrastructure device' `
            -Confidence 'Medium' `
            -Basis 'Network hostname or management-service pattern matched.'
    }

    # Final generic classifications. Avoid calling every TTL=128 device a
    # laptop. Without hostname/service evidence the scanner reports the
    # broader class and preserves the uncertainty in Confidence.
    if ($osType -eq 'Windows-like') {
        $manufacturer = Get-ComputerManufacturerHint -Vendor $vendor -HostName $hostName
        if (!$manufacturer) { $manufacturer = if ($vendor -ne 'Unknown') { $vendor } else { 'PC / OEM unknown' } }
        return New-DeviceFingerprintResult `
            -DeviceType 'Windows Computer' `
            -DeviceFamily 'Computer' `
            -Manufacturer $manufacturer `
            -ModelHint 'Windows endpoint, form factor unknown' `
            -Confidence 'Low' `
            -Basis 'Windows TTL heuristic detected without a more specific fingerprint.'
    }

    if ($osType -eq 'Linux / Unix-like') {
        $manufacturer = if ($vendor -ne 'Unknown') { $vendor } else { 'Linux / OEM unknown' }
        return New-DeviceFingerprintResult `
            -DeviceType 'Linux / Unix Device' `
            -DeviceFamily 'Computer / Embedded' `
            -Manufacturer $manufacturer `
            -ModelHint 'Linux/Unix endpoint, device class unknown' `
            -Confidence 'Low' `
            -Basis 'Linux/Unix TTL heuristic detected without a more specific fingerprint.'
    }

    return New-DeviceFingerprintResult
}

function Invoke-HttpFingerprint {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$IPAddress,

        [Parameter(Mandatory)]
        [int]$Port,

        [int]$TimeoutMs = 300
    )

    $result = [pscustomobject]@{
        Server = $null
        Title  = $null
    }

    if ($Port -ne 80 -and $Port -ne 8080) {
        return $result
    }

    $client = $null
    $stream = $null

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($IPAddress, $Port, $null, $null)

        if (!$async.AsyncWaitHandle.WaitOne($TimeoutMs)) {
            return $result
        }

        $client.EndConnect($async)
        $stream = $client.GetStream()
        $stream.ReadTimeout = $TimeoutMs
        $stream.WriteTimeout = $TimeoutMs

        $request = "GET / HTTP/1.0`r`nHost: $IPAddress`r`nUser-Agent: NativeNetworkScanner/1.5`r`nConnection: close`r`n`r`n"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($request)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush()

        $buffer = New-Object byte[] 8192
        $read = $stream.Read($buffer, 0, $buffer.Length)

        if ($read -le 0) {
            return $result
        }

        $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
        $headers = $text -split "`r?`n"

        foreach ($header in $headers) {
            if ($header -match '^Server\s*:\s*(.+)$') {
                $result.Server = $matches[1].Trim()
            }
        }

        if ($text -match '(?is)<title[^>]*>\s*(.*?)\s*</title>') {
            $result.Title = ($matches[1] -replace '\s+', ' ').Trim()
            if ($result.Title.Length -gt 100) {
                $result.Title = $result.Title.Substring(0,100)
            }
        }
    }
    catch {
    }
    finally {
        if ($stream) {
            try { $stream.Close() } catch {}
            try { $stream.Dispose() } catch {}
        }
        if ($client) {
            try { $client.Close() } catch {}
            try { $client.Dispose() } catch {}
        }
    }

    return $result
}

function Invoke-NetworkScan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TargetSpec,

        [Parameter(Mandatory)]
        [string]$Description,

        [string[]]$GatewayIPs = @()
    )

    $targets = @(Get-IPv4Targets -Spec $TargetSpec -MaxCount $MaxTargets)
    if (!$targets.Count) {
        throw 'No targets generated.'
    }

    Write-Host "`nScanning $Description, $($targets.Count) targets" -ForegroundColor Cyan
    Write-Host "Concurrency: $MaxConcurrency | Ping: ${PingTimeoutMs}ms | Port: ${PortTimeoutMs}ms | Ports: $($script:Ports.Count)" -ForegroundColor DarkGray

    $worker = {
        param(
            $ip,
            $ports,
            $pingTimeout,
            $portTimeout,
            $skipDns
        )

        function Get-WorkerTtlGuess {
            param([int]$Ttl)

            if ($Ttl -le 0) {
                return 'Unknown'
            }
            if ($Ttl -le 68) {
                return 'Linux / Unix-like'
            }
            if ($Ttl -le 128) {
                return 'Windows-like'
            }
            return 'Network device / Other'
        }

        function Invoke-WorkerHttpFingerprint {
            param(
                [string]$IPAddress,
                [int]$Port,
                [int]$TimeoutMs
            )

            $result = [pscustomobject]@{
                Server = $null
                Title  = $null
            }

            if ($Port -ne 80 -and $Port -ne 8080) {
                return $result
            }

            $client = $null
            $stream = $null

            try {
                $client = New-Object System.Net.Sockets.TcpClient
                $async = $client.BeginConnect($IPAddress, $Port, $null, $null)

                if (!$async.AsyncWaitHandle.WaitOne($TimeoutMs)) {
                    return $result
                }

                $client.EndConnect($async)
                $stream = $client.GetStream()
                $stream.ReadTimeout = $TimeoutMs
                $stream.WriteTimeout = $TimeoutMs

                $request = "GET / HTTP/1.0`r`nHost: $IPAddress`r`nUser-Agent: NativeNetworkScanner/1.5`r`nConnection: close`r`n`r`n"
                $bytes = [System.Text.Encoding]::ASCII.GetBytes($request)
                $stream.Write($bytes, 0, $bytes.Length)
                $stream.Flush()

                $buffer = New-Object byte[] 8192
                $read = $stream.Read($buffer, 0, $buffer.Length)

                if ($read -le 0) {
                    return $result
                }

                $text = [System.Text.Encoding]::UTF8.GetString($buffer, 0, $read)
                $lines = $text -split "`r?`n"

                foreach ($line in $lines) {
                    if ($line -match '^Server\s*:\s*(.+)$') {
                        $result.Server = $matches[1].Trim()
                    }
                }

                if ($text -match '(?is)<title[^>]*>\s*(.*?)\s*</title>') {
                    $result.Title = ($matches[1] -replace '\s+', ' ').Trim()
                    if ($result.Title.Length -gt 100) {
                        $result.Title = $result.Title.Substring(0,100)
                    }
                }
            }
            catch {
            }
            finally {
                if ($stream) {
                    try { $stream.Close() } catch {}
                    try { $stream.Dispose() } catch {}
                }
                if ($client) {
                    try { $client.Close() } catch {}
                    try { $client.Dispose() } catch {}
                }
            }

            return $result
        }

        $pingStatus = 'TimedOut'
        $rtt = $null
        $ttl = $null
        $hostName = $null
        $errorText = $null
        $portResults = New-Object 'System.Collections.Generic.List[object]'
        $httpResults = New-Object 'System.Collections.Generic.List[object]'

        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $reply = $ping.Send($ip, $pingTimeout)
            $pingStatus = [string]$reply.Status

            if ($reply.Status -eq 'Success') {
                $rtt = [math]::Round([double]$reply.RoundtripTime, 1)
                if ($reply.Options) {
                    $ttl = [int]$reply.Options.Ttl
                }
            }

            $ping.Dispose()
        }
        catch {
            $pingStatus = 'Error'
            $errorText = $_.Exception.Message
        }

        if ($pingStatus -eq 'Success' -and !$skipDns) {
            try {
                $asyncDns = [System.Net.Dns]::BeginGetHostEntry($ip, $null, $null)
                if ($asyncDns.AsyncWaitHandle.WaitOne($pingTimeout)) {
                    try {
                        $dnsResult = [System.Net.Dns]::EndGetHostEntry($asyncDns)
                        if ($dnsResult.HostName) {
                            $hostName = [string]$dnsResult.HostName
                        }
                    }
                    catch {
                    }
                }
            }
            catch {
            }
        }

        if ($pingStatus -eq 'Success') {
            foreach ($port in $ports) {
                $tcp = $null
                $state = 'Closed'

                try {
                    $tcp = New-Object System.Net.Sockets.TcpClient
                    $tcp.NoDelay = $true
                    $asyncConnect = $tcp.BeginConnect($ip, [int]$port, $null, $null)

                    if ($asyncConnect.AsyncWaitHandle.WaitOne($portTimeout)) {
                        try {
                            $tcp.EndConnect($asyncConnect)
                        }
                        catch {
                        }

                        if ($tcp.Connected) {
                            $state = 'Open'
                        }
                        else {
                            $state = 'Closed'
                        }
                    }
                    else {
                        $state = 'Filtered'
                    }
                }
                catch {
                    $state = 'Closed'
                }
                finally {
                    if ($tcp) {
                        try { $tcp.Close() } catch {}
                        try { $tcp.Dispose() } catch {}
                    }
                }

                [void]$portResults.Add([pscustomobject]@{
                    Port  = [int]$port
                    State = $state
                })
            }
        }

        $openPorts = @(
            $portResults |
            Where-Object { $_.State -eq 'Open' } |
            ForEach-Object { $_.Port } |
            Sort-Object
        )

        $filteredPorts = @(
            $portResults |
            Where-Object { $_.State -eq 'Filtered' } |
            ForEach-Object { $_.Port } |
            Sort-Object
        )

        foreach ($openHttpPort in @($openPorts | Where-Object { $_ -eq 80 -or $_ -eq 8080 })) {
            $http = Invoke-WorkerHttpFingerprint $ip $openHttpPort ([math]::Max(150, $portTimeout))
            if ($http.Server -or $http.Title) {
                [void]$httpResults.Add([pscustomobject]@{
                    Port   = [int]$openHttpPort
                    Server = $http.Server
                    Title  = $http.Title
                })
            }
        }

        $httpServer = (($httpResults | Where-Object Server | Select-Object -ExpandProperty Server -First 1) | ForEach-Object { [string]$_ })
        $httpTitle = (($httpResults | Where-Object Title | Select-Object -ExpandProperty Title -First 1) | ForEach-Object { [string]$_ })

        return [pscustomobject]@{
            IPAddress      = [string]$ip
            HostName       = $hostName
            PingStatus     = $pingStatus
            ResponseTimeMs = $rtt
            TTL            = $ttl
            OSType         = Get-WorkerTtlGuess ([int]$(if ($null -eq $ttl) { 0 } else { $ttl }))
            Open           = @($openPorts)
            Filtered       = @($filteredPorts)
            HttpServer     = $httpServer
            HttpTitle      = $httpTitle
            Error          = $errorText
        }
    }

    $pool = [runspacefactory]::CreateRunspacePool(1, $MaxConcurrency)
    $pool.ApartmentState = [Threading.ApartmentState]::MTA
    $pool.Open()

    $jobs = New-Object 'System.Collections.Generic.List[object]'
    $results = New-Object 'System.Collections.Generic.List[object]'
    $stopwatch = [Diagnostics.Stopwatch]::StartNew()

    try {
        foreach ($ip in $targets) {
            $powershell = [powershell]::Create()
            $powershell.RunspacePool = $pool

            [void]$powershell.AddScript($worker.ToString())
            [void]$powershell.AddArgument($ip)
            [void]$powershell.AddArgument([int[]]$script:Ports)
            [void]$powershell.AddArgument($PingTimeoutMs)
            [void]$powershell.AddArgument($PortTimeoutMs)
            [void]$powershell.AddArgument($script:NoDns)

            $asyncResult = $powershell.BeginInvoke()

            [void]$jobs.Add([pscustomobject]@{
                PS     = $powershell
                Async  = $asyncResult
                IP     = $ip
                Done   = $false
            })
        }

        $done = 0
        $total = $jobs.Count
        $lastPercent = -1

        while ($done -lt $total) {
            foreach ($job in @($jobs | Where-Object { $_.Async.IsCompleted -and !$_.Done })) {
                try {
                    foreach ($result in @($job.PS.EndInvoke($job.Async))) {
                        [void]$results.Add($result)
                    }
                }
                catch {
                    [void]$results.Add([pscustomobject]@{
                        IPAddress      = $job.IP
                        HostName       = $null
                        PingStatus     = 'Error'
                        ResponseTimeMs = $null
                        TTL            = $null
                        OSType         = 'Unknown'
                        Open           = @()
                        Filtered       = @()
                        HttpServer     = $null
                        HttpTitle      = $null
                        Error          = $_.Exception.Message
                    })
                }

                $job.Done = $true
                $done++

                try { $job.PS.Dispose() } catch {}
            }

            $percent = [int](($done / [double]$total) * 100)

            if ($percent -ne $lastPercent) {
                Write-Progress `
                    -Activity 'Network scan' `
                    -Status ("{0}% | {1}/{2} complete | {3} remaining | {4:N1}s" -f `
                        $percent,
                        $done,
                        $total,
                        ($total - $done),
                        $stopwatch.Elapsed.TotalSeconds) `
                    -PercentComplete $percent

                $lastPercent = $percent
            }

            if ($done -lt $total) {
                Start-Sleep -Milliseconds 35
            }
        }

        Write-Progress -Activity 'Network scan' -Completed
    }
    finally {
        foreach ($job in $jobs) {
            try {
                if (!$job.Async.IsCompleted) {
                    $job.PS.Stop()
                }
            }
            catch {
            }

            try { $job.PS.Dispose() } catch {}
        }

        try { $pool.Close() } catch {}
        try { $pool.Dispose() } catch {}
    }

    $stopwatch.Stop()

    $arp = Get-ArpTable
    $rows = New-Object 'System.Collections.Generic.List[object]'
    $active = 0

    foreach ($result in $results) {
        $mac = $null
        if ($result.PingStatus -eq 'Success') {
            $mac = Resolve-HostMacAddress -IPAddress ([string]$result.IPAddress) -ArpTable $arp
        }

        $macVendor = Resolve-MacVendor $mac
        $vendorInfo = Resolve-InferredVendor `
            -MacVendor $macVendor `
            -HostName ([string]$result.HostName) `
            -HttpServer ([string]$result.HttpServer) `
            -HttpTitle ([string]$result.HttpTitle) `
            -OpenPorts @($result.Open) `
            -OSType ([string]$result.OSType)

        $formattedOpenPorts = @(
            $result.Open | ForEach-Object {
                $number = [int]$_
                if ($script:PortNames.ContainsKey($number)) {
                    "$number/$($script:PortNames[$number])"
                }
                else {
                    [string]$number
                }
            }
        )

        $state = 'Unreachable'
        if ($result.PingStatus -eq 'Success') {
            if (@($result.Filtered).Count -gt 0 -and @($result.Open).Count -eq 0) {
                $state = 'Partial'
            }
            else {
                $state = 'Alive'
            }
        }

        if ($state -ne 'Unreachable') {
            $active++
        }

        $row = [pscustomobject]@{
            ScanTime          = Get-Date
            IPAddress         = [string]$result.IPAddress
            HostName          = $(if ($result.HostName) { [string]$result.HostName } else { $null })
            MACAddress        = $mac
            MACType           = Get-MacAddressType $mac
            Vendor            = [string]$vendorInfo.Vendor
            VendorSource      = [string]$vendorInfo.Source
            State             = $state
            ResponseTimeMs    = $result.ResponseTimeMs
            TTL               = $result.TTL
            OSType            = [string]$result.OSType
            DeviceType        = 'Unknown'
            DeviceFamily      = 'Unknown'
            Manufacturer      = $null
            ModelHint         = $null
            Confidence        = 'Low'
            DetectionBasis    = $null
            OpenPortNumbers   = @($result.Open)
            OpenPorts         = ($formattedOpenPorts -join ', ')
            FilteredPorts     = (@($result.Filtered) -join ', ')
            HttpServer        = $(if ($result.HttpServer) { [string]$result.HttpServer } else { $null })
            HttpTitle         = $(if ($result.HttpTitle) { [string]$result.HttpTitle } else { $null })
            Gateway           = $(if ($GatewayIPs -contains [string]$result.IPAddress) { 'Yes' } else { 'No' })
            Error             = $(if ($result.Error) { [string]$result.Error } else { $null })
        }

        $fingerprint = Get-DeviceFingerprint $row $GatewayIPs
        $row.DeviceType = $fingerprint.DeviceType
        $row.DeviceFamily = $fingerprint.DeviceFamily
        $row.Manufacturer = $fingerprint.Manufacturer
        $row.ModelHint = $fingerprint.ModelHint
        $row.Confidence = $fingerprint.Confidence
        $row.DetectionBasis = $fingerprint.Basis

        # Vendor is always populated when the fingerprint engine has a
        # defensible manufacturer identity. Keep the OUI result when present.
        $currentVendor = [string]$row.Vendor
        if ([string]::IsNullOrWhiteSpace($currentVendor) -or $currentVendor -eq 'Unknown') {
            $fingerprintManufacturer = [string]$fingerprint.Manufacturer
            if ($fingerprintManufacturer -and
                $fingerprintManufacturer -notmatch 'unknown|vendor unknown|OEM unknown|PC / OEM|Linux / OEM') {
                $row.Vendor = "$fingerprintManufacturer (inferred)"
                $row.VendorSource = 'Device fingerprint'
            }
            elseif ($fingerprintManufacturer) {
                $row.Vendor = $fingerprintManufacturer
            }
        }

        [void]$rows.Add($row)
    }

    $script:CurrentResults = @(
        $rows |
        Sort-Object `
            @{ Expression = { if ($_.State -eq 'Alive') { 0 } elseif ($_.State -eq 'Partial') { 1 } else { 2 } } }, `
            @{ Expression = { [uint32](Convert-IPv4ToUInt32 $_.IPAddress) } }
    )

    $script:LastTargetSpec = $TargetSpec
    $script:LastDescription = $Description
    $script:LastGatewayIPs = @($GatewayIPs)

    $script:LastStats = [pscustomobject]@{
        Target          = $Description
        Total           = $targets.Count
        Active          = $active
        Unreachable     = $targets.Count - $active
        Duration        = $stopwatch.Elapsed
        DurationText    = $stopwatch.Elapsed.ToString('hh\:mm\:ss\.fff')
        HostsPerSecond  = $(if ($stopwatch.Elapsed.TotalSeconds -gt 0) { [math]::Round($targets.Count / $stopwatch.Elapsed.TotalSeconds, 1) } else { 0 })
        CoveragePercent = 100
    }

    Show-ScanSummary
    Show-ResultsTable -Results $script:CurrentResults

    return @($script:CurrentResults)
}

function Format-Cell {
    [CmdletBinding()]
    param(
        [AllowNull()]
        [string]$Value,

        [int]$Width
    )

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    $text = $text -replace '[\r\n\t]', ' '

    if ($text.Length -gt $Width) {
        return $text.Substring(0, [math]::Max(1, $Width - 1)) + '…'
    }

    return $text.PadRight($Width)
}

function Show-ResultsTable {
    [CmdletBinding()]
    param(
        [object[]]$Results,
        [int]$Limit = 300
    )

    if (!$Results -or !$Results.Count) {
        Write-Host 'No results.' -ForegroundColor Yellow
        return
    }

    Write-Host ''
    Write-Host 'STATE      IP              HOSTNAME                    DEVICE TYPE                 VENDOR                  PORTS' -ForegroundColor White
    Write-Host '---------- --------------- --------------------------- --------------------------- ----------------------- ------------------------------' -ForegroundColor DarkGray

    $count = 0

    foreach ($result in $Results) {
        if ($count -ge $Limit) {
            break
        }

        $line = '{0} {1} {2} {3} {4} {5}' -f `
            (Format-Cell $result.State 10),
            (Format-Cell $result.IPAddress 15),
            (Format-Cell $result.HostName 27),
            (Format-Cell $result.DeviceType 27),
            (Format-Cell $result.Vendor 23),
            (Format-Cell $result.OpenPorts 30)

        $color = switch ($result.State) {
            'Alive'       { 'Green' }
            'Partial'     { 'Yellow' }
            default       { 'Red' }
        }

        Write-Host $line -ForegroundColor $color
        $count++
    }

    if ($Results.Count -gt $Limit) {
        Write-Host "Displayed $Limit of $($Results.Count) rows." -ForegroundColor DarkGray
    }
}

function Show-ScanSummary {
    [CmdletBinding()]
    param()

    if (!$script:LastStats) {
        return
    }

    $stats = $script:LastStats

    Write-Host "`n-------------------- SCAN SUMMARY --------------------" -ForegroundColor Cyan
    Write-Host "Target       : $($stats.Target)"
    Write-Host "Coverage     : $($stats.Total) addresses ($($stats.CoveragePercent)%)"
    Write-Host "Active       : $($stats.Active)" -ForegroundColor Green
    Write-Host "Unreachable  : $($stats.Unreachable)" -ForegroundColor Red
    Write-Host "Duration     : $($stats.DurationText)"
    Write-Host "Throughput   : $($stats.HostsPerSecond) hosts/sec"
    Write-Host '-------------------------------------------------------' -ForegroundColor Cyan
}

function Show-Networks {
    [CmdletBinding()]
    param()

    $networks = @(Get-LocalIPv4Networks)

    if (!$networks.Count) {
        Write-Host 'No active IPv4 network was detected.' -ForegroundColor Yellow
        return @()
    }

    Write-Host 'Detected active IPv4 networks:' -ForegroundColor Cyan

    for ($i = 0; $i -lt $networks.Count; $i++) {
        $gateway = if ($networks[$i].Gateway) { $networks[$i].Gateway } else { '-' }

        Write-Host (
            '[{0}] {1,-24} IP {2,-15} Gateway {3,-15} {4}' -f `
                ($i + 1),
                $networks[$i].InterfaceAlias,
                $networks[$i].IPAddress,
                $gateway,
                $networks[$i].CIDR
        )
    }

    return $networks
}

function Invoke-AutoScan {
    [CmdletBinding()]
    param()

    $networks = @(Show-Networks)
    if (!$networks.Count) {
        return
    }

    $selection = Read-Host 'Select network number'
    $index = 0

    if (![int]::TryParse($selection, [ref]$index) -or
        $index -lt 1 -or
        $index -gt $networks.Count) {
        Write-Host 'Invalid selection.' -ForegroundColor Yellow
        return
    }

    $network = $networks[$index - 1]
    $gatewayIPs = @()

    if ($network.Gateway) {
        $gatewayIPs += [string]$network.Gateway
    }

    try {
        Invoke-NetworkScan `
            -TargetSpec $network.CIDR `
            -Description "$($network.CIDR) [$($network.InterfaceAlias)]" `
            -GatewayIPs $gatewayIPs | Out-Null
    }
    catch {
        Write-Host "Scan failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-CustomScan {
    [CmdletBinding()]
    param()

    $spec = Read-Host 'Enter CIDR/range, for example 192.168.1.0/24 or 192.168.1.1-254'
    if (!$spec) {
        return
    }

    try {
        $count = @(Get-IPv4Targets $spec $MaxTargets).Count
        $gatewayIPs = @(
            Get-LocalIPv4Networks |
            ForEach-Object { $_.Gateway } |
            Where-Object { $_ } |
            Sort-Object -Unique
        )

        Write-Host "Validated $count targets." -ForegroundColor DarkGray

        Invoke-NetworkScan `
            -TargetSpec $spec `
            -Description $spec `
            -GatewayIPs $gatewayIPs | Out-Null
    }
    catch {
        Write-Host "Scan failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-Rescan {
    [CmdletBinding()]
    param()

    if (!$script:LastTargetSpec) {
        Write-Host 'No previous scan exists.' -ForegroundColor Yellow
        return
    }

    try {
        Invoke-NetworkScan `
            -TargetSpec $script:LastTargetSpec `
            -Description $script:LastDescription `
            -GatewayIPs $script:LastGatewayIPs | Out-Null
    }
    catch {
        Write-Host "Rescan failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-FilterSort {
    [CmdletBinding()]
    param()

    if (!$script:CurrentResults.Count) {
        Write-Host 'No results.' -ForegroundColor Yellow
        return
    }

    $fields = @(
        'State','IPAddress','HostName','MACAddress','MACType','Vendor','VendorSource',
        'DeviceType','DeviceFamily','Manufacturer','ModelHint','Confidence',
        'OSType','Gateway','OpenPorts','FilteredPorts','TTL','ResponseTimeMs',
        'HttpServer','HttpTitle'
    )

    Write-Host ('Fields: ' + ($fields -join ', ')) -ForegroundColor DarkGray

    $sortField = Read-Host 'Sort field, blank for none'
    if ($sortField -and $fields -notcontains $sortField) {
        Write-Host 'Invalid sort field.' -ForegroundColor Yellow
        return
    }

    $descending = $false
    if ($sortField) {
        $descending = (Read-Host 'Descending? y/n') -match '^[Yy]$'
    }

    $filterField = Read-Host 'Filter field, blank for none'
    $results = @($script:CurrentResults)

    if ($filterField) {
        if ($fields -notcontains $filterField) {
            Write-Host 'Invalid filter field.' -ForegroundColor Yellow
            return
        }

        $filterValue = Read-Host 'Filter value, wildcard * supported'
        $results = @(
            $results |
            Where-Object { [string]($_.$filterField) -like $filterValue }
        )
    }

    if ($sortField) {
        if ($descending) {
            $results = @($results | Sort-Object $sortField -Descending)
        }
        else {
            $results = @($results | Sort-Object $sortField)
        }
    }

    Show-ResultsTable -Results $results -Limit 500

    if ((Read-Host 'Open in Out-GridView when available? y/n') -match '^[Yy]$') {
        if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
            $results | Out-GridView -Title 'Native Network Scanner'
        }
        else {
            Write-Host 'Out-GridView is not installed. Console table remains available.' -ForegroundColor Yellow
        }
    }
}

function Show-HostDetails {
    [CmdletBinding()]
    param()

    if (!$script:CurrentResults.Count) {
        Write-Host 'No results.' -ForegroundColor Yellow
        return
    }

    $ip = Read-Host 'Host IP'
    $hostResult = @(
        $script:CurrentResults |
        Where-Object { $_.IPAddress -eq $ip } |
        Select-Object -First 1
    )

    if (!$hostResult.Count) {
        Write-Host 'Host not found.' -ForegroundColor Yellow
        return
    }

    $properties = @(
        'IPAddress','HostName','State','MACAddress','MACType','Vendor','VendorSource',
        'ResponseTimeMs','TTL','OSType','DeviceType','DeviceFamily',
        'Manufacturer','ModelHint','Confidence','DetectionBasis',
        'OpenPorts','FilteredPorts','HttpServer','HttpTitle','Gateway',
        'Error','ScanTime'
    )

    Write-Host "`n-------------------- HOST DETAILS --------------------" -ForegroundColor Cyan

    foreach ($property in $properties) {
        Write-Host ('{0,-18}: {1}' -f $property, [string]$hostResult[0].$property)
    }

    Write-Host '-------------------------------------------------------' -ForegroundColor Cyan
}

function Export-CurrentResults {
    [CmdletBinding()]
    param()

    if (!$script:CurrentResults.Count) {
        Write-Host 'No results.' -ForegroundColor Yellow
        return
    }

    $format = Read-Host 'Export C=CSV, J=JSON, B=Both'
    if ($format -notmatch '^[CcJjBb]$') {
        Write-Host 'Invalid selection.' -ForegroundColor Yellow
        return
    }

    $defaultFolder = Join-Path $env:TEMP $script:AppName

    if (!(Test-Path -LiteralPath $defaultFolder)) {
        New-Item -ItemType Directory -Path $defaultFolder -Force | Out-Null
    }

    $folderInput = Read-Host "Output folder, Enter for $defaultFolder"
    $folder = if ($folderInput) { $folderInput } else { $defaultFolder }

    try {
        if (!(Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }

        $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

        if ($format -match '^[CcBb]$') {
            $csvPath = Join-Path $folder "NativeNetworkScanner_$stamp.csv"
            $script:CurrentResults | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "CSV: $csvPath" -ForegroundColor Green
        }

        if ($format -match '^[JjBb]$') {
            $jsonPath = Join-Path $folder "NativeNetworkScanner_$stamp.json"
            $script:CurrentResults | ConvertTo-Json -Depth 8 | Set-Content -Path $jsonPath -Encoding UTF8
            Write-Host "JSON: $jsonPath" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Export failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Configure-Ports {
    [CmdletBinding()]
    param()

    Write-Host "Current ports: $($script:Ports -join ', ')" -ForegroundColor Cyan
    $inputValue = Read-Host 'New comma-separated TCP ports, Enter to keep'

    if (!$inputValue) {
        return
    }

    try {
        $newPorts = @(
            $inputValue -split ',' |
            ForEach-Object {
                $value = 0

                if (![int]::TryParse($_.Trim(), [ref]$value) -or $value -lt 1 -or $value -gt 65535) {
                    throw "Invalid port: $_"
                }

                $value
            } |
            Sort-Object -Unique
        )

        if (!$newPorts.Count) {
            throw 'At least one port is required.'
        }

        $script:Ports = $newPorts
        Write-Host "Updated ports: $($script:Ports -join ', ')" -ForegroundColor Green
    }
    catch {
        Write-Host "Port update failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Update-OuiDatabase {
    [CmdletBinding()]
    param()

    try {
        $folder = Split-Path $script:OuiCachePath -Parent

        if (!(Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }

        $response = Invoke-WebRequest `
            -Uri 'https://standards-oui.ieee.org/oui.txt' `
            -UseBasicParsing `
            -TimeoutSec 20 `
            -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace([string]$response.Content)) {
            throw 'Downloaded OUI file was empty.'
        }

        [IO.File]::WriteAllText($script:OuiCachePath, [string]$response.Content)
        $script:OuiDatabase = Import-OuiDatabase

        Write-Host "OUI entries loaded: $($script:OuiDatabase.Count)" -ForegroundColor Green
    }
    catch {
        Write-Host "OUI refresh failed. Embedded data remains active. $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Show-Menu {
    [CmdletBinding()]
    param()

    Write-Host "`n================ NATIVE NETWORK SCANNER ================" -ForegroundColor Cyan
    Write-Host '[1] Scan detected subnet'
    Write-Host '[2] Scan custom CIDR / range'
    Write-Host '[3] Rescan last target'
    Write-Host '[4] View current results'
    Write-Host '[5] Filter / sort results'
    Write-Host '[6] Detailed host information'
    Write-Host '[7] Export CSV / JSON'
    Write-Host '[8] Configure TCP ports'
    Write-Host '[9] Refresh IEEE OUI database'
    Write-Host '[0] Exit'
    Write-Host '==========================================================' -ForegroundColor Cyan
}

if ($RefreshOui) {
    Update-OuiDatabase
}

while ($true) {
    Show-Menu
    $choice = Read-Host 'Select an option'

    switch ($choice) {
        '1' {
            Invoke-AutoScan
        }
        '2' {
            Invoke-CustomScan
        }
        '3' {
            Invoke-Rescan
        }
        '4' {
            if ($script:LastStats) {
                Show-ScanSummary
            }
            Show-ResultsTable -Results $script:CurrentResults -Limit 500
        }
        '5' {
            Invoke-FilterSort
        }
        '6' {
            Show-HostDetails
        }
        '7' {
            Export-CurrentResults
        }
        '8' {
            Configure-Ports
        }
        '9' {
            Update-OuiDatabase
        }
        '0' {
            Write-Host 'Scanner exited.' -ForegroundColor Cyan
            return
        }
        default {
            Write-Host 'Invalid selection.' -ForegroundColor Yellow
        }
    }

    if ($choice -ne '0') {
        Read-Host 'Press Enter to continue' | Out-Null
    }
}


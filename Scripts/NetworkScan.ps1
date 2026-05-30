
<#
Get-NetNeighbor -AddressFamily IPv4 `
    | Where-Object {$_.State -ne 'Unreachable' -and $_.State -ne 'Permanent'} `
    | Sort-Object LinkLayerAddress -Unique `
    | Format-Table -AutoSize IPAddress, LinkLayerAddress, State
#>

<#
.SYNOPSIS
    Performs a network scan to discover active hosts and their MAC addresses.

.DESCRIPTION
    This script scans a specified IP range to identify active hosts on the network.
    It uses parallel processing for efficiency and collects both IP and MAC address
    information for each responsive host. The script is designed to work across all
    PowerShell versions (1.0 and above).

.PARAMETER IPRange
    Specifies the IP range to scan. Can be a CIDR notation (e.g., "192.168.1.0/24") 
    or an IP address with subnet mask (e.g., "192.168.1.0/255.255.255.0").

.PARAMETER Timeout
    Specifies the timeout in milliseconds for each ping attempt. Default is 500ms.

.PARAMETER Threads
    Specifies the number of concurrent threads to use. Default is 100.

.EXAMPLE
    .\NetworkScanner.ps1 -IPRange "192.168.1.0/24"
    
    Scans the entire 192.168.1.0/24 subnet with default timeout and thread settings.

.EXAMPLE
    .\NetworkScanner.ps1 -IPRange "10.0.0.0/16" -Timeout 1000 -Threads 50
    
    Scans the 10.0.0.0/16 network with a 1-second timeout and 50 concurrent threads.

.NOTES
    File Name      : NetworkScanner.ps1
    Prerequisite   : PowerShell 1.0 or later
    Version        : 1.1
    Change History : Fixed bitwise operation compatibility with PS 1.0/2.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, Position=0, HelpMessage="IP range to scan (CIDR notation, e.g. 192.168.1.0/24)")]
    [ValidateNotNullOrEmpty()]
    [string]$IPRange,
    
    [Parameter(Mandatory=$false, HelpMessage="Timeout in milliseconds for each ping")]
    [ValidateRange(100, 10000)]
    [int]$Timeout = 500,
    
    [Parameter(Mandatory=$false, HelpMessage="Number of concurrent threads")]
    [ValidateRange(1, 500)]
    [int]$Threads = 100
)

#region Helper Functions

function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("Info", "Warning", "Error", "Success", "Verbose")]
        [string]$Level = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        "Info"     { Write-Host $logMessage -ForegroundColor Cyan }
        "Warning"  { Write-Warning $logMessage }
        "Error"    { Write-Error $logMessage }
        "Success"  { Write-Host $logMessage -ForegroundColor Green }
        "Verbose"  { Write-Verbose $logMessage }
    }
}

# Cross-version compatible bitwise operations
function Get-BitwiseAnd {
    param([UInt32]$Left, [UInt32]$Right)
    
    # Manual bitwise AND implementation that works on all PS versions
    $result = 0
    for ($i = 0; $i -lt 32; $i++) {
        $leftBit = [Math]::Floor($Left / [Math]::Pow(2, $i)) % 2
        $rightBit = [Math]::Floor($Right / [Math]::Pow(2, $i)) % 2
        $resultBit = $leftBit -band $rightBit
        $result += $resultBit * [Math]::Pow(2, $i)
    }
    
    return [UInt32]$result
}

function Get-BitwiseShiftLeft {
    param([UInt32]$Value, [int]$Shift)
    
    # Manual bitwise shift left implementation that works on all PS versions
    if ($Shift -le 0) { return $Value }
    if ($Shift -ge 32) { return 0 }
    
    return [UInt32]($Value * [Math]::Pow(2, $Shift))
}

function Get-BitwiseShiftRight {
    param([UInt32]$Value, [int]$Shift)
    
    # Manual bitwise shift right implementation that works on all PS versions
    if ($Shift -le 0) { return $Value }
    if ($Shift -ge 32) { return 0 }
    
    return [UInt32]([Math]::Floor($Value / [Math]::Pow(2, $Shift)))
}

function Get-BitwiseComplement {
    param([UInt32]$Value)
    
    # Manual bitwise complement (~) implementation for all PS versions
    $result = 0
    for ($i = 0; $i -lt 32; $i++) {
        $bit = [Math]::Floor($Value / [Math]::Pow(2, $i)) % 2
        $complementBit = 1 - $bit
        $result += $complementBit * [Math]::Pow(2, $i)
    }
    
    return [UInt32]$result
}

function Get-ParsedCIDR {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CIDRNotation
    )
    try {
        # Ensure we have a single, trimmed string
        if ($CIDRNotation -is [array]) {
            Write-LogMessage -Level Warning -Message "Multiple CIDR values detected; using the first one."
            $CIDRNotation = $CIDRNotation[0]
        }
        $CIDRNotation = $CIDRNotation.Trim()

        # Updated regex to match an IP and CIDR bits
        if ($CIDRNotation -match '^(\d{1,3}(?:\.\d{1,3}){3})/(\d{1,2})$') {
            $baseIP = $matches[1]
            $cidrBits = [int]$matches[2]

            # Split the base IP on period and ensure exactly 4 octets
            $octets = $baseIP -split '\.'
            if ($octets.Count -ne 4) {
                throw "Invalid IP address format: $baseIP"
            }

            # Validate each octet is between 0 and 255
            foreach ($octet in $octets) {
                $octetValue = [int]$octet
                if ($octetValue -lt 0 -or $octetValue -gt 255) {
                    throw "Invalid IP address: $baseIP"
                }
            }

            if ($cidrBits -lt 0 -or $cidrBits -gt 32) {
                throw "Invalid CIDR bits: $cidrBits. Must be between 0 and 32."
            }

            # Calculate the subnet size and mask value
            $subnetSize = [UInt32][Math]::Pow(2, (32 - $cidrBits))
            $maskValue = [UInt32]0
            if ($cidrBits -gt 0) {
                $maskValue = [UInt32]([Math]::Pow(2, 32) - [Math]::Pow(2, (32 - $cidrBits)))
            }

            # Convert IP to UInt32 using proper shifting
            $ipUint32 = [UInt32]0
            for ($i = 0; $i -lt 4; $i++) {
                $octetValue = [UInt32]([int]$octets[$i])
                $shiftAmount = 24 - ($i * 8)
                $ipUint32 += Get-BitwiseShiftLeft -Value $octetValue -Shift $shiftAmount
            }

            # Calculate network address
            $networkUint32 = Get-BitwiseAnd -Left $ipUint32 -Right $maskValue

            # Determine first and last usable addresses (if applicable)
            $firstUsableUint32 = $networkUint32 + 1
            $lastUsableUint32 = $networkUint32 + $subnetSize - 2

            # Convert the first and last IP addresses to dotted-quad strings
            $firstIP = @(
                (Get-BitwiseShiftRight -Value $firstUsableUint32 -Shift 24) % 256,
                (Get-BitwiseShiftRight -Value $firstUsableUint32 -Shift 16) % 256,
                (Get-BitwiseShiftRight -Value $firstUsableUint32 -Shift 8) % 256,
                $firstUsableUint32 % 256
            ) -join '.'

            $lastIP = @(
                (Get-BitwiseShiftRight -Value $lastUsableUint32 -Shift 24) % 256,
                (Get-BitwiseShiftRight -Value $lastUsableUint32 -Shift 16) % 256,
                (Get-BitwiseShiftRight -Value $lastUsableUint32 -Shift 8) % 256,
                $lastUsableUint32 % 256
            ) -join '.'

            return @{
                FirstIP       = $firstIP
                LastIP        = $lastIP
                TotalHosts    = $subnetSize - 2  # Exclude network and broadcast addresses
                NetworkAddress = $baseIP
                CIDRBits      = $cidrBits
            }
        }
        else {
            throw "Invalid CIDR notation. Expected format: x.x.x.x/y"
        }
    }
    catch {
        Write-LogMessage -Level Error -Message "Error parsing CIDR notation: $_"
        throw
    }
}

function Get-IPRange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$StartIP,
        
        [Parameter(Mandatory=$true)]
        [string]$EndIP
    )
    
    try {
        # Convert IP addresses to integers for comparison using a compatible method
        $startIPInt = ConvertTo-IPInt -IPAddress $StartIP
        $endIPInt = ConvertTo-IPInt -IPAddress $EndIP
        
        # Validate range
        if ($endIPInt -lt $startIPInt) {
            throw "End IP address must be greater than or equal to start IP address."
        }
        
        # Generate IP addresses in the range
        $ipRange = @()
        for ($i = $startIPInt; $i -le $endIPInt; $i++) {
            $ipRange += ConvertFrom-IPInt -IPInt $i
        }
        
        return $ipRange
    }
    catch {
        Write-LogMessage -Level Error -Message "Error generating IP range: $_"
        throw
    }
}

# Convert IP address to integer - compatible with all PS versions
function ConvertTo-IPInt {
    param(
        [Parameter(Mandatory = $true)]
        [string]$IPAddress
    )
    $IPAddress = $IPAddress.Trim()
    $octets = $IPAddress -split '\.'
    if ($octets.Count -ne 4) {
        throw "Invalid IP address format: $IPAddress"
    }
    $ipInt = [UInt32]0
    for ($i = 0; $i -lt 4; $i++) {
        $ipInt += [UInt32]([int]$octets[$i]) * [Math]::Pow(256, (3 - $i))
    }
    return $ipInt
}
# Convert integer to IP address - compatible with all PS versions
function ConvertFrom-IPInt {
    param(
        [Parameter(Mandatory = $true)]
        [UInt32]$IPInt
    )
    $octets = @(
        [Math]::Floor($IPInt / [Math]::Pow(256, 3)) % 256,
        [Math]::Floor($IPInt / [Math]::Pow(256, 2)) % 256,
        [Math]::Floor($IPInt / [Math]::Pow(256, 1)) % 256,
        [Math]::Floor($IPInt / [Math]::Pow(256, 0)) % 256
    )
    return $octets -join '.'
}
function Test-HostOnline {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$IPAddress,
        
        [Parameter(Mandatory=$false)]
        [int]$Timeout = 500
    )
    
    try {
        # First try using System.Net.NetworkInformation.Ping - available on PS 2.0+
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $result = $ping.Send($IPAddress, $Timeout)
            return ($null -ne $result -and $result.Status -eq 'Success')
        }
        catch [System.Management.Automation.MethodInvocationException] {
            # Fallback to legacy ping command
            Write-Verbose "Falling back to legacy ping command for $IPAddress"
            $pingResult = ping -n 1 -w $Timeout $IPAddress 2>$null
            return ($pingResult -match 'Reply from')
        }
    }
    catch {
        Write-Verbose "Error pinging ${IPAddress}: $_"
        return $false
    }
}

function Get-MACAddress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$IPAddress
    )
    
    try {
        # Try multiple methods to get MAC address
        
        # Method 1: Get-NetNeighbor (PowerShell 3.0+ on Windows 8/Server 2012+)
        if (Get-Command -Name Get-NetNeighbor -ErrorAction SilentlyContinue) {
            $neighbor = Get-NetNeighbor -IPAddress $IPAddress -ErrorAction SilentlyContinue | 
                       Select-Object -First 1
            if ($neighbor -and $neighbor.LinkLayerAddress) {
                return $neighbor.LinkLayerAddress, "Reachable"
            }
        }
        
        # Method 2: Parse ARP table
        $arpResult = & arp -a $IPAddress 2>$null
        
        if ($arpResult -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+(([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2}))') {
            $macAddress = $matches[2]
            # Standardize MAC address format
            $macAddress = ($macAddress -replace '-', ':').ToUpper()
            return $macAddress, "Reachable"
        }
        
        # If we're here, we can ping the host but couldn't get MAC
        return "Unknown", "Reachable"
    }
    catch {
        Write-Verbose "Error getting MAC address for $IPAddress`: $_"
        return "Unknown", "Unknown"
    }
}

#endregion

#region Main Execution

# Validate PowerShell version and load required assemblies
try {
    # PowerShell 2.0+ has access to runspace functionality out of the box
    $useRunspaces = $false
    
    try {
        $psVersion = $PSVersionTable.PSVersion.Major
        
        if ($psVersion -ge 2) {
            Write-LogMessage -Level Verbose -Message "PowerShell version $psVersion detected. Using built-in threading capabilities."
            $useRunspaces = $true
        }
        else {
            Write-LogMessage -Level Warning -Message "PowerShell version 1.0 detected. Using sequential scanning (slower)."
            $useRunspaces = $false
        }
    }
    catch {
        # On very old PowerShell versions, PSVersionTable might not exist
        Write-LogMessage -Level Warning -Message "PowerShell version detection failed. Using sequential scanning."
        $useRunspaces = $false
    }
}
catch {
    Write-LogMessage -Level Warning -Message "Unable to determine PowerShell version. Using sequential scanning."
    $useRunspaces = $false
}

# Parse IP range
try {
    $networkInfo = Get-ParsedCIDR -CIDRNotation $IPRange
    Write-LogMessage -Level Info -Message "Scanning network: $($networkInfo.NetworkAddress)/$($networkInfo.CIDRBits)"
    Write-LogMessage -Level Info -Message "IP Range: $($networkInfo.FirstIP) to $($networkInfo.LastIP)"
    Write-LogMessage -Level Info -Message "Total hosts to scan: $($networkInfo.TotalHosts)"
    
    $allIPs = Get-IPRange -StartIP $networkInfo.FirstIP -EndIP $networkInfo.LastIP
}
catch {
    Write-LogMessage -Level Error -Message "Failed to parse IP range: $_"
    exit 1
}

# Initialize results array
$results = @()

# Display scan parameters
Write-LogMessage -Level Info -Message "Starting network scan with timeout of ${Timeout}ms and $Threads max concurrent operations"
$startTime = Get-Date

# Scan IP addresses
if ($useRunspaces) {
    # Test if runspaces are available
    try {
        # Create runspace pool for parallel processing
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $Threads)
        $runspacePool.Open()
        
        $runspaces = @()
        $scriptBlock = {
            param($ip, $timeout)
            
            if (Test-HostOnline -IPAddress $ip -Timeout $timeout) {
                $macInfo = Get-MACAddress -IPAddress $ip
                
                [PSCustomObject]@{
                    IPAddress = $ip
                    LinkLayerAddress = $macInfo[0]
                    State = $macInfo[1]
                    IsOnline = $true
                }
            }
            else {
                [PSCustomObject]@{
                    IPAddress = $ip
                    LinkLayerAddress = $null
                    State = "Unreachable"
                    IsOnline = $false
                }
            }
        }
        
        # Create and start runspaces for each IP
        foreach ($ip in $allIPs) {
            $powerShell = [powershell]::Create().AddScript($scriptBlock).AddParameter("ip", $ip).AddParameter("timeout", $Timeout)
            $powerShell.RunspacePool = $runspacePool
            
            $runspaces += [PSCustomObject]@{
                PowerShell = $powerShell
                Runspace = $powerShell.BeginInvoke()
                IP = $ip
            }
        }
        
        # Poll for completion and collect results
        $completed = 0
        $total = $allIPs.Count
        
        do {
            $stillRunning = $false
            
            foreach ($runspace in $runspaces | Where-Object { $null -ne $_.Runspace }) {
                if ($runspace.Runspace.IsCompleted) {
                    $result = $runspace.PowerShell.EndInvoke($runspace.Runspace)
                    if ($result -and $result.IsOnline) {
                        $results += $result
                    }
                    
                    $runspace.PowerShell.Dispose()
                    $runspace.Runspace = $null
                    $runspace.PowerShell = $null
                    
                    $completed++
                    if ($completed % 10 -eq 0 -or $completed -eq $total) {
                        Write-Progress -Activity "Scanning Network" -Status "Scanning IPs" -PercentComplete (($completed / $total) * 100) -CurrentOperation "Completed $completed of $total"
                    }
                }
                else {
                    $stillRunning = $true
                }
            }
            
            if ($stillRunning) {
                Start-Sleep -Milliseconds 100
            }
        } while ($stillRunning)
        
        # Cleanup
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
    catch {
        Write-LogMessage -Level Warning -Message "Parallel processing failed: $_. Falling back to sequential scanning."
        $useRunspaces = $false
        
        # If we failed to use runspaces, we'll fall through to the sequential scanning code below
    }
}

# Sequential scanning for PowerShell 1.0 or if parallel processing failed
if (-not $useRunspaces) {
    $total = $allIPs.Count
    $completed = 0
    
    foreach ($ip in $allIPs) {
        Write-Progress -Activity "Scanning Network" -Status "Scanning IPs" -PercentComplete (($completed / $total) * 100) -CurrentOperation "Scanning $ip"
        
        if (Test-HostOnline -IPAddress $ip -Timeout $Timeout) {
            $macInfo = Get-MACAddress -IPAddress $ip
            
            $results += [PSCustomObject]@{
                IPAddress = $ip
                LinkLayerAddress = $macInfo[0]
                State = $macInfo[1]
                IsOnline = $true
            }
        }
        
        $completed++
    }
}

# Calculate scan duration
$endTime = Get-Date
$duration = $endTime - $startTime
$formattedDuration = "{0:00}:{1:00}:{2:00}" -f $duration.Hours, $duration.Minutes, $duration.Seconds

# Filter out unreachable hosts and format results
$finalResults = $results | Where-Object { $_.IsOnline } | Select-Object IPAddress, LinkLayerAddress, State

# Display results
Write-LogMessage -Level Success -Message "Scan completed in $formattedDuration. Found $($finalResults.Count) active hosts."
$finalResults | Format-Table -AutoSize

#endregion
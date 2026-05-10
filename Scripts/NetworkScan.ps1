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
    Version        : 1.2
    Change History : Fixed variable scope issues and improved error handling
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
    
    if ($Shift -le 0) { return $Value }
    if ($Shift -ge 32) { return 0 }
    
    return [UInt32]($Value * [Math]::Pow(2, $Shift))
}

function Get-BitwiseShiftRight {
    param([UInt32]$Value, [int]$Shift)
    
    if ($Shift -le 0) { return $Value }
    if ($Shift -ge 32) { return 0 }
    
    return [UInt32]([Math]::Floor($Value / [Math]::Pow(2, $Shift)))
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

            # Split the base IP and validate
            $octets = $baseIP -split '\.'
            if ($octets.Count -ne 4) {
                throw "Invalid IP address format: $baseIP"
            }

            # Validate each octet
            foreach ($octet in $octets) {
                $octetValue = [int]$octet
                if ($octetValue -lt 0 -or $octetValue -gt 255) {
                    throw "Invalid IP address: $baseIP"
                }
            }

            if ($cidrBits -lt 0 -or $cidrBits -gt 32) {
                throw "Invalid CIDR bits: $cidrBits. Must be between 0 and 32."
            }

            # Calculate subnet size and mask
            $subnetSize = [UInt32][Math]::Pow(2, (32 - $cidrBits))
            $maskValue = [UInt32]0
            if ($cidrBits -gt 0) {
                $maskValue = [UInt32]([Math]::Pow(2, 32) - [Math]::Pow(2, (32 - $cidrBits)))
            }

            # Convert IP to UInt32
            $ipUint32 = [UInt32]0
            for ($i = 0; $i -lt 4; $i++) {
                $octetValue = [UInt32]([int]$octets[$i])
                $shiftAmount = 24 - ($i * 8)
                $ipUint32 += Get-BitwiseShiftLeft -Value $octetValue -Shift $shiftAmount
            }

            # Calculate network address
            $networkUint32 = Get-BitwiseAnd -Left $ipUint32 -Right $maskValue

            # Determine first and last usable addresses
            $firstUsableUint32 = $networkUint32 + 1
            $lastUsableUint32 = $networkUint32 + $subnetSize - 2

            # Convert to IP strings
            $firstIP = @(
                (Get-BitwiseShiftRight -Value $firstUsableUint32 -Shift 24) -band 255,
                (Get-BitwiseShiftRight -Value $firstUsableUint32 -Shift 16) -band 255,
                (Get-BitwiseShiftRight -Value $firstUsableUint32 -Shift 8) -band 255,
                $firstUsableUint32 -band 255
            ) -join '.'

            $lastIP = @(
                (Get-BitwiseShiftRight -Value $lastUsableUint32 -Shift 24) -band 255,
                (Get-BitwiseShiftRight -Value $lastUsableUint32 -Shift 16) -band 255,
                (Get-BitwiseShiftRight -Value $lastUsableUint32 -Shift 8) -band 255,
                $lastUsableUint32 -band 255
            ) -join '.'

            return @{
                FirstIP       = $firstIP
                LastIP        = $lastIP
                TotalHosts    = $subnetSize - 2
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
        $startIPInt = ConvertTo-IPInt -IPAddress $StartIP
        $endIPInt = ConvertTo-IPInt -IPAddress $EndIP
        
        if ($endIPInt -lt $startIPInt) {
            throw "End IP address must be greater than or equal to start IP address."
        }
        
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
    
    [UInt32]$ipInt = 0
    for ($i = 0; $i -lt 4; $i++) {
        $ipInt += [UInt32]([int]$octets[$i]) * [Math]::Pow(256, (3 - $i))
    }
    return $ipInt
}

function ConvertFrom-IPInt {
    param(
        [Parameter(Mandatory = $true)]
        [UInt32]$IPInt
    )
    
    $octets = @(
        [int]([Math]::Floor($IPInt / 16777216) % 256),
        [int]([Math]::Floor($IPInt / 65536) % 256),
        [int]([Math]::Floor($IPInt / 256) % 256),
        [int]($IPInt % 256)
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
        # Try System.Net.NetworkInformation.Ping first
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $result = $ping.Send($IPAddress, $Timeout)
            $ping.Dispose()
            return ($null -ne $result -and $result.Status.ToString() -eq 'Success')
        }
        catch {
            # Fallback to legacy ping command
            Write-Verbose "Falling back to legacy ping command for $IPAddress"
            $pingResult = & ping -n 1 -w $Timeout $IPAddress 2>$null
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
        # Try Get-NetNeighbor first (PowerShell 3.0+)
        if (Get-Command -Name Get-NetNeighbor -ErrorAction SilentlyContinue) {
            try {
                $neighbor = Get-NetNeighbor -IPAddress $IPAddress -ErrorAction SilentlyContinue | 
                           Select-Object -First 1
                if ($neighbor -and $neighbor.LinkLayerAddress -and $neighbor.LinkLayerAddress -ne '00-00-00-00-00-00') {
                    return $neighbor.LinkLayerAddress, $neighbor.State
                }
            }
            catch {
                Write-Verbose "Get-NetNeighbor failed for $IPAddress`: $_"
            }
        }
        
        # Parse ARP table as fallback
        $arpResult = & arp -a $IPAddress 2>$null
        
        if ($arpResult -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+(([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2}))') {
            $macAddress = $matches[2]
            $macAddress = ($macAddress -replace '-', ':').ToUpper()
            return $macAddress, "Reachable"
        }
        
        return "Unknown", "Reachable"
    }
    catch {
        Write-Verbose "Error getting MAC address for $IPAddress`: $_"
        return "Unknown", "Unknown"
    }
}

#endregion

#region Main Execution

# Initialize variables
$useRunspaces = $false
$results = @()

# Determine PowerShell capabilities
try {
    if ($PSVersionTable -and $PSVersionTable.PSVersion.Major -ge 2) {
        Write-LogMessage -Level Verbose -Message "PowerShell version $($PSVersionTable.PSVersion.Major) detected. Using parallel processing."
        $useRunspaces = $true
    }
    else {
        Write-LogMessage -Level Warning -Message "PowerShell version 1.0 detected. Using sequential scanning."
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

# Display scan parameters
Write-LogMessage -Level Info -Message "Starting network scan with timeout of ${Timeout}ms and $Threads max concurrent operations"
$startTime = Get-Date

# Main scanning logic
if ($useRunspaces) {
    try {
        # Create runspace pool
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $Threads)
        $runspacePool.Open()
        
        # Define the script block with all required functions
        $scriptBlock = {
            param($ip, $timeout)
            
            # Embed required functions in scriptblock
            function Test-HostOnline {
                param([string]$IPAddress, [int]$Timeout = 500)
                try {
                    $ping = New-Object System.Net.NetworkInformation.Ping
                    $result = $ping.Send($IPAddress, $Timeout)
                    $ping.Dispose()
                    return ($null -ne $result -and $result.Status.ToString() -eq 'Success')
                }
                catch {
                    $pingResult = & ping -n 1 -w $Timeout $IPAddress 2>$null
                    return ($pingResult -match 'Reply from')
                }
            }
            
            function Get-MACAddress {
                param([string]$IPAddress)
                try {
                    if (Get-Command -Name Get-NetNeighbor -ErrorAction SilentlyContinue) {
                        $neighbor = Get-NetNeighbor -IPAddress $IPAddress -ErrorAction SilentlyContinue | Select-Object -First 1
                        if ($neighbor -and $neighbor.LinkLayerAddress -and $neighbor.LinkLayerAddress -ne '00-00-00-00-00-00') {
                            return $neighbor.LinkLayerAddress, $neighbor.State
                        }
                    }
                    
                    $arpResult = & arp -a $IPAddress 2>$null
                    if ($arpResult -match '(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\s+(([0-9A-Fa-f]{2}[:-]){5}([0-9A-Fa-f]{2}))') {
                        $macAddress = $matches[2]
                        $macAddress = ($macAddress -replace '-', ':').ToUpper()
                        return $macAddress, "Reachable"
                    }
                    
                    return "Unknown", "Reachable"
                }
                catch {
                    return "Unknown", "Unknown"
                }
            }
            
            # Main logic
            if (Test-HostOnline -IPAddress $ip -Timeout $timeout) {
                $macInfo = Get-MACAddress -IPAddress $ip
                
                [PSCustomObject]@{
                    IPAddress = $ip
                    LinkLayerAddress = $macInfo[0]
                    State = $macInfo[1]
                    IsOnline = $true
                }
            }
        }
        
        $runspaces = @()
        
        # Create runspaces for each IP
        foreach ($ip in $allIPs) {
            $powerShell = [powershell]::Create().AddScript($scriptBlock).AddParameter("ip", $ip).AddParameter("timeout", $Timeout)
            $powerShell.RunspacePool = $runspacePool
            
            $runspaces += [PSCustomObject]@{
                PowerShell = $powerShell
                Runspace = $powerShell.BeginInvoke()
                IP = $ip
            }
        }
        
        # Collect results
        $completed = 0
        $total = $allIPs.Count
        
        do {
            $stillRunning = $false
            
            for ($i = 0; $i -lt $runspaces.Count; $i++) {
                if ($null -ne $runspaces[$i].Runspace -and $runspaces[$i].Runspace.IsCompleted) {
                    $result = $runspaces[$i].PowerShell.EndInvoke($runspaces[$i].Runspace)
                    if ($result -and $result.IsOnline) {
                        $results += $result
                    }
                    
                    $runspaces[$i].PowerShell.Dispose()
                    $runspaces[$i].Runspace = $null
                    $runspaces[$i].PowerShell = $null
                    
                    $completed++
                    if ($completed % 10 -eq 0 -or $completed -eq $total) {
                        Write-Progress -Activity "Scanning Network" -Status "Scanning IPs" -PercentComplete (($completed / $total) * 100) -CurrentOperation "Completed $completed of $total"
                    }
                }
                elseif ($null -ne $runspaces[$i].Runspace) {
                    $stillRunning = $true
                }
            }
            
            if ($stillRunning) {
                Start-Sleep -Milliseconds 100
            }
        } while ($stillRunning)
        
        $runspacePool.Close()
        $runspacePool.Dispose()
    }
    catch {
        Write-LogMessage -Level Warning -Message "Parallel processing failed: $_. Falling back to sequential scanning."
        $useRunspaces = $false
        $results = @()  # Reset results array
    }
}

# Sequential scanning fallback
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

# Calculate scan duration and display results
$endTime = Get-Date
$duration = $endTime - $startTime
$formattedDuration = "{0:00}:{1:00}:{2:00}" -f $duration.Hours, $duration.Minutes, $duration.Seconds

$finalResults = $results | Where-Object { $_.IsOnline } | Select-Object IPAddress, LinkLayerAddress, State

Write-Progress -Activity "Scanning Network" -Completed
Write-LogMessage -Level Success -Message "Scan completed in $formattedDuration. Found $($finalResults.Count) active hosts."

if ($finalResults.Count -gt 0) {
    $finalResults | Format-Table -AutoSize
}
else {
    Write-LogMessage -Level Warning -Message "No active hosts found in the specified range."
}

#endregion

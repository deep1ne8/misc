# Pervasive Database Connection Troubleshooter
# This script diagnoses and attempts to fix Pervasive database connection issues (Error 3110)
# Created: March 7, 2025

# Configuration Variables - Modify these for your environment
$PervasiveServerName = "MY-Apps.digney.local" # Replace with your actual Pervasive database server
$PervasivePort = 1583 # Default Pervasive port, change if different
$PervasiveServiceNames = @("PSQL", "Pervasive.SQL", "Pervasive PSQL Workgroup Engine", "Pervasive PSQL Workgroup Engine")
$LogPath = "$env:SYSTEMROOT\TEMP\PervasiveTroubleshooter.log"

# Function to write to both console and log file
function Write-Log {
    param (
        [string]$Message,
        [string]$Type = "INFO" # INFO, ERROR, SUCCESS, WARNING
    )
    
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$TimeStamp] [$Type] $Message"
    
    # Output to console with color coding
    switch ($Type) {
        "ERROR" { Write-Host $LogMessage -ForegroundColor Red }
        "SUCCESS" { Write-Host $LogMessage -ForegroundColor Green }
        "WARNING" { Write-Host $LogMessage -ForegroundColor Yellow }
        "INFO" { Write-Host $LogMessage -ForegroundColor Cyan }
        default { Write-Host $LogMessage }
    }
    
    # Append to log file
    Add-Content -Path $LogPath -Value $LogMessage
}

# Start with a clean log file
if (Test-Path $LogPath) {
    Remove-Item $LogPath -Force
}

Write-Log "Starting Pervasive database connection troubleshooting script" "INFO"
Write-Log "Target server: $PervasiveServerName" "INFO"

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log "Script is not running with administrative privileges. Some operations may fail." "WARNING"
} else {
    Write-Log "Running with administrative privileges" "INFO"
}

# Step 1: Test basic network connectivity to the server
Write-Log "Testing network connectivity to $PervasiveServerName..." "INFO"
try {
    $pingResult = Test-Connection -ComputerName $PervasiveServerName -Count 2 -ErrorAction Stop
    Write-Log "Network connectivity to $PervasiveServerName is available (Latency: $($pingResult[0].ResponseTime) ms)" "SUCCESS"
} catch {
    Write-Log "Cannot reach $PervasiveServerName - Network connectivity issue: $($_.Exception.Message)" "ERROR"
    Write-Log "Recommendations:" "INFO"
    Write-Log "  - Verify network cable connections" "INFO"
    Write-Log "  - Check if the server is online" "INFO"
    Write-Log "  - Verify VPN connection if connecting remotely" "INFO"
    Write-Log "  - Check firewall settings" "INFO"
}

# Step 2: Test port connectivity for Pervasive
Write-Log "Testing Pervasive port connectivity on port $PervasivePort..." "INFO"
$portTest = New-Object System.Net.Sockets.TcpClient
try {
    $portTestResult = $portTest.BeginConnect($PervasiveServerName, $PervasivePort, $null, $null)
    $asyncResult = $portTestResult.AsyncWaitHandle.WaitOne(1000, $false)
    if ($asyncResult) {
        Write-Log "Port $PervasivePort is open on $PervasiveServerName" "SUCCESS"
    } else {
        Write-Log "Port $PervasivePort is closed or filtered on $PervasiveServerName" "ERROR"
        Write-Log "Recommendations:" "INFO"
        Write-Log "  - Check if Pervasive service is running on the server" "INFO"
        Write-Log "  - Verify firewall allows traffic on port $PervasivePort" "INFO"
    }
} catch {
    Write-Log "Error testing port connectivity: $($_.Exception.Message)" "ERROR"
} finally {
    $portTest.Close()
}

# Step 3: Check local Pervasive client services
Write-Log "Checking local Pervasive client services..." "INFO"
$foundPervasiveService = $false

foreach ($serviceName in $PervasiveServiceNames) {
    try {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            $foundPervasiveService = $true
            Write-Log "Found Pervasive service: $($service.DisplayName) - Status: $($service.Status)" "INFO"
            
            if ($service.Status -ne "Running") {
                Write-Log "Attempting to start $($service.DisplayName)..." "WARNING"
                try {
                    Start-Service -Name $serviceName -ErrorAction Stop
                    $updatedService = Get-Service -Name $serviceName
                    if ($updatedService.Status -eq "Running") {
                        Write-Log "Successfully started $($service.DisplayName)" "SUCCESS"
                    } else {
                        Write-Log "Failed to start $($service.DisplayName)" "ERROR"
                    }
                } catch {
                    Write-Log "Error starting service: $($_.Exception.Message)" "ERROR"
                }
            }
        }
    } catch {
        # Service not found, continue to next
    }
}

if (-not $foundPervasiveService) {
    Write-Log "No Pervasive services found on this machine" "WARNING"
    Write-Log "The Pervasive client might not be installed correctly" "WARNING"
}

# Step 4: Check registry settings for Pervasive client
Write-Log "Checking Pervasive registry settings..." "INFO"
$registryPaths = @(
    "HKLM:\SOFTWARE\Pervasive Software",
    "HKLM:\SOFTWARE\WOW6432Node\Pervasive Software"
)

$foundRegistry = $false
foreach ($path in $registryPaths) {
    if (Test-Path $path) {
        $foundRegistry = $true
        Write-Log "Found Pervasive registry path: $path" "INFO"
        try {
            $regItems = Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue
            foreach ($item in $regItems) {
                if ($item.Name -like "*ServerName*" -or $item.Name -like "*DBNAMES*") {
                    $properties = Get-ItemProperty -Path $item.PSPath -ErrorAction SilentlyContinue
                    foreach ($prop in $properties.PSObject.Properties) {
                        if ($prop.Name -notlike "PS*") {
                            Write-Log "Registry key: $($item.Name) - $($prop.Name): $($prop.Value)" "INFO"
                        }
                    }
                }
            }
        } catch {
            Write-Log "Error accessing registry: $($_.Exception.Message)" "ERROR"
        }
    }
}

if (-not $foundRegistry) {
    Write-Log "No Pervasive registry entries found" "WARNING"
    Write-Log "The Pervasive client might not be installed or configured properly" "WARNING"
}

# Step 5: Check for Pervasive settings in the Windows hosts file
Write-Log "Checking hosts file for Pervasive server entries..." "INFO"
$hostsPath = "$env:windir\System32\drivers\etc\hosts"
if (Test-Path $hostsPath) {
    $hostsContent = Get-Content $hostsPath
    $pervasiveHostEntry = $hostsContent | Where-Object { $_ -match $PervasiveServerName }
    
    if ($pervasiveHostEntry) {
        Write-Log "Found entry for Pervasive server in hosts file: $pervasiveHostEntry" "INFO"
    } else {
        Write-Log "No entry found for $PervasiveServerName in hosts file" "INFO"
    }
} else {
    Write-Log "Hosts file not found" "WARNING"
}

# Step 6: Attempt to restart the Pervasive client services as a potential fix
if ($foundPervasiveService -and $isAdmin) {
    Write-Log "Attempting to restart all Pervasive client services as a potential fix..." "INFO"
    foreach ($serviceName in $PervasiveServiceNames) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
        if ($service) {
            try {
                Restart-Service -Name $serviceName -Force -ErrorAction Stop
                Start-Sleep -Seconds 2
                $updatedService = Get-Service -Name $serviceName
                Write-Log "$($service.DisplayName) restart attempt - New Status: $($updatedService.Status)" "INFO"
            } catch {
                Write-Log "Error restarting $($service.DisplayName): $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# Step 7: Create a test file to check local write permissions
Write-Log "Testing local file system write permissions..." "INFO"
$testFilePath = "$env:SYSTEMROOT\TEMP\pervasive_test.txt"
try {
    "Pervasive test file - Created at $(Get-Date)" | Out-File -FilePath $testFilePath -Force
    if (Test-Path $testFilePath) {
        Write-Log "Successfully created test file at $testFilePath" "SUCCESS"
        Remove-Item $testFilePath -Force
    } else {
        Write-Log "Failed to create test file" "ERROR"
    }
} catch {
    Write-Log "Error testing file system: $($_.Exception.Message)" "ERROR"
    Write-Log "This may indicate disk or permission issues" "WARNING"
}

# Step 8: Summary and recommendations
Write-Log "===== Troubleshooting Summary =====" "INFO"
Write-Log "Checking if any critical errors were found..." "INFO"

$logContent = Get-Content $LogPath
$criticalErrors = $logContent | Where-Object { $_ -match "\[ERROR\]" }
$warnings = $logContent | Where-Object { $_ -match "\[WARNING\]" }

if ($criticalErrors) {
    Write-Log "Critical errors were found during troubleshooting:" "ERROR"
    foreach ($error in $criticalErrors) {
        Write-Log "  - $($error.Substring($error.IndexOf(']') + 2))" "ERROR"
    }
    
    Write-Log "Recommended fixes:" "INFO"
    
    if ($logContent -match "Network connectivity") {
        Write-Log "  - Check network connection to the server $PervasiveServerName" "INFO"
        Write-Log "  - Verify server is online and reachable" "INFO"
    }
    
    if ($logContent -match "Port $PervasivePort is closed") {
        Write-Log "  - Ensure Pervasive database engine is running on the server" "INFO"
        Write-Log "  - Check firewall settings on both client and server" "INFO"
    }
    
    if ($logContent -match "Failed to start") {
        Write-Log "  - Check Pervasive client installation" "INFO"
        Write-Log "  - Reinstall Pervasive client if necessary" "INFO"
    }
} elseif ($warnings) {
    Write-Log "Only warnings found - the system may work but with issues" "WARNING"
} else {
    Write-Log "No critical issues detected. If the problem persists:" "SUCCESS"
    Write-Log "  - Try restarting the Accounts Payable application" "INFO"
    Write-Log "  - Verify database credentials in the application" "INFO"
}

Write-Log "Troubleshooting complete. Log file saved to: $LogPath" "INFO"
Write-Log "If issues persist, please contact your database administrator with this log file." "INFO"
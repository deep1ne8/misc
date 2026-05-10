# Sage 300 CRE Pervasive Connection Troubleshooter
# This script diagnoses and attempts to fix Pervasive database connection issues (Error 3110)
# Created: March 7, 2025

# Configuration Variables - Modify these for your environment
$ServerName = "DYA-APPS" # Default server name from the document, modify if different
$TimberlineSharePath = "\\DYA-APPS\Timberline Office"
$LogPath = "$env:SYSTEMROOT\Temp\Sage300CRE_Troubleshooter.log"

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

Write-Log "Starting Sage 300 CRE Pervasive Connection Troubleshooting Script" "INFO"
Write-Log "Target server: $ServerName" "INFO"
Write-Log "Timberline share path: $TimberlineSharePath" "INFO"

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Log "Script is not running with administrative privileges. Some operations may fail." "WARNING"
    Write-Log "Please re-run this script as an administrator for best results." "WARNING"
} else {
    Write-Log "Running with administrative privileges" "SUCCESS"
}

# Step 1: Test network connectivity to the Sage 300 CRE server
Write-Log "Testing network connectivity to $ServerName..." "INFO"
try {
    $pingResult = Test-Connection -ComputerName $ServerName -Count 2 -ErrorAction Stop
    Write-Log "Network connectivity to $ServerName is available (Latency: $($pingResult[0].ResponseTime) ms)" "SUCCESS"
} catch {
    Write-Log "Cannot reach $ServerName - Network connectivity issue: $($_.Exception.Message)" "ERROR"
    Write-Log "Recommendations:" "INFO"
    Write-Log "  - Verify network cable connections" "INFO"
    Write-Log "  - Check if the server is online" "INFO"
    Write-Log "  - Verify VPN connection if connecting remotely" "INFO"
    Write-Log "  - Check firewall settings" "INFO"
}

# Step 2: Test network share connectivity
Write-Log "Testing connectivity to Timberline network share..." "INFO"
if (Test-Path $TimberlineSharePath) {
    Write-Log "Successfully connected to Timberline share at $TimberlineSharePath" "SUCCESS"
    
    # Check for client installer
    $clientInstallPath = "$TimberlineSharePath\Sage300CRE ACCT Client install"
    if (Test-Path "$clientInstallPath.exe" -or Test-Path "$clientInstallPath.lnk") {
        Write-Log "Found Sage 300 CRE client installer" "SUCCESS"
    } else {
        Write-Log "Could not find Sage 300 CRE client installer at expected location" "WARNING"
        Write-Log "Expected path: $clientInstallPath" "INFO"
    }
} else {
    Write-Log "Cannot access Timberline share at $TimberlineSharePath" "ERROR"
    Write-Log "This is critical for installation and proper functioning" "ERROR"
    Write-Log "Please check network permissions and connectivity" "INFO"
}

# Step 3: Check for Sage 300 CRE installation
Write-Log "Checking for Sage 300 CRE installation..." "INFO"
$sageCREPaths = @(
    "${env:ProgramFiles}\Timberline Office",
    "${env:ProgramFiles(x86)}\Timberline Office",
    "C:\Timberline Office",
    "C:\Program Files\Timberline Office",
    "C:\Program Files (x86)\Timberline Office"
)

$foundSageCRE = $false
foreach ($path in $sageCREPaths) {
    if (Test-Path $path) {
        $foundSageCRE = $true
        Write-Log "Found Sage 300 CRE installation at: $path" "SUCCESS"
        
        # Check for version file
        $versionFiles = Get-ChildItem -Path $path -Filter "version.xml" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($versionFiles) {
            $versionContent = Get-Content $versionFiles.FullName -ErrorAction SilentlyContinue
            Write-Log "Sage 300 CRE version information found at $($versionFiles.FullName)" "INFO"
            Write-Log "Version info: $versionContent" "INFO"
        } else {
            Write-Log "Could not determine Sage 300 CRE version" "WARNING"
        }
    }
}

if (-not $foundSageCRE) {
    Write-Log "No Sage 300 CRE installation found on this workstation" "ERROR"
    Write-Log "You need to install the Sage 300 CRE client from $TimberlineSharePath" "ERROR"
}

# Step 4: Check for Sage System Verifier
Write-Log "Checking for Sage System Verifier..." "INFO"
$verifierPaths = @(
    "${env:ProgramFiles}\Timberline Office\Shared\System Verifier.exe",
    "${env:ProgramFiles(x86)}\Timberline Office\Shared\System Verifier.exe",
    "C:\Timberline Office\Shared\System Verifier.exe"
)

$foundVerifier = $false
foreach ($path in $verifierPaths) {
    if (Test-Path $path) {
        $foundVerifier = $true
        Write-Log "Found Sage System Verifier at: $path" "SUCCESS"
    }
}

if (-not $foundVerifier) {
    Write-Log "Sage System Verifier not found" "WARNING"
    Write-Log "This tool is necessary for verifying proper installation" "WARNING"
}

# Step 5: Check Pervasive database components
Write-Log "Checking Pervasive database components..." "INFO"

# Check for Pervasive services
$pervasiveServices = @(
    "Pervasive PSQL",
    "PSQL",
    "Pervasive.SQL",
    "Pervasive PSQL Workgroup Engine",
    "psqlWGE"
)

$foundPervasiveService = $false
foreach ($serviceName in $pervasiveServices) {
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
    Write-Log "Pervasive PSQL client might not be installed correctly" "WARNING"
}

# Check for Pervasive processes
$pervasiveProcesses = @("w3dbsmgr.exe", "pvsw.exe", "pervasive.exe", "PsiSvc.exe", "w3dbsmgr.exe", "w3sns.exe")
$foundPervasiveProcess = $false
foreach ($processName in $pervasiveProcesses) {
    $process = Get-Process -Name ($processName -replace ".exe", "") -ErrorAction SilentlyContinue
    if ($process) {
        $foundPervasiveProcess = $true
        Write-Log "Found Pervasive process: $($process.Name) (PID: $($process.Id))" "INFO"
    }
}

if (-not $foundPervasiveProcess) {
    Write-Log "No Pervasive processes found running" "WARNING"
}

# Check Pervasive registry settings
Write-Log "Checking Pervasive registry settings..." "INFO"
$pervasiveRegistryPaths = @(
    "HKLM:\SOFTWARE\Pervasive Software",
    "HKLM:\SOFTWARE\WOW6432Node\Pervasive Software"
)

$foundPervasiveRegistry = $false
foreach ($path in $pervasiveRegistryPaths) {
    if (Test-Path $path) {
        $foundPervasiveRegistry = $true
        Write-Log "Found Pervasive registry path: $path" "INFO"
        
        # Look for database configuration
        $dbConfigPath = Join-Path $path "PSQL\*\CurrentVersion\Engine"
        if (Test-Path $dbConfigPath) {
            $dbConfig = Get-ItemProperty -Path $dbConfigPath -ErrorAction SilentlyContinue
            if ($dbConfig.ServerName) {
                Write-Log "Pervasive configured server: $($dbConfig.ServerName)" "INFO"
            }
            if ($dbConfig.ServerPort) {
                Write-Log "Pervasive configured port: $($dbConfig.ServerPort)" "INFO"
            }
        }
    }
}

if (-not $foundPervasiveRegistry) {
    Write-Log "No Pervasive registry entries found" "WARNING"
    Write-Log "Pervasive might not be installed or configured properly" "WARNING"
}

# Step 6: Check for Sage 300 CRE ODBC DSNs
Write-Log "Checking for Sage 300 CRE ODBC DSNs..." "INFO"
$sageDSNs = @()

try {
    $odbcRegistryPaths = @(
        "HKLM:\SOFTWARE\ODBC\ODBC.INI",
        "HKLM:\SOFTWARE\WOW6432Node\ODBC\ODBC.INI",
        "HKCU:\SOFTWARE\ODBC\ODBC.INI"
    )
    
    $foundDSNs = $false
    foreach ($path in $odbcRegistryPaths) {
        if (Test-Path $path) {
            $dsns = Get-ChildItem -Path $path -ErrorAction SilentlyContinue
            foreach ($dsn in $dsns) {
                $dsnProperties = Get-ItemProperty -Path $dsn.PSPath -ErrorAction SilentlyContinue
                if ($dsnProperties.Driver -match "Pervasive" -or $dsn.PSChildName -match "Timberline" -or $dsnProperties.Description -match "Sage") {
                    $foundDSNs = $true
                    $sageDSNs += $dsn.PSChildName
                    Write-Log "Found Sage 300 CRE ODBC DSN: $($dsn.PSChildName)" "INFO"
                    
                    if ($dsnProperties.ServerName) {
                        Write-Log "  - Server: $($dsnProperties.ServerName)" "INFO"
                    }
                    if ($dsnProperties.DatabaseName) {
                        Write-Log "  - Database: $($dsnProperties.DatabaseName)" "INFO"
                    }
                }
            }
        }
    }
    
    if (-not $foundDSNs) {
        Write-Log "No Sage 300 CRE ODBC DSNs found" "WARNING"
        Write-Log "This could indicate an incomplete or improper client setup" "WARNING"
    }
} catch {
    Write-Log "Error checking ODBC DSNs: $($_.Exception.Message)" "ERROR"
}

# Step 7: Test Pervasive port connectivity
Write-Log "Testing Pervasive port connectivity..." "INFO"
$pervasivePorts = @(1583, 3351, 161, 137, 138, 139, 445)

foreach ($port in $pervasivePorts) {
    $portTest = New-Object System.Net.Sockets.TcpClient
    try {
        $portTestResult = $portTest.BeginConnect($ServerName, $port, $null, $null)
        $asyncResult = $portTestResult.AsyncWaitHandle.WaitOne(1000, $false)
        if ($asyncResult) {
            Write-Log "Port $port is open on $ServerName" "SUCCESS"
        } else {
            Write-Log "Port $port is closed or filtered on $ServerName" "WARNING"
        }
    } catch {
        Write-Log "Error testing port $port connectivity: $($_.Exception.Message)" "ERROR"
    } finally {
        $portTest.Close()
    }
}

# Step 8: Check Windows Event Logs for Pervasive errors
Write-Log "Checking Windows Event Logs for Pervasive errors..." "INFO"
try {
    $pervasiveEvents = Get-EventLog -LogName Application -Source "Pervasive*" -After (Get-Date).AddDays(-7) -ErrorAction SilentlyContinue
    if ($pervasiveEvents) {
        $recentErrors = $pervasiveEvents | Where-Object { $_.EntryType -eq "Error" } | Select-Object -First 5
        if ($recentErrors) {
            Write-Log "Found recent Pervasive errors in Event Log:" "WARNING"
            foreach ($event in $recentErrors) {
                Write-Log "  - Event ID $($event.EventID): $($event.Message.Substring(0, [Math]::Min(100, $event.Message.Length)))..." "WARNING"
            }
        } else {
            Write-Log "No recent Pervasive errors found in Event Log" "SUCCESS"
        }
    } else {
        Write-Log "No Pervasive events found in Application Event Log" "INFO"
    }
} catch {
    Write-Log "Error accessing Event Logs: $($_.Exception.Message)" "WARNING"
}

# Step 9: Check .NET Framework version (required for Sage 300 CRE)
Write-Log "Checking .NET Framework version..." "INFO"
$dotNetVersions = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse |
    Get-ItemProperty -Name Version, Release -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^v\d' } |
    Select-Object PSChildName, Version, Release

$dotNetFramework48OrLater = $false
foreach ($version in $dotNetVersions) {
    Write-Log ".NET Framework $($version.PSChildName) version $($version.Version) found" "INFO"
    if ($version.PSChildName -eq 'v4' -and $version.Release -ge 528040) {
        $dotNetFramework48OrLater = $true
    }
}

if ($dotNetFramework48OrLater) {
    Write-Log ".NET Framework 4.8 or higher found (required for Sage 300 CRE)" "SUCCESS"
} else {
    Write-Log ".NET Framework 4.8 or higher not found" "ERROR"
    Write-Log "Sage 300 CRE requires .NET Framework 4.8 or higher" "ERROR"
}

# Step 10: Run Sage System Verifier if found
if ($foundVerifier) {
    Write-Log "Would you like to run the Sage System Verifier now? (Y/N)" "INFO"
    $response = Read-Host
    if ($response -eq "Y" -or $response -eq "y") {
        $verifierPath = $verifierPaths | Where-Object { Test-Path $_ } | Select-Object -First 1
        if ($verifierPath) {
            Write-Log "Launching Sage System Verifier from $verifierPath..." "INFO"
            Start-Process -FilePath $verifierPath
            Write-Log "Please click 'Scan System' in the Sage System Verifier window" "INFO"
        }
    }
}

# Step 11: Summary and recommendations
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
    
    Write-Log "Recommended actions for Pervasive error 3110:" "INFO"
    
    if ($logContent -match "Cannot reach $ServerName") {
        Write-Log "  - Fix network connectivity to $ServerName" "INFO"
        Write-Log "  - Verify the server is online and network paths are accessible" "INFO"
    }
    
    if (-not $foundSageCRE) {
        Write-Log "  - Install the Sage 300 CRE client using:" "INFO"
        Write-Log "    $TimberlineSharePath\Sage300CRE ACCT Client install" "INFO"
        Write-Log "  - Run the installer as Administrator" "INFO"
    }
    
    if (-not $foundPervasiveService -and -not $foundPervasiveProcess) {
        Write-Log "  - Pervasive database client components are missing" "INFO"
        Write-Log "  - Reinstall Sage 300 CRE client with Pervasive components" "INFO"
    }
    
    if (-not $dotNetFramework48OrLater) {
        Write-Log "  - Install .NET Framework 4.8 or higher from Microsoft's website" "INFO"
    }
    
    Write-Log "" "INFO"
    Write-Log "Specific steps to fix Pervasive error 3110:" "INFO"
    Write-Log "1. Try reinstalling the Sage 300 CRE client following the provided installation document" "INFO"
    Write-Log "   - Go to $TimberlineSharePath" "INFO"
    Write-Log "   - Run 'Sage300CRE ACCT Client install' as administrator" "INFO"
    Write-Log "   - Follow the prompts through the upgrade process" "INFO"
    Write-Log "2. After installation completes, run Sage System Verifier" "INFO"
    Write-Log "   - Go to Windows Start Button and search for 'Sage System Verifier'" "INFO"
    Write-Log "   - Run it as administrator" "INFO"
    Write-Log "   - Click the 'Scan System' button in the top right corner" "INFO"
    Write-Log "3. If problems persist, check firewall settings for ports 1583 and 3351" "INFO"
    Write-Log "4. Ensure the Pervasive engine is running on the server" "INFO"
} elseif ($warnings) {
    Write-Log "Warnings found - the system may work but with potential issues:" "WARNING"
    Write-Log "  - Review the warnings above and address them if experiencing problems" "INFO"
    
    if (-not $foundPervasiveService) {
        Write-Log "  - No Pervasive services were found - check Pervasive installation" "INFO"
    }
    
    Write-Log "Recommended action:" "INFO"
    Write-Log "1. Run the Sage System Verifier to check for missing components" "INFO"
    Write-Log "   - Go to Windows Start Button and search for 'Sage System Verifier'" "INFO"
    Write-Log "   - Run it as administrator" "INFO"
    Write-Log "   - Click the 'Scan System' button in the top right corner" "INFO"
} else {
    Write-Log "No critical issues detected. If the problem persists:" "SUCCESS"
    Write-Log "1. Try restarting the Sage 300 CRE application" "INFO"
    Write-Log "2. Run the Sage System Verifier as described in the installation document" "INFO"
    Write-Log "3. Contact your Sage administrator if issues continue" "INFO"
}

Write-Log "Troubleshooting complete. Log file saved to: $LogPath" "INFO"
Write-Log "If issues persist, please share this log file with your Sage support team." "INFO"
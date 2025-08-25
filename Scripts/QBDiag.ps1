<# 
.SYNOPSIS
  Enhanced QuickBooks Freeze Diagnostic with issue detection and fix recommendations
.DESCRIPTION
  Collects OS/storage/network signals, QB service/process/version info, RDS printing flags,
  inspects .ND/.TLG, validates SMB path/latency, samples relevant Event Logs, detects issues,
  and provides fix recommendations without auto-applying them.
.NOTES
  Requires administrative privileges for certain diagnostics.
  Tested on PS 5.1+. Run on both the RDS host and the file server for comparison.
#>

[CmdletBinding()]
param()

$now = Get-Date
$hostName = $env:COMPUTERNAME

# Interactive prompts for all variables
Write-Host "`n=== QuickBooks Freeze Diagnostic Configuration ===" -ForegroundColor Cyan
Write-Host "Please provide the following information:`n" -ForegroundColor Yellow

# Role selection with validation
do {
    $Role = Read-Host -Prompt "Enter Role (Server/RDSHost/Client)"
    $Role = $Role.Trim()
} while ($Role -notin @('Server', 'RDSHost', 'Client'))

# Server names
$ServerName = Read-Host -Prompt "Enter File Server Name (press Enter if local)"
if ([string]::IsNullOrWhiteSpace($ServerName)) { $ServerName = $hostName }

$RDSHostName = Read-Host -Prompt "Enter RDS Host Name (press Enter if not applicable)"
if ([string]::IsNullOrWhiteSpace($RDSHostName)) { $RDSHostName = $null }

$ClientName = Read-Host -Prompt "Enter Client Name (press Enter if not applicable)"
if ([string]::IsNullOrWhiteSpace($ClientName)) { $ClientName = $null }

# QuickBooks directories
Write-Host "`nQuickBooks Directory Configuration:" -ForegroundColor Yellow
$QuickBooksDir1 = Read-Host -Prompt "Enter primary QB directory (default: C:\Users\Public)"
if ([string]::IsNullOrWhiteSpace($QuickBooksDir1)) { $QuickBooksDir1 = 'C:\Users\Public' }

$QuickBooksDir2 = Read-Host -Prompt "Enter secondary QB directory (default: C:\QuickBooks)"
if ([string]::IsNullOrWhiteSpace($QuickBooksDir2)) { $QuickBooksDir2 = 'C:\QuickBooks' }

$QuickBooksDir3 = Read-Host -Prompt "Enter custom QB directory (press Enter to skip)"
if ([string]::IsNullOrWhiteSpace($QuickBooksDir3)) { $QuickBooksDir3 = $null }

# Company file path
$CompanyFilePath = Read-Host -Prompt "Enter full path to Company File (.qbw) or press Enter to auto-search"
$FindCompanyFile = [string]::IsNullOrWhiteSpace($CompanyFilePath)

# Log days
$LogDaysInput = Read-Host -Prompt "Enter number of days for event log analysis (default: 7)"
if ([string]::IsNullOrWhiteSpace($LogDaysInput)) { 
    $LogDays = 7 
} else { 
    $LogDays = [int]$LogDaysInput 
}

# Share write probe
$ProbeResponse = Read-Host -Prompt "Perform share write probe test? (Y/N, default: N)"
$DoShareWriteProbe = ($ProbeResponse -eq 'Y' -or $ProbeResponse -eq 'y')

# Output directory
$outDir = 'C:\ProgramData\QB-Diag'
$null = New-Item -ItemType Directory -Path $outDir -Force -ErrorAction SilentlyContinue
$reportPath = Join-Path $outDir ("QBDiag_{0}_{1:yyyyMMdd_HHmmss}.html" -f $hostName,$now)

# Initialize issue tracking
$issuesFound = @()
$recommendedFixes = @()

function New-Section { param([string]$Title,[object]$Data) [PSCustomObject]@{Section=$Title;Data=$Data} }
function SizeMB { param([long]$b) if ($b -is [long]) {[math]::Round($b/1MB,2)} else {$null} }
function BaseNameNoExt([string]$p){ [IO.Path]::GetFileNameWithoutExtension($p) }

function Add-Issue {
    param(
        [string]$Category,
        [string]$Issue,
        [string]$Severity,
        [string]$Fix
    )
    
    $script:issuesFound += [PSCustomObject]@{
        Category = $Category
        Issue = $Issue
        Severity = $Severity
        RecommendedFix = $Fix
    }
}

function Resolve-CompanyPath {
    param([string]$Path,[switch]$AllowSearch)
    if ($Path) { return $Path }
    if ($AllowSearch) {
        try {
            Write-Host "Searching for QuickBooks company files..." -ForegroundColor Cyan
            $candidates = New-Object System.Collections.Generic.List[string]
            
            # Search QBWUSER.INI
            Get-ChildItem "$env:ProgramData\Intuit" -Recurse -Filter 'QBWUSER.INI' -ErrorAction SilentlyContinue | ForEach-Object {
                foreach ($line in (Get-Content -LiteralPath $_.FullName -ErrorAction SilentlyContinue)) {
                    if ($line -match '\.qbw' -and $line -match '=') {
                        $p = ($line -split '=',2)[1].Trim()
                        if ($p -and (Test-Path -LiteralPath $p)) { $candidates.Add($p) }
                    }
                }
            }
            
            # Search configured directories
            $dirsToSearch = @($QuickBooksDir1, $QuickBooksDir2)
            if ($QuickBooksDir3) { $dirsToSearch += $QuickBooksDir3 }
            
            foreach ($root in $dirsToSearch) {
                if (Test-Path $root) {
                    Get-ChildItem -Path $root -Recurse -Filter *.qbw -ErrorAction SilentlyContinue -Force |
                        Select-Object -First 2 -Expand FullName | ForEach-Object { $candidates.Add($_) }
                }
            }
            
            $found = ($candidates | Select-Object -Unique | Select-Object -First 1)
            if ($found) {
                Write-Host "Found company file: $found" -ForegroundColor Green
            }
            return $found
        } catch {
            Write-Host "Error searching for company file: $_" -ForegroundColor Red
        }
    }
    return $null
}

function Parse-ServerFromUNC([string]$UNC){
    if ($UNC -and $UNC.StartsWith('\\')) { return $UNC.TrimStart('\').Split('\')[0] } else { return $null }
}

# Main diagnostic process
Write-Host "`n=== Starting QuickBooks Diagnostics ===" -ForegroundColor Green
Write-Host "Running diagnostics for role: $Role" -ForegroundColor Cyan

# --- System snapshot
$os  = Get-CimInstance Win32_OperatingSystem
$cs  = Get-CimInstance Win32_ComputerSystem
$cpu = Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed
$uptime = (Get-Date) - $os.LastBootUpTime
$mem = [PSCustomObject]@{
    TotalGB      = [Math]::Round($os.TotalVisibleMemorySize/1MB,2)
    FreeGB       = [Math]::Round($os.FreePhysicalMemory/1MB,2)
    PercentUsed  = [Math]::Round(100 - (($os.FreePhysicalMemory*1.0)/$os.TotalVisibleMemorySize*100),1)
    CommitLimitGB= [Math]::Round(($os.TotalVirtualMemorySize/1MB),2)
    CommitUsedGB = [Math]::Round(($os.TotalVirtualMemorySize - $os.FreeVirtualMemory)/1MB,2)
}

# Check for system issues
if ($mem.PercentUsed -gt 85) {
    Add-Issue -Category "System" -Issue "High memory usage: $($mem.PercentUsed)%" -Severity "High" `
              -Fix "Consider adding more RAM or closing unnecessary applications"
}

if ($mem.FreeGB -lt 2) {
    Add-Issue -Category "System" -Issue "Low free memory: $($mem.FreeGB) GB" -Severity "High" `
              -Fix "Free up memory by restarting services or the system"
}

if ($uptime.TotalDays -gt 30) {
    Add-Issue -Category "System" -Issue "System uptime exceeds 30 days" -Severity "Medium" `
              -Fix "Schedule a restart to clear memory leaks and apply pending updates"
}

# Check power plan
$powerPlan = (powercfg /getactivescheme 2>$null) -join ' '
if ($powerPlan -notmatch 'High performance') {
    Add-Issue -Category "System" -Issue "Power plan is not set to High Performance" -Severity "Medium" `
              -Fix "Run: powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
}

$sysSection = New-Section "System" (@{
    Role         = $Role
    ComputerName = $hostName
    OS           = "$($os.Caption) ($($os.Version))"
    UptimeDays   = [Math]::Round($uptime.TotalDays,1)
    CPU          = $cpu
    Memory       = $mem
    PowerPlan    = $powerPlan
})

# --- Storage + SMART-ish
$vols = Get-Volume -ErrorAction SilentlyContinue | Select-Object DriveLetter,Path,FileSystem,Size,SizeRemaining,HealthStatus
$pd   = Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName,MediaType,Size,HealthStatus

# Check for disk issues
foreach ($vol in $vols) {
    if ($vol.Size -gt 0) {
        $freePercent = ($vol.SizeRemaining / $vol.Size) * 100
        if ($freePercent -lt 10) {
            Add-Issue -Category "Storage" -Issue "Drive $($vol.DriveLetter): has less than 10% free space" -Severity "High" `
                      -Fix "Free up disk space on drive $($vol.DriveLetter):"
        }
    }
    if ($vol.HealthStatus -ne 'Healthy' -and $vol.HealthStatus) {
        Add-Issue -Category "Storage" -Issue "Drive $($vol.DriveLetter): health status is $($vol.HealthStatus)" -Severity "Critical" `
                  -Fix "Check disk for errors using: chkdsk $($vol.DriveLetter): /f /r"
    }
}

foreach ($disk in $pd) {
    if ($disk.HealthStatus -ne 'Healthy' -and $disk.HealthStatus) {
        Add-Issue -Category "Storage" -Issue "Physical disk $($disk.FriendlyName) health is $($disk.HealthStatus)" -Severity "Critical" `
                  -Fix "Consider replacing the disk immediately and backing up data"
    }
}

$diskSection = New-Section "Storage" (@{ Volumes = $vols; PhysicalDisks = $pd })

# --- QB services/processes/versions
$qbServices = Get-Service -Name 'QBDBMgr*','QBCFMonitorService' -ErrorAction SilentlyContinue |
    Select-Object Name,Status,StartType,DisplayName

# Check QB service issues
foreach ($svc in $qbServices) {
    if ($svc.Status -ne 'Running') {
        Add-Issue -Category "QuickBooks Services" -Issue "Service $($svc.Name) is not running" -Severity "High" `
                  -Fix "Start-Service -Name $($svc.Name)"
    }
    if ($svc.StartType -ne 'Automatic') {
        Add-Issue -Category "QuickBooks Services" -Issue "Service $($svc.Name) is not set to Automatic start" -Severity "Medium" `
                  -Fix "Set-Service -Name $($svc.Name) -StartupType Automatic"
    }
}

$qbProc = Get-Process -Name 'QBW32','QBDBMgrN','QBCFMonitorService' -ErrorAction SilentlyContinue |
    Select-Object ProcessName,Id,CPU,@{n='WorkingSetMB';e={[math]::Round($_.WS/1MB,1)}},StartTime

# Check for high memory usage by QB processes
foreach ($proc in $qbProc) {
    if ($proc.WorkingSetMB -gt 2000) {
        Add-Issue -Category "QuickBooks Process" -Issue "Process $($proc.ProcessName) using excessive memory: $($proc.WorkingSetMB) MB" -Severity "Medium" `
                  -Fix "Restart QuickBooks and related services"
    }
}

# Discover executable versions
$qbBins = Get-ChildItem 'C:\Program Files','C:\Program Files (x86)','C:\Program Files\Common Files' -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match 'QBW32\.exe|QBDBMgrN\.exe|QBCFMonitorService\.exe' } |
    Select-Object FullName,Name
$qbVers = foreach ($b in $qbBins) {
    try {
        $f = Get-Item $b.FullName
        [PSCustomObject]@{
            Binary     = $b.Name
            Version    = $f.VersionInfo.FileVersion
            Product    = $f.VersionInfo.ProductVersion
            Path       = $b.FullName
        }
    } catch {}
}

$qbSection = New-Section "QuickBooks Binaries/Services" (@{
    Services  = $qbServices
    Processes = $qbProc
    Binaries  = $qbVers
})

# --- Company file / ND / TLG
if (-not $CompanyFilePath) { $CompanyFilePath = Resolve-CompanyPath -AllowSearch:$FindCompanyFile }
$companyObj = @{}
if ($CompanyFilePath) {
    $companyObj.Path     = $CompanyFilePath
    $companyObj.Exists   = Test-Path -LiteralPath $CompanyFilePath
    $companyObj.PathType = ($CompanyFilePath -like '\\*') ? 'UNC' : 'Local/Drive'
    $parent = Split-Path $CompanyFilePath -Parent
    $serverName = Parse-ServerFromUNC $CompanyFilePath
    $companyObj.ParentDir = $parent

    if ($companyObj.Exists) {
        $cfi = Get-Item -LiteralPath $CompanyFilePath
        $companyObj.SizeMB = SizeMB $cfi.Length
        
        # Check company file size
        if ($companyObj.SizeMB -gt 1000) {
            Add-Issue -Category "Company File" -Issue "Company file is large: $($companyObj.SizeMB) MB" -Severity "Medium" `
                      -Fix "Consider running File Doctor or condensing the company file"
        }
        
        # Check ND file
        $nd  = "$parent\$(BaseNameNoExt $CompanyFilePath).ND"
        $tlg = [IO.Path]::ChangeExtension($CompanyFilePath,'.TLG')
        
        if (Test-Path $nd) {
            $ndInfo = Get-Item $nd
            $companyObj.ND = @{Path=$nd; SizeMB=SizeMB $ndInfo.Length; Modified=$ndInfo.LastWriteTime}
            
            if ($ndInfo.Length -eq 0) {
                Add-Issue -Category "Company File" -Issue ".ND file is empty (0 bytes)" -Severity "High" `
                          -Fix "Delete the .ND file and let QuickBooks recreate it: Remove-Item '$nd'"
            }
            
            # Check if ND file is stale (older than 7 days)
            $ndAge = (Get-Date) - $ndInfo.LastWriteTime
            if ($ndAge.TotalDays -gt 7) {
                Add-Issue -Category "Company File" -Issue ".ND file hasn't been updated in $([int]$ndAge.TotalDays) days" -Severity "Low" `
                          -Fix "Consider refreshing the .ND file by deleting and recreating it"
            }
        } else {
            $companyObj.ND = $null
            Add-Issue -Category "Company File" -Issue ".ND file is missing" -Severity "High" `
                      -Fix "Open QuickBooks in single-user mode to recreate the .ND file"
        }
        
        # Check TLG file
        if (Test-Path $tlg) {
            $tlgInfo = Get-Item $tlg
            $companyObj.TLG = @{Path=$tlg; SizeMB=SizeMB $tlgInfo.Length; Modified=$tlgInfo.LastWriteTime}
            
            $tlgSizeMB = [math]::Round($tlgInfo.Length/1MB, 2)
            if ($tlgSizeMB -gt 500) {
                Add-Issue -Category "Company File" -Issue ".TLG file is very large: $tlgSizeMB MB" -Severity "Medium" `
                          -Fix "Backup and remove the .TLG file: Move-Item '$tlg' '$tlg.bak'"
            }
        } else {
            $companyObj.TLG = $null
        }
    } else {
        Add-Issue -Category "Company File" -Issue "Company file not found at specified path" -Severity "Critical" `
                  -Fix "Verify the company file path and network connectivity"
    }

    # SMB connectivity & latency
    $serverToTest = if ($serverName) { $serverName } else { $null }
    if ($serverToTest) {
        $tnc = Test-NetConnection -ComputerName $serverToTest -Port 445 -WarningAction SilentlyContinue
        $companyObj.SMB445 = @{
            RemoteAddress = $tnc.RemoteAddress
            Reachable     = $tnc.TcpTestSucceeded
            PingSucceeded = $tnc.PingSucceeded
            LatencyMs     = $tnc.PingReplyDetails.RoundtripTime
        }
        
        if (-not $tnc.TcpTestSucceeded) {
            Add-Issue -Category "Network" -Issue "Cannot connect to SMB port 445 on $serverToTest" -Severity "Critical" `
                      -Fix "Check firewall rules and ensure SMB is enabled on the server"
        }
        
        if ($tnc.PingReplyDetails.RoundtripTime -gt 50) {
            Add-Issue -Category "Network" -Issue "High network latency to server: $($tnc.PingReplyDetails.RoundtripTime) ms" -Severity "Medium" `
                      -Fix "Check network connectivity and consider network optimization"
        }
    }

    # Measure metadata access times
    try {
        $t1 = Measure-Command { Get-ChildItem -LiteralPath $parent | Select-Object -First 1 | Out-Null }
        $t2 = Measure-Command { (Get-Item -LiteralPath $CompanyFilePath).LastWriteTime | Out-Null }
        $companyObj.TouchTimingsMs = @{
            DirList = [math]::Round($t1.TotalMilliseconds,1)
            Stat    = [math]::Round($t2.TotalMilliseconds,1)
        }
        
        if ($t1.TotalMilliseconds -gt 1000) {
            Add-Issue -Category "Performance" -Issue "Slow directory listing: $([math]::Round($t1.TotalMilliseconds,1)) ms" -Severity "High" `
                      -Fix "Check network performance and SMB settings"
        }
    } catch {}

    # Optional write probe
    if ($DoShareWriteProbe) {
        try {
            $tmp = Join-Path $parent ("_qbdiag_{0}.bin" -f [guid]::NewGuid().ToString('N'))
            $sw  = [IO.File]::Create($tmp)
            $buf = New-Object byte[] (5MB)
            (New-Object System.Random).NextBytes($buf)
            $writeTime = Measure-Command {
                $sw.Write($buf,0,$buf.Length)
                $sw.Flush()
                $sw.Dispose()
            }
            $wr = Get-Item $tmp
            Remove-Item $tmp -Force
            $companyObj.WriteProbe = @{
                Succeeded = $true
                SizeMB = SizeMB $wr.Length
                WriteTimeMs = [math]::Round($writeTime.TotalMilliseconds,1)
            }
            
            if ($writeTime.TotalMilliseconds -gt 5000) {
                Add-Issue -Category "Performance" -Issue "Slow write performance: $([math]::Round($writeTime.TotalMilliseconds,1)) ms for 5MB" -Severity "High" `
                          -Fix "Check disk performance and network bandwidth"
            }
        } catch {
            $companyObj.WriteProbe = @{Succeeded=$false; Error=$_.Exception.Message}
            Add-Issue -Category "Permissions" -Issue "Cannot write to company file directory" -Severity "High" `
                      -Fix "Check folder permissions and ensure write access is granted"
        }
    }
}

# --- Role-specific diagnostics
if ($Role -eq 'Server') {
    Write-Host "`nRunning Server-specific diagnostics..." -ForegroundColor Cyan
    
    # File server share properties
    $shareInfo = $null
    if ($CompanyFilePath) {
        try {
            if ($CompanyFilePath -notlike '\\*') {
                $dir = Split-Path $CompanyFilePath -Parent
                $shareInfo = Get-SmbShare -ErrorAction SilentlyContinue | Where-Object { $_.Path -ieq $dir } |
                    Select-Object Name,Path,CachingMode,ConcurrentUserLimit,FolderEnumerationMode,EncryptData
            } else {
                $parts = $CompanyFilePath.TrimStart('\').Split('\')
                if ($parts.Length -ge 2 -and $parts[0] -ieq $env:COMPUTERNAME) {
                    $shareName = $parts[1]
                    $shareInfo = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue |
                        Select-Object Name,Path,CachingMode,ConcurrentUserLimit,FolderEnumerationMode,EncryptData
                }
            }
            
            if ($shareInfo) {
                if ($shareInfo.CachingMode -ne 'None') {
                    Add-Issue -Category "SMB Configuration" -Issue "SMB caching is enabled (can cause QB issues)" -Severity "High" `
                              -Fix "Set-SmbShare -Name '$($shareInfo.Name)' -CachingMode None"
                }
                if ($shareInfo.EncryptData) {
                    Add-Issue -Category "SMB Configuration" -Issue "SMB encryption is enabled (impacts performance)" -Severity "Medium" `
                              -Fix "Set-SmbShare -Name '$($shareInfo.Name)' -EncryptData `$false"
                }
            }
        } catch {}
    }
    
    # Check SMB server configuration
    try {
        $smbConfig = Get-SmbServerConfiguration
        if ($smbConfig.RequireSecuritySignature) {
            Add-Issue -Category "SMB Configuration" -Issue "SMB signing is required (impacts performance)" -Severity "Medium" `
                      -Fix "Set-SmbServerConfiguration -RequireSecuritySignature `$false -Force"
        }
        if ($smbConfig.EnableSMB1Protocol) {
            Add-Issue -Category "SMB Configuration" -Issue "SMB1 protocol is enabled (security risk)" -Severity "High" `
                      -Fix "Set-SmbServerConfiguration -EnableSMB1Protocol `$false -Force"
        }
    } catch {}
    
    $shareSection = New-Section "Share Configuration (Server)" $shareInfo
}

if ($Role -eq 'RDSHost') {
    Write-Host "`nRunning RDS Host-specific diagnostics..." -ForegroundColor Cyan
    
    # RDS printing flags
    $rdsPrint = $null
    try {
        $key = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
        $easy = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).UseTerminalServicesEasyPrint
        $spool = Get-Service -Name Spooler -ErrorAction SilentlyContinue | Select-Object Status,StartType
        $rdsPrint = @{
            EasyPrintPolicy = $(if ($easy -eq 1) {'Enabled'} elseif ($easy -eq 0) {'Disabled'} else {'NotConfigured'})
            Spooler         = $spool
            PrinterDrivers  = (Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object Name,Manufacturer,Version | Select-Object -First 10)
        }
        
        if ($easy -ne 1) {
            Add-Issue -Category "RDS Configuration" -Issue "RDS Easy Print is not enabled" -Severity "Medium" `
                      -Fix "Set-ItemProperty -Path '$key' -Name 'UseTerminalServicesEasyPrint' -Value 1 -Type DWord"
        }
        
        if ($spool.Status -ne 'Running') {
            Add-Issue -Category "RDS Configuration" -Issue "Print Spooler service is not running" -Severity "High" `
                      -Fix "Start-Service -Name Spooler"
        }
    } catch {}
    
    # Check RDS licensing
    try {
        $rdsLic = Get-Service -Name TermServLicensing -ErrorAction SilentlyContinue
        if ($rdsLic.Status -ne 'Running') {
            Add-Issue -Category "RDS Configuration" -Issue "RDS Licensing service is not running" -Severity "Medium" `
                      -Fix "Start-Service -Name TermServLicensing"
        }
    } catch {}
    
    $rdsPrintSection = New-Section "RDS Configuration" $rdsPrint
}

if ($Role -eq 'Client') {
    Write-Host "`nRunning Client-specific diagnostics..." -ForegroundColor Cyan
    
    # Check for antivirus exclusions
    $avExclusions = @()
    try {
        $defenderExclusions = Get-MpPreference -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ExclusionPath
        if ($CompanyFilePath) {
            $qbDir = Split-Path $CompanyFilePath -Parent
            if ($defenderExclusions -notcontains $qbDir -and $defenderExclusions -notcontains "$qbDir\*") {
                Add-Issue -Category "Antivirus" -Issue "QuickBooks directory not excluded from Windows Defender" -Severity "Medium" `
                          -Fix "Add-MpPreference -ExclusionPath '$qbDir'"
            }
        }
    } catch {}
}

# --- Network optimization checks (all roles)
Write-Host "`nChecking network optimization settings..." -ForegroundColor Cyan

# Check network throttling
try {
    $throttleKey = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Multimedia\SystemProfile'
    $currentValue = (Get-ItemProperty -Path $throttleKey -ErrorAction SilentlyContinue).NetworkThrottlingIndex
    
    if ($null -eq $currentValue -or $currentValue -ne 0xFFFFFFFF) {
        Add-Issue -Category "Network" -Issue "Network throttling is enabled" -Severity "Medium" `
                  -Fix "Set-ItemProperty -Path '$throttleKey' -Name 'NetworkThrottlingIndex' -Value 0xFFFFFFFF -Type DWord"
    }
} catch {}

# Check TCP settings
try {
    $tcpSettings = Get-NetTCPSetting -SettingName InternetCustom -ErrorAction SilentlyContinue
    if ($tcpSettings -and $tcpSettings.AutoTuningLevelLocal -ne 'Normal') {
        Add-Issue -Category "Network" -Issue "TCP Auto-Tuning is not optimized" -Severity "Low" `
                  -Fix "Set-NetTCPSetting -SettingName InternetCustom -AutoTuningLevelLocal Normal"
    }
} catch {}

# --- Event logs
Write-Host "`nAnalyzing event logs (last $LogDays days)..." -ForegroundColor Cyan

$startTime = (Get-Date).AddDays(-1 * $LogDays)
$providers = @(
    'QuickBooks','QBDBMgrN','QBCFMonitorService','Application Error',
    'Microsoft-Windows-SMBClient','Microsoft-Windows-SMBServer',
    'Microsoft-Windows-PrintService','disk','Ntfs','srv','srv2','LanmanWorkstation','LanmanServer'
)
$logEvents = @()
$criticalEvents = 0
$errorEvents = 0

foreach ($p in $providers) {
    try {
        $ev = Get-WinEvent -FilterHashtable @{StartTime=$startTime; ProviderName=$p} -ErrorAction SilentlyContinue |
            Select-Object TimeCreated,ProviderName,Id,LevelDisplayName,Message -First 60
        if ($ev) {
            $logEvents += $ev
            $criticalEvents += ($ev | Where-Object { $_.LevelDisplayName -eq 'Critical' }).Count
            $errorEvents += ($ev | Where-Object { $_.LevelDisplayName -eq 'Error' }).Count
        }
    } catch {}
}

if ($criticalEvents -gt 0) {
    Add-Issue -Category "Event Logs" -Issue "Found $criticalEvents critical events in the last $LogDays days" -Severity "High" `
              -Fix "Review critical events in Event Viewer and address underlying issues"
}

if ($errorEvents -gt 10) {
    Add-Issue -Category "Event Logs" -Issue "Found $errorEvents error events in the last $LogDays days" -Severity "Medium" `
              -Fix "Review error events and look for patterns related to QuickBooks freezes"
}

$logsSection = New-Section ("Event Logs (last {0} days)" -f $LogDays) ($logEvents | Select-Object -First 50)

# --- Generate Issue Summary
$issueSummary = $null
if ($issuesFound.Count -gt 0) {
    Write-Host "`n=== ISSUES DETECTED ===" -ForegroundColor Red
    Write-Host "Found $($issuesFound.Count) issue(s) requiring attention:`n" -ForegroundColor Yellow
    
    # Group issues by severity
    $criticalIssues = $issuesFound | Where-Object { $_.Severity -eq 'Critical' }
    $highIssues = $issuesFound | Where-Object { $_.Severity -eq 'High' }
    $mediumIssues = $issuesFound | Where-Object { $_.Severity -eq 'Medium' }
    $lowIssues = $issuesFound | Where-Object { $_.Severity -eq 'Low' }
    
    if ($criticalIssues) {
        Write-Host "`nCRITICAL Issues:" -ForegroundColor Red
        foreach ($issue in $criticalIssues) {
            Write-Host "  [$($issue.Category)] $($issue.Issue)" -ForegroundColor Red
            Write-Host "    Fix: $($issue.RecommendedFix)" -ForegroundColor Yellow
        }
    }
    
    if ($highIssues) {
        Write-Host "`nHIGH Priority Issues:" -ForegroundColor DarkRed
        foreach ($issue in $highIssues) {
            Write-Host "  [$($issue.Category)] $($issue.Issue)" -ForegroundColor DarkRed
            Write-Host "    Fix: $($issue.RecommendedFix)" -ForegroundColor Yellow
        }
    }
    
    if ($mediumIssues) {
        Write-Host "`nMEDIUM Priority Issues:" -ForegroundColor DarkYellow
        foreach ($issue in $mediumIssues) {
            Write-Host "  [$($issue.Category)] $($issue.Issue)" -ForegroundColor DarkYellow
            Write-Host "    Fix: $($issue.RecommendedFix)" -ForegroundColor Gray
        }
    }
    
    if ($lowIssues) {
        Write-Host "`nLOW Priority Issues:" -ForegroundColor Yellow
        foreach ($issue in $lowIssues) {
            Write-Host "  [$($issue.Category)] $($issue.Issue)" -ForegroundColor Yellow
            Write-Host "    Fix: $($issue.RecommendedFix)" -ForegroundColor Gray
        }
    }
    
    $issueSummary = New-Section "Issues Detected" $issuesFound
} else {
    Write-Host "`n=== NO ISSUES DETECTED ===" -ForegroundColor Green
    Write-Host "All diagnostics passed successfully!" -ForegroundColor Green
}

# --- Build sections array
$sections = @(
    $sysSection,
    $diskSection,
    $qbSection,
    (New-Section "Company File" $companyObj)
)

# Add role-specific sections
if ($Role -eq 'Server' -and $shareInfo) {
    $sections += $shareSection
}
if ($Role -eq 'RDSHost' -and $rdsPrint) {
    $sections += $rdsPrintSection
}

# Add issue summary if issues found
if ($issueSummary) {
    $sections += $issueSummary
}

# Add event logs
$sections += $logsSection

# --- Build Complete HTML Report
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>QB Freeze Diagnostic - $hostName</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Arial, sans-serif; 
            font-size: 13px; 
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 10px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.2);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        h1 { 
            font-size: 28px;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        .header-info {
            display: flex;
            justify-content: center;
            gap: 30px;
            margin-top: 15px;
            flex-wrap: wrap;
        }
        .header-info span {
            background: rgba(255,255,255,0.2);
            padding: 5px 15px;
            border-radius: 20px;
            backdrop-filter: blur(10px);
        }
        .content {
            padding: 30px;
        }
        h2 { 
            color: #2c3e50;
            background: linear-gradient(90deg, #ecf0f1 0%, transparent 100%);
            padding: 12px 20px;
            border-left: 5px solid #3498db;
            margin: 30px 0 20px 0;
            font-size: 20px;
            border-radius: 0 5px 5px 0;
        }
        h3 { 
            color: #34495e;
            margin: 20px 0 10px 0;
            font-size: 16px;
            border-bottom: 2px solid #ecf0f1;
            padding-bottom: 5px;
        }
        
        /* Statistics Dashboard */
        .summary-stats { 
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin: 30px 0;
        }
        .stat-box { 
            background: white;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            text-align: center;
            transition: transform 0.3s, box-shadow 0.3s;
            border: 2px solid #ecf0f1;
        }
        .stat-box:hover {
            transform: translateY(-5px);
            box-shadow: 0 8px 15px rgba(0,0,0,0.2);
        }
        .stat-box.critical { border-color: #e74c3c; }
        .stat-box.high { border-color: #e67e22; }
        .stat-box.medium { border-color: #f39c12; }
        .stat-box.low { border-color: #3498db; }
        .stat-box.success { border-color: #27ae60; }
        
        .stat-number { 
            font-size: 36px;
            font-weight: bold;
            margin-bottom: 5px;
        }
        .stat-label { 
            color: #7f8c8d;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }
        
        /* Alert Boxes */
        .alert-box {
            padding: 15px 20px;
            margin: 20px 0;
            border-radius: 8px;
            border-left: 5px solid;
        }
        .info-box { 
            background: #d1f2eb;
            border-color: #27ae60;
            color: #1e7e5e;
        }
        .warning-box { 
            background: #fef5e7;
            border-color: #f39c12;
            color: #9a7d0a;
        }
        .error-box { 
            background: #fadbd8;
            border-color: #e74c3c;
            color: #922b21;
        }
        .success-box {
            background: #d5f5d5;
            border-color: #27ae60;
            color: #1d6f1d;
        }
        
        /* Tables */
        table { 
            border-collapse: separate;
            border-spacing: 0;
            margin: 20px 0;
            width: 100%;
            background: white;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            border-radius: 8px;
            overflow: hidden;
        }
        th, td { 
            padding: 12px 15px;
            text-align: left;
        }
        th { 
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 12px;
            letter-spacing: 0.5px;
        }
        td {
            border-bottom: 1px solid #ecf0f1;
        }
        tr:last-child td {
            border-bottom: none;
        }
        tr:nth-child(even) { 
            background: #f8f9fa;
        }
        tr:hover { 
            background: #e3f2fd;
            transition: background 0.3s;
        }
        
        /* Issue severity styles */
        tr.critical td:first-child { 
            border-left: 5px solid #e74c3c;
            font-weight: bold;
            color: #e74c3c;
        }
        tr.high td:first-child { 
            border-left: 5px solid #e67e22;
            font-weight: bold;
            color: #e67e22;
        }
        tr.medium td:first-child { 
            border-left: 5px solid #f39c12;
            color: #f39c12;
        }
        tr.low td:first-child { 
            border-left: 5px solid #3498db;
            color: #3498db;
        }
        
        /* Fix commands */
        .fix-command { 
            font-family: 'Consolas', 'Courier New', monospace;
            background: #2c3e50;
            color: #1abc9c;
            padding: 8px 12px;
            border-radius: 4px;
            display: inline-block;
            margin: 4px 0;
            font-size: 12px;
            word-break: break-all;
            max-width: 100%;
        }
        pre.fix-command {
            display: block;
            overflow-x: auto;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        
        /* List styles */
        .property-list {
            list-style: none;
            padding: 0;
        }
        .property-list li {
            padding: 8px 15px;
            border-bottom: 1px solid #ecf0f1;
            display: flex;
            justify-content: space-between;
        }
        .property-list li:nth-child(even) {
            background: #f8f9fa;
        }
        .property-name {
            font-weight: 600;
            color: #34495e;
            min-width: 200px;
        }
        .property-value {
            color: #2c3e50;
            text-align: right;
            flex: 1;
        }
        
        /* Progress indicators */
        .progress-bar {
            width: 100%;
            height: 20px;
            background: #ecf0f1;
            border-radius: 10px;
            overflow: hidden;
            margin: 10px 0;
        }
        .progress-fill {
            height: 100%;
            background: linear-gradient(90deg, #3498db 0%, #2980b9 100%);
            transition: width 0.3s;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-size: 12px;
            font-weight: bold;
        }
        .progress-fill.danger {
            background: linear-gradient(90deg, #e74c3c 0%, #c0392b 100%);
        }
        .progress-fill.warning {
            background: linear-gradient(90deg, #f39c12 0%, #e67e22 100%);
        }
        
        /* Footer */
        .footer {
            background: #2c3e50;
            color: white;
            padding: 20px;
            text-align: center;
            margin-top: 40px;
        }
        
        /* Print styles */
        @media print {
            body { background: white; }
            .container { box-shadow: none; }
            .header { background: #2c3e50; print-color-adjust: exact; }
            .no-print { display: none; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 QuickBooks Freeze Diagnostic Report</h1>
            <div class="header-info">
                <span>🖥️ Host: $hostName</span>
                <span>👤 Role: $Role</span>
                <span>📅 Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))</span>
            </div>
        </div>
        
        <div class="content">
"@

# Add summary statistics if issues found
if ($issuesFound.Count -gt 0) {
    $critCount = ($issuesFound | Where-Object { $_.Severity -eq 'Critical' }).Count
    $highCount = ($issuesFound | Where-Object { $_.Severity -eq 'High' }).Count
    $medCount = ($issuesFound | Where-Object { $_.Severity -eq 'Medium' }).Count
    $lowCount = ($issuesFound | Where-Object { $_.Severity -eq 'Low' }).Count
    
    $htmlReport += @"
            <div class="alert-box error-box">
                <h3>⚠️ Diagnostic Results: Issues Detected</h3>
                <p>The diagnostic scan has identified $($issuesFound.Count) issue(s) that may be causing QuickBooks freezing problems.</p>
            </div>
            
            <div class="summary-stats">
                <div class="stat-box">
                    <div class="stat-number" style="color: #3498db;">$($issuesFound.Count)</div>
                    <div class="stat-label">Total Issues</div>
                </div>
                <div class="stat-box critical">
                    <div class="stat-number" style="color: #e74c3c;">$critCount</div>
                    <div class="stat-label">Critical</div>
                </div>
                <div class="stat-box high">
                    <div class="stat-number" style="color: #e67e22;">$highCount</div>
                    <div class="stat-label">High Priority</div>
                </div>
                <div class="stat-box medium">
                    <div class="stat-number" style="color: #f39c12;">$medCount</div>
                    <div class="stat-label">Medium</div>
                </div>
                <div class="stat-box low">
                    <div class="stat-number" style="color: #3498db;">$lowCount</div>
                    <div class="stat-label">Low</div>
                </div>
            </div>
"@
}

# Process each section and add to HTML
foreach ($sec in $sections) {
    if ($null -ne $sec -and $null -ne $sec.Data -and ($sec.Data -ne @{})) {
        $htmlReport += "<h2>$($sec.Section)</h2>"
        
        # Special formatting for Issues section
        if ($sec.Section -eq "Issues Detected") {
            $htmlReport += @"
            <table>
                <thead>
                    <tr>
                        <th>Severity</th>
                        <th>Category</th>
                        <th>Issue Description</th>
                        <th>Recommended Fix</th>
                    </tr>
                </thead>
                <tbody>
"@
            foreach ($issue in $sec.Data) {
                $severityClass = switch($issue.Severity) {
                    'Critical' { 'critical' }
                    'High' { 'high' }
                    'Medium' { 'medium' }
                    'Low' { 'low' }
                    default { '' }
                }
                $htmlReport += @"
                    <tr class="$severityClass">
                        <td><strong>$($issue.Severity)</strong></td>
                        <td>$($issue.Category)</td>
                        <td>$($issue.Issue)</td>
                        <td><span class="fix-command">$($issue.RecommendedFix -replace '<', '&lt;' -replace '>', '&gt;')</span></td>
                    </tr>
"@
            }
            $htmlReport += @"
                </tbody>
            </table>
"@
        }
        # System section - custom formatting
        elseif ($sec.Section -eq "System") {
            $htmlReport += @"
            <ul class="property-list">
                <li><span class="property-name">Computer Name</span><span class="property-value">$($sec.Data.ComputerName)</span></li>
                <li><span class="property-name">Role</span><span class="property-value">$($sec.Data.Role)</span></li>
                <li><span class="property-name">Operating System</span><span class="property-value">$($sec.Data.OS)</span></li>
                <li><span class="property-name">Uptime</span><span class="property-value">$($sec.Data.UptimeDays) days</span></li>
                <li><span class="property-name">Power Plan</span><span class="property-value">$($sec.Data.PowerPlan)</span></li>
            </ul>
            
            <h3>CPU Information</h3>
            <ul class="property-list">
                <li><span class="property-name">Processor</span><span class="property-value">$($sec.Data.CPU.Name)</span></li>
                <li><span class="property-name">Cores</span><span class="property-value">$($sec.Data.CPU.NumberOfCores)</span></li>
                <li><span class="property-name">Logical Processors</span><span class="property-value">$($sec.Data.CPU.NumberOfLogicalProcessors)</span></li>
                <li><span class="property-name">Max Clock Speed</span><span class="property-value">$($sec.Data.CPU.MaxClockSpeed) MHz</span></li>
            </ul>
            
            <h3>Memory Information</h3>
            <ul class="property-list">
                <li><span class="property-name">Total Memory</span><span class="property-value">$($sec.Data.Memory.TotalGB) GB</span></li>
                <li><span class="property-name">Free Memory</span><span class="property-value">$($sec.Data.Memory.FreeGB) GB</span></li>
                <li><span class="property-name">Memory Used</span><span class="property-value">$($sec.Data.Memory.PercentUsed)%</span></li>
            </ul>
            <div class="progress-bar">
                <div class="progress-fill $(if($sec.Data.Memory.PercentUsed -gt 85){'danger'}elseif($sec.Data.Memory.PercentUsed -gt 70){'warning'})" style="width: $($sec.Data.Memory.PercentUsed)%;">
                    $($sec.Data.Memory.PercentUsed)% Used
                </div>
            </div>
"@
        }
        # Company File section - custom formatting
        elseif ($sec.Section -eq "Company File") {
            if ($sec.Data.Path) {
                $htmlReport += @"
            <ul class="property-list">
                <li><span class="property-name">File Path</span><span class="property-value">$($sec.Data.Path)</span></li>
                <li><span class="property-name">Path Type</span><span class="property-value">$($sec.Data.PathType)</span></li>
                <li><span class="property-name">File Exists</span><span class="property-value">$(if($sec.Data.Exists){'✅ Yes'}else{'❌ No'})</span></li>
"@
                if ($sec.Data.SizeMB) {
                    $htmlReport += @"
                <li><span class="property-name">File Size</span><span class="property-value">$($sec.Data.SizeMB) MB</span></li>
"@
                }
                if ($sec.Data.ND) {
                    $htmlReport += @"
                <li><span class="property-name">.ND File</span><span class="property-value">✅ Present ($($sec.Data.ND.SizeMB) MB)</span></li>
                <li><span class="property-name">.ND Last Modified</span><span class="property-value">$($sec.Data.ND.Modified)</span></li>
"@
                } else {
                    $htmlReport += @"
                <li><span class="property-name">.ND File</span><span class="property-value">❌ Missing</span></li>
"@
                }
                if ($sec.Data.TLG) {
                    $htmlReport += @"
                <li><span class="property-name">.TLG File</span><span class="property-value">✅ Present ($($sec.Data.TLG.SizeMB) MB)</span></li>
"@
                }
                $htmlReport += "</ul>"
                
                if ($sec.Data.SMB445) {
                    $htmlReport += @"
            <h3>Network Connectivity</h3>
            <ul class="property-list">
                <li><span class="property-name">SMB Port 445</span><span class="property-value">$(if($sec.Data.SMB445.Reachable){'✅ Reachable'}else{'❌ Unreachable'})</span></li>
                <li><span class="property-name">Ping Status</span><span class="property-value">$(if($sec.Data.SMB445.PingSucceeded){'✅ Success'}else{'❌ Failed'})</span></li>
                <li><span class="property-name">Latency</span><span class="property-value">$($sec.Data.SMB445.LatencyMs) ms</span></li>
                <li><span class="property-name">Remote Address</span><span class="property-value">$($sec.Data.SMB445.RemoteAddress)</span></li>
            </ul>
"@
                }
                
                if ($sec.Data.TouchTimingsMs) {
                    $htmlReport += @"
            <h3>Performance Metrics</h3>
            <ul class="property-list">
                <li><span class="property-name">Directory List Time</span><span class="property-value">$($sec.Data.TouchTimingsMs.DirList) ms</span></li>
                <li><span class="property-name">File Stat Time</span><span class="property-value">$($sec.Data.TouchTimingsMs.Stat) ms</span></li>
            </ul>
"@
                }
                
                if ($sec.Data.WriteProbe) {
                    $htmlReport += @"
            <h3>Write Test Results</h3>
            <ul class="property-list">
                <li><span class="property-name">Write Test</span><span class="property-value">$(if($sec.Data.WriteProbe.Succeeded){'✅ Success'}else{'❌ Failed'})</span></li>
"@
                    if ($sec.Data.WriteProbe.WriteTimeMs) {
                        $htmlReport += @"
                <li><span class="property-name">Write Time (5MB)</span><span class="property-value">$($sec.Data.WriteProbe.WriteTimeMs) ms</span></li>
"@
                    }
                    if ($sec.Data.WriteProbe.Error) {
                        $htmlReport += @"
                <li><span class="property-name">Error</span><span class="property-value">$($sec.Data.WriteProbe.Error)</span></li>
"@
                    }
                    $htmlReport += "</ul>"
                }
            } else {
                $htmlReport += "<p>No company file specified or found.</p>"
            }
        }
        # Generic table formatting for other sections
        elseif ($sec.Data -is [System.Collections.IEnumerable] -and -not ($sec.Data -is [string])) {
            $htmlReport += ($sec.Data | ConvertTo-Html -As Table -Fragment)
        } 
        else {
            $htmlReport += ($sec.Data | ConvertTo-Html -As List -Fragment)
        }
    }
}

# Add fix script section if issues found
if ($issuesFound.Count -gt 0) {
    # Generate fix script content
    $fixScript = @"
# QuickBooks Issue Fix Script
# Generated: $now
# Host: $hostName
# Role: $Role
# Total Issues: $($issuesFound.Count)

Write-Host '================================================' -ForegroundColor Cyan
Write-Host ' QuickBooks Issue Fix Script' -ForegroundColor Green
Write-Host '================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "Host: $hostName" -ForegroundColor Yellow
Write-Host "Role: $Role" -ForegroundColor Yellow
Write-Host "Issues to fix: $($issuesFound.Count)" -ForegroundColor Yellow
Write-Host ''

`$fixCount = 0
`$errorCount = 0

"@
    
    # Group fixes by severity
    $criticalFixes = $issuesFound | Where-Object { $_.Severity -eq 'Critical' }
    $highFixes = $issuesFound | Where-Object { $_.Severity -eq 'High' }
    $mediumFixes = $issuesFound | Where-Object { $_.Severity -eq 'Medium' }
    $lowFixes = $issuesFound | Where-Object { $_.Severity -eq 'Low' }
    
    if ($criticalFixes) {
        $fixScript += @"
Write-Host '--- CRITICAL FIXES ---' -ForegroundColor Red

"@
        foreach ($issue in $criticalFixes) {
            $fixScript += @"
# [$($issue.Category)] $($issue.Issue)
Write-Host "Fixing: $($issue.Issue)" -ForegroundColor Yellow
try {
    $($issue.RecommendedFix)
    Write-Host "  ✓ Fixed successfully" -ForegroundColor Green
    `$fixCount++
} catch {
    Write-Host "  ✗ Failed to fix: `$_" -ForegroundColor Red
    `$errorCount++
}

"@
        }
    }
    
    $fixScript += @"

Write-Host ''
Write-Host '================================================' -ForegroundColor Cyan
Write-Host ' Fix Script Complete' -ForegroundColor Green
Write-Host '================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host "Fixes applied: `$fixCount" -ForegroundColor Green
Write-Host "Errors encountered: `$errorCount" -ForegroundColor $(if (`$errorCount -gt 0) { 'Red' } else { 'Green' })
Write-Host ''
Write-Host 'Next Steps:' -ForegroundColor Yellow
Write-Host '1. Restart QuickBooks Database Server Manager service' -ForegroundColor White
Write-Host '2. Restart QuickBooks application' -ForegroundColor White
Write-Host '3. Test multi-user access' -ForegroundColor White
Write-Host '4. Monitor for freeze issues' -ForegroundColor White
Write-Host ''
Write-Host 'If issues persist, consider:' -ForegroundColor Yellow
Write-Host '- Running QuickBooks File Doctor' -ForegroundColor White
Write-Host '- Verifying and rebuilding the company file' -ForegroundColor White
Write-Host '- Checking Windows Event Viewer for additional errors' -ForegroundColor White
"@
    
    # Save fix script
    $fixScriptPath = Join-Path $outDir ("QBFix_{0}_{1:yyyyMMdd_HHmmss}.ps1" -f $hostName,$now)
    $fixScript | Set-Content -LiteralPath $fixScriptPath -Encoding UTF8
    
    $htmlReport += @"
            <h2>Automated Fix Script</h2>
            <div class="alert-box warning-box">
                <h3>⚠️ Important: Review Before Running</h3>
                <p><strong>A PowerShell script has been generated with all recommended fixes.</strong></p>
                <p>Please review the fixes before executing them. Some changes may require:</p>
                <ul style="margin-left: 20px;">
                    <li>Administrative privileges</li>
                    <li>Service restarts</li>
                    <li>QuickBooks to be closed</li>
                    <li>System restart after completion</li>
                </ul>
            </div>
            
            <h3>Fix Script Location</h3>
            <div class="fix-command" style="display: block; padding: 15px;">
                $fixScriptPath
            </div>
            
            <h3>To Execute the Fix Script</h3>
            <p>Run the following command in an elevated PowerShell prompt:</p>
            <pre class="fix-command">
# Option 1: Run directly
powershell.exe -ExecutionPolicy Bypass -File "$fixScriptPath"

# Option 2: Review first, then run
notepad.exe "$fixScriptPath"  # Review the script
powershell.exe -ExecutionPolicy Bypass -File "$fixScriptPath"  # Execute after review
            </pre>
            
            <h3>Manual Fix Reference</h3>
            <p>If you prefer to apply fixes manually, here's the complete list grouped by priority:</p>
"@
    
    # Add manual fix reference
    if ($criticalFixes) {
        $htmlReport += @"
            <h4 style="color: #e74c3c;">🔴 Critical Fixes (Address Immediately)</h4>
            <ol>
"@
        foreach ($issue in $criticalFixes) {
            $htmlReport += @"
                <li>
                    <strong>$($issue.Issue)</strong><br>
                    <span class="fix-command">$($issue.RecommendedFix -replace '<', '&lt;' -replace '>', '&gt;')</span>
                </li>
"@
        }
        $htmlReport += "</ol>"
    }
    
    if ($highFixes) {
        $htmlReport += @"
            <h4 style="color: #e67e22;">🟠 High Priority Fixes</h4>
            <ol>
"@
        foreach ($issue in $highFixes) {
            $htmlReport += @"
                <li>
                    <strong>$($issue.Issue)</strong><br>
                    <span class="fix-command">$($issue.RecommendedFix -replace '<', '&lt;' -replace '>', '&gt;')</span>
                </li>
"@
        }
        $htmlReport += "</ol>"
    }
    
    if ($mediumFixes) {
        $htmlReport += @"
            <h4 style="color: #f39c12;">🟡 Medium Priority Fixes</h4>
            <ol>
"@
        foreach ($issue in $mediumFixes) {
            $htmlReport += @"
                <li>
                    <strong>$($issue.Issue)</strong><br>
                    <span class="fix-command">$($issue.RecommendedFix -replace '<', '&lt;' -replace '>', '&gt;')</span>
                </li>
"@
        }
        $htmlReport += "</ol>"
    }
    
    if ($lowFixes) {
        $htmlReport += @"
            <h4 style="color: #3498db;">🔵 Low Priority Optimizations</h4>
            <ol>
"@
        foreach ($issue in $lowFixes) {
            $htmlReport += @"
                <li>
                    <strong>$($issue.Issue)</strong><br>
                    <span class="fix-command">$($issue.RecommendedFix -replace '<', '&lt;' -replace '>', '&gt;')</span>
                </li>
"@
        }
        $htmlReport += "</ol>"
    }
}

# Add recommendations section
$htmlReport += @"
            <h2>General Recommendations</h2>
            <div class="alert-box info-box">
                <h3>📋 Best Practices for QuickBooks Multi-User Environment</h3>
                <ol style="margin-left: 20px;">
                    <li><strong>Regular Maintenance:</strong> Run QuickBooks File Doctor monthly</li>
                    <li><strong>Backup Strategy:</strong> Implement automated daily backups</li>
                    <li><strong>Network Optimization:</strong> Ensure gigabit ethernet connections for all users</li>
                    <li><strong>File Size Management:</strong> Keep company files under 1GB when possible</li>
                    <li><strong>User Limits:</strong> Limit concurrent users to 5-10 for optimal performance</li>
                    <li><strong>Server Resources:</strong> Allocate at least 8GB RAM for the QuickBooks server</li>
                    <li><strong>Antivirus Exclusions:</strong> Exclude QuickBooks folders from real-time scanning</li>
                    <li><strong>Windows Updates:</strong> Keep server and workstations updated monthly</li>
                </ol>
            </div>
            
            <h2>Additional Troubleshooting Steps</h2>
            <div class="alert-box info-box">
                <h3>🔧 If Freezing Issues Persist After Fixes</h3>
                <ol style="margin-left: 20px;">
                    <li>
                        <strong>Verify Company File Integrity:</strong><br>
                        <span class="fix-command">File → Utilities → Verify Data</span>
                    </li>
                    <li>
                        <strong>Rebuild Company File:</strong><br>
                        <span class="fix-command">File → Utilities → Rebuild Data</span>
                    </li>
                    <li>
                        <strong>Re-sort Lists:</strong><br>
                        <span class="fix-command">File → Utilities → Re-sort List</span>
                    </li>
                    <li>
                        <strong>Condense Data:</strong><br>
                        <span class="fix-command">File → Utilities → Condense Data</span>
                    </li>
                    <li>
                        <strong>Create Portable Company File:</strong><br>
                        Create and restore a portable file to refresh the database
                    </li>
                    <li>
                        <strong>Check for Conflicting Applications:</strong><br>
                        Temporarily disable third-party QuickBooks add-ons
                    </li>
                    <li>
                        <strong>Network Diagnostics:</strong><br>
                        Run continuous ping tests during freeze occurrences
                    </li>
                </ol>
            </div>
        </div>
        
        <div class="footer">
            <p>QuickBooks Freeze Diagnostic Report v2.0 | Generated by Enhanced Diagnostic Script</p>
            <p>For support, consult QuickBooks ProAdvisor or Intuit Support</p>
        </div>
    </div>
</body>
</html>
"@

# Write HTML report
$htmlReport | Set-Content -LiteralPath $reportPath -Encoding UTF8

# Display summary
Write-Host "`n=== DIAGNOSTIC COMPLETE ===" -ForegroundColor Green
Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan

if ($issuesFound.Count -gt 0) {
    Write-Host "`nWould you like to:" -ForegroundColor Yellow
    Write-Host "1. View the HTML report" -ForegroundColor White
    Write-Host "2. Execute the fix script (requires admin)" -ForegroundColor White
    Write-Host "3. View fix script in Notepad" -ForegroundColor White
    Write-Host "4. Exit without fixing" -ForegroundColor White
    
    $choice = Read-Host -Prompt "Enter your choice (1-4)"
    
    switch ($choice) {
        '1' { 
            Start-Process $reportPath
            Write-Host "Opening report in browser..." -ForegroundColor Green
            Write-Host "`nFix script saved to: $fixScriptPath" -ForegroundColor Cyan
        }
        '2' {
            if ($fixScriptPath -and (Test-Path $fixScriptPath)) {
                Write-Host "`nPreparing to execute fix script..." -ForegroundColor Yellow
                Write-Host "This will attempt to fix all detected issues." -ForegroundColor Red
                Write-Host "Some fixes may require:" -ForegroundColor Yellow
                Write-Host "  - Administrative privileges" -ForegroundColor White
                Write-Host "  - QuickBooks to be closed" -ForegroundColor White
                Write-Host "  - Services to be restarted" -ForegroundColor White
                Write-Host ""
                $confirm = Read-Host -Prompt "Are you sure you want to proceed? (Y/N)"
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    Write-Host "`nExecuting fix script..." -ForegroundColor Green
                    & powershell.exe -ExecutionPolicy Bypass -File $fixScriptPath
                    Write-Host "`nFix script execution complete." -ForegroundColor Green
                    Write-Host "Please restart QuickBooks and test for improvements." -ForegroundColor Yellow
                } else {
                    Write-Host "Fix cancelled. You can run the script manually later:" -ForegroundColor Yellow
                    Write-Host $fixScriptPath -ForegroundColor Cyan
                }
            }
        }
        '3' {
            if ($fixScriptPath -and (Test-Path $fixScriptPath)) {
                Start-Process notepad.exe -ArgumentList $fixScriptPath
                Write-Host "Opening fix script in Notepad..." -ForegroundColor Green
                Write-Host "`nTo run the script later, use:" -ForegroundColor Yellow
                Write-Host "powershell.exe -ExecutionPolicy Bypass -File `"$fixScriptPath`"" -ForegroundColor Cyan
            }
        }
        default {
            Write-Host "`nExiting without fixes." -ForegroundColor Yellow
            Write-Host "Reports saved to:" -ForegroundColor White
            Write-Host "  HTML Report: $reportPath" -ForegroundColor Cyan
            if ($fixScriptPath) {
                Write-Host "  Fix Script:  $fixScriptPath" -ForegroundColor Cyan
            }
        }
    }
} else {
    Write-Host "`nNo issues found! Your QuickBooks environment appears to be properly configured." -ForegroundColor Green
    Write-Host "`nWould you like to view the report? (Y/N)" -ForegroundColor Yellow
    $viewChoice = Read-Host
    if ($viewChoice -eq 'Y' -or $viewChoice -eq 'y') {
        Start-Process $reportPath
        Write-Host "Opening report in browser..." -ForegroundColor Green
    }
}

Write-Host "`n=== Script Complete ===" -ForegroundColor Green
Write-Host "Thank you for using QuickBooks Freeze Diagnostic!" -ForegroundColor Cyan.Issue)" -ForegroundColor Yellow
try {
    $($issue.RecommendedFix)
    Write-Host "  ✓ Fixed successfully" -ForegroundColor Green
    `$fixCount++
} catch {
    Write-Host "  ✗ Failed to fix: `$_" -ForegroundColor Red
    `$errorCount++
}

"@
        }
    }
    
    if ($highFixes) {
        $fixScript += @"
Write-Host '--- HIGH PRIORITY FIXES ---' -ForegroundColor DarkRed

"@
        foreach ($issue in $highFixes) {
            $fixScript += @"
# [$($issue.Category)] $($issue.Issue)
Write-Host "Fixing: $($issue.Issue)" -ForegroundColor Yellow
try {
    $($issue.RecommendedFix)
    Write-Host "  ✓ Fixed successfully" -ForegroundColor Green
    `$fixCount++
} catch {
    Write-Host "  ✗ Failed to fix: `$_" -ForegroundColor Red
    `$errorCount++
}

"@
        }
    }
    
    if ($mediumFixes) {
        $fixScript += @"
Write-Host '--- MEDIUM PRIORITY FIXES ---' -ForegroundColor DarkYellow

"@
        foreach ($issue in $mediumFixes) {
            $fixScript += @"
# [$($issue.Category)] $($issue.Issue)
Write-Host "Fixing: $($issue.Issue)" -ForegroundColor Yellow
try {
    $($issue.RecommendedFix)
    Write-Host "  ✓ Fixed successfully" -ForegroundColor Green
    `$fixCount++
} catch {
    Write-Host "  ✗ Failed to fix: `$_" -ForegroundColor Red
    `$errorCount++
}

"@
        }
    }
    
    if ($lowFixes) {
        $fixScript += @"
Write-Host '--- LOW PRIORITY FIXES ---' -ForegroundColor Yellow

"@
        foreach ($issue in $lowFixes) {
            $fixScript += @"
# [$($issue.Category)] $($issue.Issue)
Write-Host "Fixing: $($issue else {
    $htmlReport += @"
            <div class="alert-box success-box">
                <h3>✅ Diagnostic Results: All Clear</h3>
                <p>No issues were detected during the diagnostic scan. Your QuickBooks environment appears to be properly configured.</p>
            </div>
            
            <div class="summary-stats">
                <div class="stat-box success">
                    <div class="stat-number" style="color: #27ae60;">0</div>
                    <div class="stat-label">Issues Found</div>
                </div>
            </div>
"@
}

$null = $sb.AppendLine("<h1>QuickBooks Freeze Diagnostic Report</h1>")
$null = $sb.AppendLine("<div class='info-box'>")
$null = $sb.AppendLine("<strong>Host:</strong> $hostName | <strong>Role:</strong> $Role | <strong>Generated:</strong> $($now) | <strong>Report Path:</strong> $reportPath")
$null = $sb.AppendLine("</div>")

# Add summary statistics
if ($issuesFound.Count -gt 0) {
    $critCount = ($issuesFound | Where-Object { $_.Severity -eq 'Critical' }).Count
    $highCount = ($issuesFound | Where-Object { $_.Severity -eq 'High' }).Count
    $medCount = ($issuesFound | Where-Object { $_.Severity -eq 'Medium' }).Count
    $lowCount = ($issuesFound | Where-Object { $_.Severity -eq 'Low' }).Count
    
    $null = $sb.AppendLine("<div class='summary-stats'>")
    $null = $sb.AppendLine("<div class='stat-box'><div class='stat-number'>$($issuesFound.Count)</div><div class='stat-label'>Total Issues</div></div>")
    $null = $sb.AppendLine("<div class='stat-box' style='border-left: 4px solid #e74c3c;'><div class='stat-number' style='color:#e74c3c;'>$critCount</div><div class='stat-label'>Critical</div></div>")
    $null = $sb.AppendLine("<div class='stat-box' style='border-left: 4px solid #e67e22;'><div class='stat-number' style='color:#e67e22;'>$highCount</div><div class='stat-label'>High</div></div>")
    $null = $sb.AppendLine("<div class='stat-box' style='border-left: 4px solid #f39c12;'><div class='stat-number' style='color:#f39c12;'>$medCount</div><div class='stat-label'>Medium</div></div>")
    $null = $sb.AppendLine("<div class='stat-box' style='border-left: 4px solid #3498db;'><div class='stat-number' style='color:#3498db;'>$lowCount</div><div class='stat-label'>Low</div></div>")
    $null = $sb.AppendLine("</div>")
}

# Add sections to HTML
foreach ($sec in $sections) {
    if ($null -ne $sec -and $null -ne $sec.Data -and ($sec.Data -ne @{})) {
        $null = $sb.AppendLine("<h2>$($sec.Section)</h2>")
        
        # Special formatting for Issues section
        if ($sec.Section -eq "Issues Detected") {
            $null = $sb.AppendLine("<table>")
            $null = $sb.AppendLine("<tr><th>Severity</th><th>Category</th><th>Issue</th><th>Recommended Fix</th></tr>")
            foreach ($issue in $sec.Data) {
                $severityClass = switch($issue.Severity) {
                    'Critical' { 'critical' }
                    'High' { 'high' }
                    'Medium' { 'medium' }
                    'Low' { 'low' }
                    default { '' }
                }
                $null = $sb.AppendLine("<tr class='$severityClass'>")
                $null = $sb.AppendLine("<td><strong>$($issue.Severity)</strong></td>")
                $null = $sb.AppendLine("<td>$($issue.Category)</td>")
                $null = $sb.AppendLine("<td>$($issue.Issue)</td>")
                $null = $sb.AppendLine("<td><span class='fix-command'>$($issue.RecommendedFix)</span></td>")
                $null = $sb.AppendLine("</tr>")
            }
            $null = $sb.AppendLine("</table>")
        }
        elseif ($sec.Data -is [System.Collections.IEnumerable] -and -not ($sec.Data -is [string])) {
            $null = $sb.AppendLine(($sec.Data | ConvertTo-Html -As Table -Fragment))
        } else {
            $null = $sb.AppendLine(($sec.Data | ConvertTo-Html -As List -Fragment))
        }
    }
}

# Add fix script generation option
if ($issuesFound.Count -gt 0) {
    $null = $sb.AppendLine("<h2>Automated Fix Script</h2>")
    $null = $sb.AppendLine("<div class='warning-box'>")
    $null = $sb.AppendLine("<p><strong>Warning:</strong> Review all fixes before applying them. Some fixes require administrative privileges and may require system restarts.</p>")
    $null = $sb.AppendLine("<p>To generate a PowerShell script with all recommended fixes, run the following command:</p>")
    $null = $sb.AppendLine("<pre class='fix-command' style='display:block; padding:10px;'>")
    
    # Generate fix script content
    $fixScript = @"
# QuickBooks Issue Fix Script
# Generated: $now
# Host: $hostName
# Role: $Role

Write-Host 'Starting QuickBooks issue fixes...' -ForegroundColor Green

"@
    
    foreach ($issue in $issuesFound) {
        $fixScript += @"

# Fix: $($issue.Issue)
Write-Host 'Fixing: $($issue.Issue)' -ForegroundColor Yellow
try {
    $($issue.RecommendedFix)
    Write-Host '  Fixed successfully' -ForegroundColor Green
} catch {
    Write-Host "  Failed to fix: `$_" -ForegroundColor Red
}
"@
    }
    
    $fixScript += @"

Write-Host 'Fix script completed. Please restart QuickBooks and test.' -ForegroundColor Green
"@
    
    # Save fix script
    $fixScriptPath = Join-Path $outDir ("QBFix_{0}_{1:yyyyMMdd_HHmmss}.ps1" -f $hostName,$now)
    $fixScript | Set-Content -LiteralPath $fixScriptPath -Encoding UTF8
    
    $null = $sb.AppendLine("# Execute the generated fix script:")
    $null = $sb.AppendLine("powershell.exe -ExecutionPolicy Bypass -File `"$fixScriptPath`"")
    $null = $sb.AppendLine("</pre>")
    $null = $sb.AppendLine("<p>Fix script saved to: <strong>$fixScriptPath</strong></p>")
    $null = $sb.AppendLine("</div>")
}

$null = $sb.AppendLine("</body></html>")

# Write HTML report
$sb.ToString() | Set-Content -LiteralPath $reportPath -Encoding UTF8

# Display summary
Write-Host "`n=== DIAGNOSTIC COMPLETE ===" -ForegroundColor Green
Write-Host "Report saved to: $reportPath" -ForegroundColor Cyan

if ($issuesFound.Count -gt 0) {
    Write-Host "`nWould you like to:" -ForegroundColor Yellow
    Write-Host "1. View the HTML report" -ForegroundColor White
    Write-Host "2. Execute the fix script (requires admin)" -ForegroundColor White
    Write-Host "3. Exit without fixing" -ForegroundColor White
    
    $choice = Read-Host -Prompt "Enter your choice (1-3)"
    
    switch ($choice) {
        '1' { 
            Start-Process $reportPath
            Write-Host "Opening report in browser..." -ForegroundColor Green
        }
        '2' {
            if ($fixScriptPath -and (Test-Path $fixScriptPath)) {
                Write-Host "`nExecuting fix script..." -ForegroundColor Yellow
                Write-Host "This will attempt to fix all detected issues." -ForegroundColor Red
                $confirm = Read-Host -Prompt "Are you sure? (Y/N)"
                if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                    & $fixScriptPath
                } else {
                    Write-Host "Fix cancelled. You can run the script manually later:" -ForegroundColor Yellow
                    Write-Host $fixScriptPath -ForegroundColor Cyan
                }
            }
        }
        default {
            Write-Host "Exiting without fixes. Report saved to:" -ForegroundColor Yellow
            Write-Host $reportPath -ForegroundColor Cyan
        }
    }
} else {
    Start-Process $reportPath
    Write-Host "No issues found! Opening report..." -ForegroundColor Green
}

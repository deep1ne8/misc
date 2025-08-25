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

$ErrorActionPreference = 'Stop'
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

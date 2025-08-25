<# 
.SYNOPSIS
  QuickBooks Freeze Diagnostic for split-role environments (RDS RemoteApp + QB DB/File Server)
.DESCRIPTION
  Collects OS/storage/network signals, QB service/process/version info, RDS printing flags,
  inspects .ND/.TLG, validates SMB path/latency, samples relevant Event Logs, and outputs HTML.
  Safe (read-only) by default. Optional tiny write probe in the share directory.
.NOTES
  Tested on PS 5.1+. Run on both the RDS host and the file server for comparison.
#>

[CmdletBinding()]
param(
  [ValidateSet('Server','RDSHost','Client')]
  [string]$Role = 'Client',

  [string]$CompanyFilePath,
  [switch]$FindCompanyFile,
  [int]$LogDays = 7,
  [switch]$DoShareWriteProbe
)

$ErrorActionPreference = 'Stop'
$now = Get-Date
$hostName = $env:COMPUTERNAME
$outDir = 'C:\ProgramData\QB-Diag'
$null = New-Item -ItemType Directory -Path $outDir -Force -ErrorAction SilentlyContinue
$reportPath = Join-Path $outDir ("QBDiag_{0}_{1:yyyyMMdd_HHmmss}.html" -f $hostName,$now)
$Server = "Read-Host -Prompt Enter Server Name:"
$RDSHost = "Read-Host -Prompt Enter RDSHost Name:"
$Client = "Read-Host -Prompt Enter Client"
$QuickBooksDir1 = 'C:\Users\Public'
$QuickBooksDir2 = 'C:\QuickBooks'
$QuickBooksDir3 = "Read-Host -Prompt 'Enter your Quickbooks directory':"

function New-Section { param([string]$Title,[object]$Data) [PSCustomObject]@{Section=$Title;Data=$Data} }
function SizeMB { param([long]$b) if ($b -is [long]) {[math]::Round($b/1MB,2)} else {$null} }
function BaseNameNoExt([string]$p){ [IO.Path]::GetFileNameWithoutExtension($p) }

function Resolve-CompanyPath {
  param([string]$Path,[switch]$AllowSearch)
  if ($Path) { return $Path }
  if ($AllowSearch) {
    try {
      $candidates = New-Object System.Collections.Generic.List[string]
      Get-ChildItem "$env:ProgramData\Intuit" -Recurse -Filter 'QBWUSER.INI' -ErrorAction SilentlyContinue | ForEach-Object {
        foreach ($line in (Get-Content -LiteralPath $_.FullName -ErrorAction SilentlyContinue)) {
          if ($line -match '\.qbw' -and $line -match '=') {
            $p = ($line -split '=',2)[1].Trim()
            if ($p -and (Test-Path -LiteralPath $p)) { $candidates.Add($p) }
          }
        }
      }
      foreach ($root in @('$QuickBooksDir1','$QuickBooksDir2','$QuickBooksDir3')) {
        if (Test-Path $root) {
          Get-ChildItem -Path $root -Recurse -Filter *.qbw -ErrorAction SilentlyContinue -Force |
            Select-Object -First 2 -Expand FullName | ForEach-Object { $candidates.Add($_) }
        }
      }
      return ($candidates | Select-Object -Unique | Select-Object -First 1)
    } catch { }
  }
  return $null
}

function Parse-ServerFromUNC([string]$UNC){
  if ($UNC -and $UNC.StartsWith('\\')) { return $UNC.TrimStart('\').Split('\')[0] } else { return $null }
}

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
$sysSection = New-Section "System" (@{
  Role         = $Role
  ComputerName = $hostName
  OS           = "$($os.Caption) ($($os.Version))"
  UptimeDays   = [Math]::Round($uptime.TotalDays,1)
  CPU          = $cpu
  Memory       = $mem
  PowerPlan    = (powercfg /getactivescheme 2>$null) -join ' '
})

# --- Storage + SMART-ish
$vols = Get-Volume -ErrorAction SilentlyContinue | Select-Object DriveLetter,Path,FileSystem,Size,SizeRemaining,HealthStatus
$pd   = Get-PhysicalDisk -ErrorAction SilentlyContinue | Select-Object FriendlyName,MediaType,Size,HealthStatus
$diskSection = New-Section "Storage" (@{ Volumes = $vols; PhysicalDisks = $pd })

# --- QB services/processes/versions
$qbServices = Get-Service -Name 'QBDBMgr*','QBCFMonitorService' -ErrorAction SilentlyContinue |
  Select-Object Name,Status,StartType,DisplayName
$qbProc = Get-Process -Name 'QBW32','QBDBMgrN','QBCFMonitorService' -ErrorAction SilentlyContinue |
  Select-Object ProcessName,Id,CPU,@{n='WorkingSetMB';e={[math]::Round($_.WS/1MB,1)}},StartTime

# discover executable versions (best effort)
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
    $nd  = "$parent\$(BaseNameNoExt $CompanyFilePath).ND"
    $tlg = [IO.Path]::ChangeExtension($CompanyFilePath,'.TLG')
    $companyObj.ND  = (Test-Path $nd)  ? @{Path=$nd;  SizeMB=SizeMB (Get-Item $nd).Length;  Modified=(Get-Item $nd).LastWriteTime} : $null
    $companyObj.TLG = (Test-Path $tlg) ? @{Path=$tlg; SizeMB=SizeMB (Get-Item $tlg).Length; Modified=(Get-Item $tlg).LastWriteTime} : $null
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
  }

  # measure a few metadata touches
  try {
    $t1 = Measure-Command { Get-ChildItem -LiteralPath $parent | Select-Object -First 1 | Out-Null }
    $t2 = Measure-Command { (Get-Item -LiteralPath $CompanyFilePath).LastWriteTime | Out-Null }
    $companyObj.TouchTimingsMs = @{
      DirList = [math]::Round($t1.TotalMilliseconds,1)
      Stat    = [math]::Round($t2.TotalMilliseconds,1)
    }
  } catch {}

  # optional tiny write probe
  if ($DoShareWriteProbe.IsPresent) {
    try {
      $tmp = Join-Path $parent ("_qbdiag_{0}.bin" -f [guid]::NewGuid().ToString('N'))
      $sw  = [IO.File]::Create($tmp)
      $buf = New-Object byte[] (5MB)
      (New-Object System.Random).NextBytes($buf)
      $sw.Write($buf,0,$buf.Length); $sw.Flush(); $sw.Dispose()
      $wr = Get-Item $tmp
      Remove-Item $tmp -Force
      $companyObj.WriteProbe = @{Succeeded=$true; SizeMB=SizeMB $wr.Length}
    } catch {
      $companyObj.WriteProbe = @{Succeeded=$false; Error=$_.Exception.Message}
    }
  }
}

# --- File server share properties (if local path on a server)
$shareInfo = $null
if ($Role -eq 'Server' -and $CompanyFilePath) {
  try {
    # if given local path, map it to share; if UNC, map share directly
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
  } catch {}
}

# --- RDS printing flags (only on RDS host)
$rdsPrint = $null
if ($Role -eq 'RDSHost') {
  try {
    $key = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Terminal Services'
    $easy = (Get-ItemProperty -Path $key -ErrorAction SilentlyContinue).UseTerminalServicesEasyPrint
    $spool = Get-Service -Name Spooler -ErrorAction SilentlyContinue | Select-Object Status,StartType
    $rdsPrint = @{
      EasyPrintPolicy = $(if ($easy -eq 1) {'Enabled'} elseif ($easy -eq 0) {'Disabled'} else {'NotConfigured'})
      Spooler         = $spool
      PrinterDrivers  = (Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object Name,Manufacturer,Version)
    }
  } catch {}
}

# --- Event logs (last X days): QB, SMB, Disk, App Errors
$startTime = (Get-Date).AddDays(-1 * $LogDays)
$providers = @(
  'QuickBooks','QBDBMgrN','QBCFMonitorService','Application Error',
  'Microsoft-Windows-SMBClient','Microsoft-Windows-SMBServer',
  'Microsoft-Windows-PrintService','disk','Ntfs','srv','srv2','LanmanWorkstation','LanmanServer'
)
$logEvents = @()
foreach ($p in $providers) {
  try {
    $ev = Get-WinEvent -FilterHashtable @{StartTime=$startTime; ProviderName=$p} -ErrorAction SilentlyContinue |
      Select-Object TimeCreated,ProviderName,Id,LevelDisplayName,Message -First 60
    if ($ev) {
      $logEvents += $ev
    }
  } catch {}
}
$logsSection = New-Section ("Event Logs (last {0} days)" -f $LogDays) ($logEvents)

# --- Build HTML
$sections = @(
  $sysSection,
  $diskSection,
  $qbSection,
  (New-Section "Company File" $companyObj),
  (New-Section "Share (server-side)" $shareInfo),
  (New-Section "RDS Printing" $rdsPrint),
  $logsSection
)

# quick HTML renderer
$sb = New-Object System.Text.StringBuilder
$null = $sb.AppendLine("<html><head><meta charset='utf-8'><title>QB Freeze Diagnostic - $hostName</title>")
$null = $sb.AppendLine("<style>body{font-family:Segoe UI,Arial;font-size:12px} table{border-collapse:collapse;margin:8px 0} th,td{border:1px solid #ddd;padding:6px} th{background:#f3f3f3;text-align:left}</style></head><body>")
$null = $sb.AppendLine("<h2>QuickBooks Freeze Diagnostic</h2><p><b>Host:</b> $hostName &nbsp; <b>Role:</b> $Role &nbsp; <b>Generated:</b> $($now)</p>")

foreach ($sec in $sections) {
  if ($null -ne $sec -and $null -ne $sec.Data -and ($sec.Data -ne @{})) {
    $null = $sb.AppendLine("<h3>$($sec.Section)</h3>")
    if ($sec.Data -is [System.Collections.IEnumerable] -and -not ($sec.Data -is [string])) {
      $null = $sb.AppendLine(($sec.Data | ConvertTo-Html -As Table -Fragment))
    } else {
      $null = $sb.AppendLine(($sec.Data | ConvertTo-Html -As List -Fragment))
    }
  }
}
$null = $sb.AppendLine("</body></html>")
$sb.ToString() | Set-Content -LiteralPath $reportPath -Encoding UTF8

Write-Host "Report written to: $reportPath"

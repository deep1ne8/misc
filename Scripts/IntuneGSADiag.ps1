<# 
Tier-0 GSA/Intune Deep-Diag
- Collects signals that commonly break Global Secure Access after upgrades/profile sign-in.
- Output: JSON + TXT summary + optional ZIP of logs.
- Safe to run on Win10/11. No external modules. No tenant calls.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [switch]$CollectLogs,                   # Include raw logs in ZIP
    [int]$EventHours = 24,                  # How many hours of logs to pull
    [string]$OutDir = "C:\Windows\Temp\GSA_Diag"
)

#region Setup
$ErrorActionPreference = "Stop"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$stamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$txt = Join-Path $OutDir "GSA_Diag_$stamp.txt"
$json = Join-Path $OutDir "GSA_Diag_$stamp.json"
$zip  = Join-Path $OutDir "GSA_Diag_$stamp.zip"
Write-Output "=== Tier-0 GSA/Intune Diagnostic ($stamp) ===" | Tee-Object -FilePath $txt -Append | Out-Null

function Add-Note { param([string]$m) $m | Tee-Object -FilePath $txt -Append | Out-Null }

function New-Finding {
    param(
        [string]$Area,
        [ValidateSet('Pass','Warn','Fail','Info')] [string]$Status,
        [string]$Message,
        [hashtable]$Details
    )
    [pscustomobject]@{
        Area    = $Area
        Status  = $Status
        Message = $Message
        Details = $Details
    }
}

$results = New-Object System.Collections.Generic.List[object]
$since = (Get-Date).AddHours(-1 * [math]::Abs($EventHours))

Add-Note "Output directory: $OutDir"
Add-Note "Log window: last $EventHours hour(s), since $since"
#endregion

#region System & OS
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $u  = whoami
    $build = [int]$os.BuildNumber
    $isWin11 = $build -ge 22000

    $results.Add( New-Finding -Area "OS" -Status "Info" -Message "OS detected" -Details @{
        Caption      = $os.Caption
        Version      = $os.Version
        Build        = $os.BuildNumber
        Win11        = $isWin11
        LastBoot     = $os.LastBootUpTime
        User         = $u
        Domain       = $cs.Domain
        PartOfDomain = $cs.PartOfDomain
    })
    Add-Note "OS: $($os.Caption) $($os.Version) (Build $build) | User: $u"
} catch {
    $results.Add( New-Finding -Area "OS" -Status "Warn" -Message "Failed to read OS/CS info" -Details @{ Error = "$_" })
}
#endregion

#region AAD Join & PRT (dsregcmd)
try {
    $dsOut = dsregcmd /status
    $dsText = $dsOut -join "`n"
    $ds = @{
        AzureAdJoined      = ($dsText -match 'AzureAdJoined\s*:\s*YES')
        WorkplaceJoined    = ($dsText -match 'WorkplaceJoined\s*:\s*YES')
        DomainJoined       = ($dsText -match 'DomainJoined\s*:\s*YES')
        DeviceId           = ($dsText | Select-String -Pattern 'DeviceId\s*:\s*([0-9a-f\-]{36})').Matches.Groups[1].Value
        TenantName         = ($dsText | Select-String -Pattern 'TenantName\s*:\s*(.+)$').Matches.Groups[1].Value.Trim()
        TenantId           = ($dsText | Select-String -Pattern 'TenantId\s*:\s*([0-9a-f\-]{36})').Matches.Groups[1].Value
        PRT                = (($dsText | Select-String -Pattern 'PRT\s*:\s*(YES|NO)').Matches.Groups[1].Value)
        SSOState           = ($dsText | Select-String -Pattern 'SSO State\s*:\s*(.+)$').Matches.Groups[1].Value.Trim()
        WamDefaultSet      = ($dsText -match 'WAM Default Set\s*:\s*YES')
        WamDefaultAuthority= ($dsText | Select-String -Pattern 'WAM Default Authority\s*:\s*(.+)$').Matches.Groups[1].Value.Trim()
    }

    $prtStatus = if($ds.PRT -eq 'YES'){'Pass'} else {'Fail'}
    $prtMsg = if($ds.PRT -eq 'YES'){'PRT present'} else {'PRT missing; WAM/Sign-in likely broken (check CA/MFA/Time/Proxy)'}
    $results.Add( New-Finding -Area "AAD/PRT" -Status $prtStatus -Message $prtMsg -Details $ds )
    Add-Note "AAD Joined: $($ds.AzureAdJoined) | PRT: $($ds.PRT) | Tenant: $($ds.TenantName)"
    
    Set-Content -Path (Join-Path $OutDir "dsregcmd_$stamp.txt") -Value $dsText -Encoding UTF8
} catch {
    $results.Add( New-Finding -Area "AAD/PRT" -Status "Warn" -Message "dsregcmd failed" -Details @{ Error = "$_" })
}
#endonregion

#region MDM Enrollment Keys & SCEP Cert (MS-Organization-Access)
try {
    $enrollRoot = "HKLM:\SOFTWARE\Microsoft\Enrollments"
    $enrolls = Get-ChildItem $enrollRoot -ErrorAction Stop | ForEach-Object {
        $p = $_.PSPath
        [pscustomobject]@{
            Key  = $_.PSChildName
            UPN  = (Get-ItemProperty -Path $p -Name UPN -ErrorAction SilentlyContinue).UPN
            DiscoveryServiceFullURL = (Get-ItemProperty -Path $p -Name DiscoveryServiceFullURL -ErrorAction SilentlyContinue).DiscoveryServiceFullURL
            EnrollmentType = (Get-ItemProperty -Path $p -Name EnrollmentType -ErrorAction SilentlyContinue).EnrollmentType
            LastRenewTime  = (Get-ItemProperty -Path $p -Name RenewalDate -ErrorAction SilentlyContinue).RenewalDate
        }
    }
    $results.Add( New-Finding -Area "MDM/Enrollment" -Status ($(if($enrolls){'Pass'} else {'Fail'})) -Message ($(if($enrolls){'Enrollment keys present'} else {'No enrollment keys found'})) -Details @{ Enrollments = $enrolls })
    Add-Note "MDM Enrollments found: $($enrolls.Count)"
} catch {
    $results.Add( New-Finding -Area "MDM/Enrollment" -Status "Warn" -Message "Failed to read enrollment keys" -Details @{ Error = "$_" })
}

# Cert check
try {
    $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.FriendlyName -eq "Microsoft Intune MDM Device CA" -or $_.Subject -match "CN=MS-Organization-Access" } | Sort-Object NotAfter -Descending | Select-Object -First 1
    if ($cert) {
        $exp = $cert.NotAfter
        $days = ($exp - (Get-Date)).TotalDays
        $status = if ($days -gt 15) {'Pass'} elseif ($days -gt 0) {'Warn'} else {'Fail'}
        $msg = "MDM access cert (MS-Organization-Access) expires in {0:N0} day(s)" -f $days
        $results.Add( New-Finding -Area "MDM/Cert" -Status $status -Message $msg -Details @{
            Subject=$cert.Subject; Thumbprint=$cert.Thumbprint; NotAfter=$cert.NotAfter; NotBefore=$cert.NotBefore
        })
        Add-Note "MDM cert: $($cert.Subject) | Expires: $($cert.NotAfter)"
    } else {
        $results.Add( New-Finding -Area "MDM/Cert" -Status "Fail" -Message "MDM access certificate not found" -Details @{} )
        Add-Note "MDM cert not found."
    }
} catch {
    $results.Add( New-Finding -Area "MDM/Cert" -Status "Warn" -Message "Cert lookup failed" -Details @{ Error = "$_" })
}
#endonregion

#region Intune Management Extension (IME) & MDM Engine
try {
    $svcIME = Get-Service -Name IntuneManagementExtension -ErrorAction SilentlyContinue
    $svcDMW = Get-Service -Name dmwappushservice -ErrorAction SilentlyContinue
    $svcOMCI = Get-Service -Name omadmclient -ErrorAction SilentlyContinue

    $details = @{
        IME_Service        = $svcIME.Status.ToString()
        Dmwappush          = $svcDMW.Status.ToString()
        OMADMClient        = $(if($svcOMCI){$svcOMCI.Status.ToString()}else{"N/A"})
    }
    $status = if(($svcIME -and $svcIME.Status -eq 'Running') -and ($svcDMW -and $svcDMW.Status -eq 'Running')){'Pass'} else {'Warn'}
    $results.Add( New-Finding -Area "Intune/Services" -Status $status -Message "IME/MDM services state collected" -Details $details )
    Add-Note "IME: $($details.IME_Service) | dmwappush: $($details.Dmwappush) | OMADM: $($details.OMADMClient)"
} catch {
    $results.Add( New-Finding -Area "Intune/Services" -Status "Warn" -Message "Service query failed" -Details @{ Error = "$_" })
}

# Scheduled tasks commonly used by MDM/IME
try {
    $tasks = @("PushLaunch","Schedule created by enrollment client","OMADMClient","EnterpriseMgmtTask")
    $taskStates = @{}
    foreach($t in $tasks){
        $task = Get-ScheduledTask -TaskName $t -ErrorAction SilentlyContinue
        if($task){
            $taskStates[$t] = $task.State.ToString()
        } else {
            $taskStates[$t] = "NotFound"
        }
    }
    $results.Add( New-Finding -Area "Intune/Tasks" -Status "Info" -Message "Key scheduled tasks enumerated" -Details $taskStates)
    Add-Note "Tasks: $((($taskStates.GetEnumerator() | %{"$_"} ) -join ' | '))"
} catch {
    $results.Add( New-Finding -Area "Intune/Tasks" -Status "Warn" -Message "Task query failed" -Details @{ Error = "$_" })
}
#endonregion

#region Local Accounts & Policy friction (common CSP failure: LocalUserGroup)
try {
    $localAdminGroup = Get-LocalGroupMember -Group "Administrators" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name
    $sysAdminUser = Get-LocalUser -ErrorAction SilentlyContinue | Where-Object Name -eq "SysAdmin"

    $details = @{
        AdminGroupMembers = $localAdminGroup
        SysAdminExists    = [bool]$sysAdminUser
        SysAdminEnabled   = $(if($sysAdminUser){$sysAdminUser.Enabled} else {$null})
    }
    $results.Add( New-Finding -Area "Local Accounts" -Status "Info" -Message "Local admin group & managed account snapshot" -Details $details)
    Add-Note "Administrators: $($localAdminGroup -join ', ') | SysAdmin exists: " + ([bool]$sysAdminUser)
} catch {
    $results.Add( New-Finding -Area "Local Accounts" -Status "Warn" -Message "Local group query failed" -Details @{ Error = "$_" })
}
#endonregion

#region Time/Proxy/DNS (token & GSA prerequisites)
try {
    $ntpSkew = [math]::Round(((Get-Date) - (Get-Date -AsUTC).ToLocalTime()).TotalSeconds,2)  # crude sanity
    $winhttp = netsh winhttp show proxy
    $dns = Get-DnsClientServerAddress -AddressFamily IPv4 | ForEach-Object {
        [pscustomobject]@{ InterfaceAlias=$_.InterfaceAlias; Servers=($_.ServerAddresses -join ',') }
    }
    $connectivity = @{}
    foreach($host in @("login.microsoftonline.com","device.login.microsoftonline.com","enterpriseregistration.windows.net")){
        try{
            $t = Test-NetConnection -ComputerName $host -Port 443 -InformationLevel Quiet -WarningAction SilentlyContinue
            $connectivity[$host] = $t
        } catch { $connectivity[$host] = $false }
    }

    $results.Add( New-Finding -Area "Network Prereqs" -Status "Info" -Message "Time/Proxy/DNS/connectivity snapshot" -Details @{
        ApproxTimeSkewSec = $ntpSkew
        WinHTTPProxy      = ($winhttp -join ' ')
        DNS               = $dns
        TLS443Reachable   = $connectivity
    })
    Add-Note "Connectivity 443: $((($connectivity.GetEnumerator() | % { ""+$_.Key+':'+$_.Value }) -join ' | '))"
} catch {
    $results.Add( New-Finding -Area "Network Prereqs" -Status "Warn" -Message "Prereq check failed" -Details @{ Error = "$_" })
}
#endonregion

#region Event Logs (AAD/MDM/WAM/IME/GSA-adjacent)
function Get-Log {
    param([string]$LogName,[datetime]$Since,[int]$Max=500,[string]$Level='*')
    try {
        $filter = @{LogName=$LogName; StartTime=$Since}
        $ev = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop | Sort-Object TimeCreated -Descending
        if($Level -ne '*'){ $ev = $ev | Where-Object { $_.LevelDisplayName -eq $Level } }
        return $ev | Select-Object -First $Max
    } catch { return @() }
}

$logMap = @(
    "Microsoft-Windows-AAD/Operational",
    "Microsoft-Windows-User Device Registration/Admin",
    "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin",
    "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational",
    "Application"
)

$logSummary = @{}
foreach($ln in $logMap){
    $events = Get-Log -LogName $ln -Since $since -Max 300
    $logSummary[$ln] = @{
        Total   = $events.Count
        Errors  = ($events | Where-Object LevelDisplayName -eq 'Error').Count
        Warnings= ($events | Where-Object LevelDisplayName -eq 'Warning').Count
    }
    if($CollectLogs -and $events){
        $path = Join-Path $OutDir ("{0}_{1}.evtx" -f ($ln -replace '[\\/]','_'), $stamp)
        try { wevtutil epl "$ln" "$path" /ow:true } catch {}
    }
}
$results.Add( New-Finding -Area "EventLogs" -Status "Info" -Message "Collected AAD/MDM/WAM/IME-adjacent logs" -Details $logSummary )
Add-Note "Event log counts (last $EventHours h): $(($logSummary.GetEnumerator() | % { $_.Key + '=' + ($_.Value.Errors) + 'E/' + ($_.Value.Warnings) + 'W' } ) -join ' | ')"
#endonregion

#region IME Logs tail
try {
    $imeLogDir = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
    $imeLatest = Get-ChildItem $imeLogDir -Filter "*.log" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 3
    $tails = @{}
    foreach($f in $imeLatest){
        $tails[$f.Name] = (Get-Content $f.FullName -Tail 100 -ErrorAction SilentlyContinue) -join "`n"
    }
    $results.Add( New-Finding -Area "IME/Logs" -Status "Info" -Message "Last 100 lines from 3 newest IME logs" -Details $tails )
    Add-Note "IME logs tailed: $(($imeLatest | Select-Object -ExpandProperty Name) -join ', ')"
} catch {
    $results.Add( New-Finding -Area "IME/Logs" -Status "Warn" -Message "Could not read IME logs" -Details @{ Error = "$_" })
}
#endonregion

#region Output & quick heuristics
# Heuristic flags to point you fast to likely root causes
$flags = @()

# PRT
if(($results | ? Area -eq 'AAD/PRT').Details.PRT -ne 'YES'){ $flags += "PRT missing (WAM/CA/MFA/Time/Proxy). Try: 'dsregcmd /leave' then re-join + Company Portal re-sync." }

# MDM cert
$mdmCert = ($results | ? Area -eq 'MDM/Cert')
if($mdmCert.Status -eq 'Fail'){ $flags += "MDM cert missing. Kick OMADM tasks and re-enroll." }
elseif($mdmCert.Status -eq 'Warn'){ $flags += "MDM cert close to expiry. Renew via OMADM tasks." }

# Services
$svc = ($results | ? Area -eq 'Intune/Services').Details
if($svc -and ($svc.IME_Service -ne 'Running' -or $svc.Dmwappush -ne 'Running')){ $flags += "IME/MDM service not healthy. Restart services and trigger sync." }

# LocalUserGroup friction (post-upgrade)
$loc = ($results | ? Area -eq 'Local Accounts').Details
if($loc -and $loc.SysAdminExists -eq $true){
    # not a failure, but highlight when Intune profile is also failing
    $flags += "Managed local admin already exists; if profile errors persist (0x87d1fde8), remove and let policy recreate."
}

# Connectivity
$net = ($results | ? Area -eq 'Network Prereqs').Details
if($net){
    if($net.ApproxTimeSkewSec -gt 300 -or $net.ApproxTimeSkewSec -lt -300){ $flags += "Large time skew detected. Fix NTP." }
    foreach($k in $net.TLS443Reachable.Keys){ if(-not $net.TLS443Reachable[$k]){ $flags += "443 to $k blocked. Check proxy/SSL intercept." } }
}

$summary = [pscustomobject]@{
    Timestamp = $stamp
    Machine   = $env:COMPUTERNAME
    User      = $env:USERNAME
    Findings  = $results
    Flags     = $flags
}

$summary | ConvertTo-Json -Depth 6 | Set-Content -Path $json -Encoding UTF8
Add-Note ""
Add-Note "Flags:"
$flags | ForEach-Object { Add-Note (" - " + $_) }

if($CollectLogs){
    try{
        $toZip = Get-ChildItem $OutDir -File | Where-Object { $_.Name -like "*$stamp*" -and $_.Extension -in ('.txt','.json','.evtx','.log') }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::CreateFromDirectory($OutDir, $zip) 2>$null
    } catch { Add-Note "ZIP creation failed: $_" }
}

Add-Note ""
Add-Note "Done. Report:"
Add-Note " - TXT:  $txt"
Add-Note " - JSON: $json"
if(Test-Path $zip){ Add-Note " - ZIP:  $zip" }

# Optional one-click remediation hints (commented for safety)
# Start-ScheduledTask -TaskName "OMADMClient" -ErrorAction SilentlyContinue
# Get-Service IntuneManagementExtension,dmwappushservice | Restart-Service -Force

Write-Host "`nSummary flags:`n - " + ($flags -join "`n - ")
Write-Host "Artifacts:`n $txt`n $json`n $zip"

<# 
 Tier-0 Intune + Global Secure Access Diagnostic
 Author: Tier-0 Escalation
 Purpose: Collects detailed signals (AAD join, PRT, MDM, certs, network) to explain
          why GSA/Intune fails to connect after user logon or OS upgrade.
 Output:  JSON + TXT in  C:\ProgramData\Tier0Diag
 Tested:  PS 5.1 + 7.4
#>

$ErrorActionPreference = "Stop"
$OutDir = "C:\Windows\Temp\IntuneDiag"
New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
$stamp = (Get-Date -Format "yyyyMMdd_HHmmss")
$TxtOut = Join-Path $OutDir "Diag_$stamp.txt"
$JsonOut = Join-Path $OutDir "Diag_$stamp.json"

function Write-Log([string]$msg) {
    $msg | Tee-Object -FilePath $TxtOut -Append | Out-Null
}
function Add-Find($Area,$Status,$Message,$Details){
    [pscustomobject]@{Area=$Area;Status=$Status;Message=$Message;Details=$Details}
}

$Findings = @()
Write-Log "=== Tier-0 GSA Diagnostic Run $stamp ==="

# ────────────────────────────────────────────────
#  OS and User Context
# ────────────────────────────────────────────────
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    $u = whoami
    $isWin11 = [int]$os.BuildNumber -ge 22000
    $Findings += Add-Find "OS" "Info" "Detected $($os.Caption) ($u)" @{
        Version=$os.Version;Build=$os.BuildNumber;Domain=$cs.Domain;Win11=$isWin11
    }
    Write-Log "OS: $($os.Caption) Build $($os.BuildNumber)  User: $u"
}
catch {
    $Findings += Add-Find "OS" "Warn" "OS query failed" @{Error=$_.Exception.Message}
}

# ────────────────────────────────────────────────
#  Azure AD Join / PRT Token Check
# ────────────────────────────────────────────────
try {
    $d = (dsregcmd /status) -join "`n"
    $prt = if($d -match "PRT\s*:\s*YES"){"YES"}else{"NO"}
    $Findings += Add-Find "AAD Join" ($(if($prt -eq "YES"){"Pass"}else{"Fail"})) "PRT=$prt" @{
        AzureAdJoined=($d -match "AzureAdJoined\s*:\s*YES");
        WorkplaceJoined=($d -match "WorkplaceJoined\s*:\s*YES")
    }
    Write-Log "AAD Joined: $($d -match 'AzureAdJoined\s*:\s*YES')  PRT=$prt"
}
catch {
    $Findings += Add-Find "AAD Join" "Warn" "dsregcmd failed" @{Error=$_.Exception.Message}
}

# ────────────────────────────────────────────────
#  Intune / MDM Certificate Check
# ────────────────────────────────────────────────
try {
    $cert = Get-ChildItem Cert:\LocalMachine\My |
        Where-Object { $_.Subject -match "MS-Organization-Access" } |
        Sort-Object NotAfter -Descending | Select-Object -First 1

    if($cert){
        $days = [math]::Round(($cert.NotAfter - (Get-Date)).TotalDays)
        $status = if($days -gt 15){"Pass"}elseif($days -gt 0){"Warn"}else{"Fail"}
        $Findings += Add-Find "MDM Cert" $status "MDM cert expires in $days days" @{
            Subject=$cert.Subject;Expires=$cert.NotAfter
        }
        Write-Log "MDM cert expires in $days days"
    }
    else {
        $Findings += Add-Find "MDM Cert" "Fail" "No MS-Organization-Access cert found" @{}
    }
}
catch {
    $Findings += Add-Find "MDM Cert" "Warn" "Cert check failed" @{Error=$_.Exception.Message}
}

# ────────────────────────────────────────────────
#  Core Service Health
# ────────────────────────────────────────────────
$Services = "IntuneManagementExtension","dmwappushservice","omadmclient"
foreach($svc in $Services){
    $s = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if($s){
        $state = $s.Status
        $Findings += Add-Find "Service" ($(if($state -eq "Running"){"Pass"}else{"Warn"})) "$svc $state" @{}
        Write-Log "$svc = $state"
    } else {
        $Findings += Add-Find "Service" "Fail" "$svc missing" @{}
    }
}

# ────────────────────────────────────────────────
#  Network & TLS Connectivity
# ────────────────────────────────────────────────
$Targets = "login.microsoftonline.com","enterpriseregistration.windows.net","device.login.microsoftonline.com"
foreach($t in $Targets){
    $ok = Test-NetConnection $t -Port 443 -InformationLevel Quiet
    $Findings += Add-Find "Network" ($(if($ok){"Pass"}else{"Fail"})) "443 to $t" @{Reachable=$ok}
    Write-Log "$t reachable: $ok"
}

# ────────────────────────────────────────────────
#  Event Log Sample (AAD & MDM Providers)
# ────────────────────────────────────────────────
$Logs = "Microsoft-Windows-AAD/Operational",
        "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin"
foreach($l in $Logs){
    try{
        $err = Get-WinEvent -LogName $l -ErrorAction SilentlyContinue |
               Where-Object {$_.LevelDisplayName -eq 'Error'} |
               Select-Object -First 3
        $Findings += Add-Find "EventLog" "Info" "$l errors: $($err.Count)" @{
            Events = $err | ForEach-Object { $_.Message.Substring(0,[math]::Min(150,$_.Message.Length)) }
        }
    } catch{}
}

# ────────────────────────────────────────────────
#  Output Results
# ────────────────────────────────────────────────
$Findings | ConvertTo-Json -Depth 4 | Out-File -Encoding UTF8 $JsonOut
$Findings | Format-Table -AutoSize | Out-String | Out-File -Encoding UTF8 $TxtOut -Append

Write-Log "Diagnostics complete."
Write-Host "`nDiagnostics complete.`nReport saved to:`n$TxtOut`n$JsonOut`n"

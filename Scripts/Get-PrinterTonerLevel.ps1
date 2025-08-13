<#
.SYNOPSIS
  Toner monitor for SNMP-capable printers with HTML email via Microsoft Graph, fully using the provided template.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PrinterIP = '10.14.0.99',
    [Parameter()][string]$Community = 'public',
    [Parameter()][int]$SnmpPort = 161,
    [Parameter()][int]$ThresholdPct = 20,
    [Parameter(Mandatory)][string]$To = 'edaniels@openapproach.com',
    [Parameter(Mandatory)][string]$From = 'cloudadmin@simonpearce.com',
    [Parameter()][string]$BaseUrl = "http://10.14.0.99"
)

# -------------------- Helpers --------------------
function Test-Command { param([string]$Name) [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue) }
function HtmlEncode([string]$s) { if ([string]::IsNullOrWhiteSpace($s)) { return "" }; return [System.Web.HttpUtility]::HtmlEncode($s) }
function Get-ColorHex { param([string]$Name) switch -Regex ($Name) { '(?i)black' { '#000'; break } '(?i)cyan' { '#00bcd4'; break } '(?i)magenta' { '#e91e63'; break } '(?i)yellow' { '#ffeb3b'; break } default { '#777' } } }

# -------------------- SNMP setup --------------------
$HaveProxx = Test-Command 'Invoke-SnmpWalk'
if (-not $HaveProxx) {
    try { Install-Module Proxx.SNMP -Scope CurrentUser -Force; $HaveProxx = Test-Command 'Invoke-SnmpWalk' } 
    catch { Write-Verbose "Proxx.SNMP not available; SNMP skipped." }
}

$Oids = @{
    SuppliesTableBase     = '1.3.6.1.2.1.43.11.1.1'
    SuppliesColorIdxBase  = '1.3.6.1.2.1.43.11.1.1.3.1'
    SuppliesDescBase      = '1.3.6.1.2.1.43.11.1.1.6.1'
    SuppliesMaxBase       = '1.3.6.1.2.1.43.11.1.1.8.1'
    SuppliesLevelBase     = '1.3.6.1.2.1.43.11.1.1.9.1'
    ColorantValueBase     = '1.3.6.1.2.1.43.12.1.1.4.1'
}

function Invoke-WalkExact { param([string]$BaseOid) if (-not $HaveProxx) { return @() }; try { $res = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $BaseOid -Community $Community -UDPport $SnmpPort -Verbose:$false -Timeout 5000; $res | ForEach-Object { [pscustomobject]@{ Oid=$_.OID; Value=$_.Value } } } catch { @() } }
function Get-PrinterSuppliesFromSnmp {
    if (-not $HaveProxx) { return @() }
    $supWalk  = Invoke-WalkExact -BaseOid $Oids.SuppliesTableBase
    if (-not $supWalk -or $supWalk.Count -eq 0) { return @() }
    $colorVal = Invoke-WalkExact -BaseOid $Oids.ColorantValueBase
    $colorMap = @{}
    foreach ($row in $colorVal) { if ($row.Oid -match '\.43\.12\.1\.1\.4\.1\.(?<cidx>\d+)$') { $colorMap[$Matches['cidx']] = ($row.Value -replace '^STRING:\s*','').Trim(' "') } }
    $rows=@{}
    foreach ($row in $supWalk) { 
        if ($row.Oid -match '\.43\.11\.1\.1\.(?<leaf>\d+)\.1\.(?<idx>\d+)$') {
            $leaf=[int]$Matches['leaf']; $idx=[int]$Matches['idx']
            if (-not $rows.ContainsKey($idx)) { $rows[$idx] = [ordered]@{} }
            $val=$row.Value -replace '^(?:STRING|INTEGER|Gauge32):\s*',''; $val=$val.Trim(' "')
            switch ($leaf) {3 {$rows[$idx]['ColorIdx']=$val} 6 {$rows[$idx]['Desc']=$val} 8 {$rows[$idx]['Max']=[int]($val -as [int])} 9 {$rows[$idx]['Level']=[int]($val -as [int])} }
        }
    }
    $result=@()
    foreach ($k in ($rows.Keys|Sort-Object)) {
        $r=$rows[$k]; $name=$null
        if ($r['ColorIdx'] -and $colorMap.ContainsKey($r['ColorIdx'])) {$name=$colorMap[$r['ColorIdx']]}
        if (-not $name -and $r['Desc']) { switch -Regex ($r['Desc']) {'(?i)black' {$name='Black'} '(?i)cyan' {$name='Cyan'} '(?i)magenta' {$name='Magenta'} '(?i)yellow' {$name='Yellow'} default {$name=$r['Desc']}} }
        $pct=$null; $lvl=$r['Level']; $max=$r['Max']
        if ($lvl -ge 0 -and $max -gt 0) {$pct=[int][math]::Min(100,[math]::Round(($lvl*100.0)/$max,0))}
        $result+=[pscustomobject]@{Source='SNMP'; Color=$name; Description=$r['Desc']; MaxCapacity=$max; LevelUnits=$lvl; Percent=$pct}
    }
    $result
}

# -------------------- HTTP fallback --------------------
function Get-PrinterSuppliesFromWeb {
    param([string]$Base)
    if ([string]::IsNullOrWhiteSpace($Base)) { return @() }
    $tryUrls = @("$Base/hp/device/info_suppliesStatus.html?tab=Home&menu=SupplyStatus")
    foreach ($u in $tryUrls) {
        try {
            $resp = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $text = $resp.Content
            if ($text) { return Convert-HPConsumablesHtml -Html $text }
        } catch { continue }
    }
    @()
}
function Convert-HPConsumablesHtml { param([string]$Html)
    $results=New-Object System.Collections.Generic.List[object]
    $colors='Black','Cyan','Magenta','Yellow'
    foreach ($c in $colors) {
        $m=[regex]::Match($Html,"(?is)($c)[^%]{0,200}?(\d{1,3})\s*%")
        if ($m.Success) { $results.Add([pscustomobject]@{Source='HTTP'; Color=$m.Groups[1].Value; Description="$($m.Groups[1].Value) Cartridge (web)"; MaxCapacity=100; LevelUnits=[int]$m.Groups[2].Value; Percent=[int]$m.Groups[2].Value}) }
    }
    return ,$results
}

# -------------------- Collect toner data --------------------
$toners=@()
$snmpRows=Get-PrinterSuppliesFromSnmp
if ($snmpRows.Count -gt 0) { $toners=$snmpRows|Where-Object {$_.Percent -ne $null}}
if ($toners.Count -eq 0 -and $BaseUrl) { $toners=Get-PrinterSuppliesFromWeb -Base $BaseUrl }

if ($toners.Count -eq 0) { Write-Error "No supply data via SNMP or HTTP."; return }

# -------------------- Render HTML --------------------
$cards=''
foreach ($row in $toners) {
    $pct=[int][math]::Max(0,[math]::Min(100,($row.Percent -as [int])))
    $lbl=HtmlEncode(($row.Color)?$row.Color:$row.Description)
    $src=HtmlEncode($row.Source)
    $cssColor=($row.Color).ToLower()
    $cards+=@"
    <div class='toner-item $cssColor'>
        <div class='toner-header'>
            <div class='toner-name'>$lbl</div>
            <div class='toner-percentage'>$pct%</div>
        </div>
        <div class='progress-container'>
            <div class='progress-bar $cssColor' style='width:$pct%;'></div>
        </div>
        <div class='cartridge-info'>$lbl Cartridge | Source: $src</div>
    </div>
"@
}

$anyLow=$toners|Where-Object{$_.Percent -le $ThresholdPct}
$alertText=if($anyLow){"Threshold met: ≤ $ThresholdPct%"}else{"All above threshold ($ThresholdPct%)"}

$body=@"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width, initial-scale=1.0'>
<title>Toner Status Dashboard</title>
<style>
body {font-family:'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width:800px; margin:0 auto; padding:20px; background-color:#f5f5f5;}
.container {background:white; border-radius:8px; padding:30px; box-shadow:0 2px 10px rgba(0,0,0,0.1);}
.header {text-align:center; margin-bottom:30px; padding-bottom:20px; border-bottom:2px solid #eee;}
.printer-ip {font-size:24px; font-weight:bold; color:#333; margin-bottom:5px;}
.status-good {color:#28a745; font-weight:600; font-size:16px;}
.toner-item {margin-bottom:25px; padding:20px; border-radius:6px; background-color:#fafafa; border-left:4px solid;}
.toner-item.black {border-left-color:#000;}
.toner-item.cyan {border-left-color:#00bcd4;}
.toner-item.magenta {border-left-color:#e91e63;}
.toner-item.yellow {border-left-color:#ffeb3b;}
.toner-header {display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;}
.toner-name {font-size:18px; font-weight:600; color:#333;}
.toner-percentage {font-size:20px; font-weight:bold; color:#333;}
.progress-container {width:100%; height:24px; background-color:#e0e0e0; border-radius:12px; overflow:hidden; margin-bottom:8px;}
.progress-bar {height:100%; border-radius:12px; transition:width 0.3s ease; position:relative;}
.progress-bar.black {background:linear-gradient(90deg,#333,#000);}
.progress-bar.cyan {background:linear-gradient(90deg,#4dd0e1,#00bcd4);}
.progress-bar.magenta {background:linear-gradient(90deg,#f06292,#e91e63);}
.progress-bar.yellow {background:linear-gradient(90deg,#fff176,#ffeb3b);}
.cartridge-info {font-size:14px; color:#666; font-style:italic;}
.footer {margin-top:30px; padding-top:20px; border-top:1px solid #eee; text-align:center; color:#888; font-size:12px;}
.threshold-indicator {display:inline-block; padding:4px 8px; border-radius:12px; font-size:12px; font-weight:600; margin-left:10px;}
.above-threshold {background-color:#d4edda; color:#155724;}
</style>
</head>
<body>
<div class='container'>
    <div class='header'>
        <div class='printer-ip'>Toner Status — $(HtmlEncode $PrinterIP)</div>
        <div class='status-good'>$alertText <span class='threshold-indicator above-threshold'>✓ GOOD</span></div>
    </div>
    $cards
    <div class='footer'>
        Retrieved on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') via $(if ($snmpRows.Count -gt 0){'SNMP (Printer-MIB)'} else {'HTTP scrape'}).
    </div>
</div>
</body>
</html>
"@

# -------------------- Send via Microsoft Graph --------------------
function Use-Graph {
    if (-not (Get-Module -ListAvailable Microsoft.Graph)) { Install-Module Microsoft.Graph -Scope CurrentUser -Force }
    Import-Module Microsoft.Graph -ErrorAction Stop
    $ctx=Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) { Connect-MgGraph -Scopes "Mail.Send" -ErrorAction Stop | Out-Null }
    elseif ($ctx.Scopes -notcontains "Mail.Send") { Disconnect-MgGraph; Connect-MgGraph -Scopes "Mail.Send" -ErrorAction Stop | Out-Null }
}

try {
    Use-Graph
    $msg=@{Subject="Toner status [HP Color LaserJet M477fnw] [$PrinterIP]"; Body=@{ContentType="HTML"; Content=$body}; ToRecipients=@(@{EmailAddress=@{Address=$To}})}
    Send-MgUserMail -UserId $From -Message $msg -SaveToSentItems:$false
    Write-Host "Email sent successfully."
} catch { Write-Warning "Graph send failed: $($_.Exception.Message)"; Write-Output $body }

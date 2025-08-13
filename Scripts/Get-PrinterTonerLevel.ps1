<#
.SYNOPSIS
  Toner monitor for SNMP-capable printers with HTML email via Microsoft Graph.

.DESCRIPTION
  Fetches toner levels via SNMP or HTTP and renders a styled HTML dashboard.
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

# -------------------- Utilities --------------------
function Test-Command { param([string]$Name) [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue) }

function Get-ColorHex {
    param([string]$Name)
    switch -Regex ($Name) {
        '(?i)black' { '#000000'; break }
        '(?i)cyan' { '#00bcd4'; break }
        '(?i)magenta' { '#e91e63'; break }
        '(?i)yellow' { '#ffeb3b'; break }
        default { '#777777' }
    }
}

# -------------------- SNMP toner levels --------------------
$snmpOid = "1.3.6.1.2.1.43.11.1.1.9"
$snmpResult = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $snmpOid -Community $Community -UDPport $SnmpPort -Verbose

if ($snmpResult.Count -eq 0) {
    Write-Warning "No SNMP data found for $PrinterIP."
} else {
    # Map returned rows to colors
    $colors = @('Black','Cyan','Magenta','Yellow')
    $toners = @()
    for ($i = 0; $i -lt $snmpResult.Count; $i++) {
        $toners += [pscustomobject]@{
            Source      = 'SNMP'
            Color       = $colors[$i]
            Description = "$($colors[$i]) Cartridge"
            Percent     = [int]$snmpResult[$i].Data
            LevelUnits  = [int]$snmpResult[$i].Data
        }
    }
}


# -------------------- HTTP fallback --------------------
function Get-PrinterSuppliesFromWeb {
    param([string]$Base)
    if ([string]::IsNullOrWhiteSpace($Base)) { return @() }

    $urls = @(
        "$Base/hp/device/info_suppliesStatus.html?tab=Home&menu=SupplyStatus",
        "$Base/DevMgmt/ConsumableConfigDyn.xml",
        "$Base/DevMgmt/ProductUsageDyn.xml"
    )

    foreach ($u in $urls) {
        try {
            $resp = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
            $text = $resp.Content
            if ($text) {
                # Parse using your previous Convert-HPConsumablesHtml logic
                $toners = Convert-HPConsumablesHtml -Html $text
                if ($toners.Count -gt 0) { return $toners }
            }
        } catch { continue }
    }
    @()
}

function Convert-HPConsumablesHtml { param([string]$Html)
    $results = New-Object System.Collections.Generic.List[object]
    $colors = 'Black','Cyan','Magenta','Yellow'
    foreach ($c in $colors) {
        $m = [regex]::Match($Html, "(?is)($c)[^%]{0,200}?(\d{1,3})\s*%")
        if ($m.Success) {
            $results.Add([pscustomobject]@{
                Source='HTTP'; Color=$m.Groups[1].Value; Description="$($m.Groups[1].Value) Cartridge"; LevelUnits=[int]$m.Groups[2].Value; Percent=[int]$m.Groups[2].Value
            })
        }
    }
    ,$results
}

# -------------------- Collect data --------------------
$toners = @()
$snmpRows = Get-PrinterSuppliesFromSnmp
if ($snmpRows.Count -gt 0) { $toners = $snmpRows }

if ($toners.Count -eq 0) { $toners = Get-PrinterSuppliesFromWeb -Base $BaseUrl }

if ($toners.Count -eq 0) {
    Write-Warning "No supply data via SNMP or HTTP. Skipping email."
    return
}

# -------------------- Render HTML --------------------
$cards = ''
foreach ($row in $toners) {
    $colorClass = ($row.Color).ToLower()
    $pct = [int][math]::Max(0,[math]::Min(100,$row.Percent))
    $summary = "$($row.Description) | Source: $($row.Source)"
    $cards += @"
    <div class='toner-item $colorClass'>
        <div class='toner-header'>
            <div class='toner-name'>$($row.Color)</div>
            <div class='toner-percentage'>$pct%</div>
        </div>
        <div class='progress-container'>
            <div class='progress-bar $colorClass' style='width:$pct%;'></div>
        </div>
        <div class='cartridge-info'>$summary</div>
    </div>
"@
}

$anyLow = $toners | Where-Object { $_.Percent -le $ThresholdPct }
$statusText = if ($anyLow) { "Threshold met: ≤ $ThresholdPct%" } else { "All above threshold ($ThresholdPct%)" }

$body = @"
<!DOCTYPE html>
<html lang='en'>
<head><meta charset='UTF-8'><style>
/* include your template CSS here */
</style></head>
<body>
<div class='container'>
    <div class='header'>
        <div class='printer-ip'>Toner Status — $PrinterIP</div>
        <div class='status-good'>$statusText</div>
    </div>
    $cards
    <div class='footer'>
        Retrieved on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') via $(if ($snmpRows.Count -gt 0){'SNMP (Printer-MIB)'} else {'HTTP scrape'}).
    </div>
</div>
</body>
</html>
"@

# -------------------- Email via Graph --------------------
function Use-Graph {
    if (-not (Get-Module -ListAvailable Microsoft.Graph)) { Install-Module Microsoft.Graph -Scope CurrentUser -Force }
    Import-Module Microsoft.Graph -ErrorAction Stop
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx -or $ctx.Scopes -notcontains "Mail.Send") {
        Connect-MgGraph -Scopes "Mail.Send" -ErrorAction Stop | Out-Null
    }
}

try {
    Use-Graph
    $msg = @{
        Subject = if ($anyLow) { "TONER LOW [$PrinterIP]" } else { "Toner status [$PrinterIP]" }
        Body    = @{ ContentType="HTML"; Content=$body }
        ToRecipients = @(@{ EmailAddress = @{ Address=$To } })
    }
    Send-MgUserMail -UserId $From -Message $msg -SaveToSentItems:$false
    Write-Host "Email sent successfully."
} catch {
    Write-Warning "Graph send failed: $($_.Exception.Message)"
    Write-Output $body
}

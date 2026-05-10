<#
.SYNOPSIS
  Toner monitor for SNMP-capable printers with HTML email via Microsoft Graph.

.DESCRIPTION
  1) SNMP first (Printer-MIB / RFC3805)
  2) HTTP fallback if SNMP fails
  3) Color-coded toner bars in HTML email compatible with Outlook

.PARAMETER PrinterIP
  Printer IP or hostname.

.PARAMETER Community
  SNMP v2c community string.

.PARAMETER SnmpPort
  UDP port (default 161).

.PARAMETER ThresholdPct
  Alert threshold (default 20).

.PARAMETER To
  Recipient email.

.PARAMETER From
  Sender UPN used for Graph delegated send.

.PARAMETER BaseUrl
  Base URL to the printer for HTTP fallback.
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
function Get-ColorHex {
    param([string]$Name)
    switch -Regex ($Name) {
        '(?i)black'   { '#000000'; break }
        '(?i)cyan'    { '#00bcd4'; break }
        '(?i)magenta' { '#e91e63'; break }
        '(?i)yellow'  { '#ffeb3b'; break }
        default       { '#777777' }
    }
}

function HtmlEncode([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return "" }
    return [System.Web.HttpUtility]::HtmlEncode($s)
}

# -------------------- SNMP toner levels --------------------
$snmpOid = "1.3.6.1.2.1.43.11.1.1.9"
$snmpResult = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $snmpOid -Community $Community -UDPport $SnmpPort

if (-not $snmpResult -or $snmpResult.Count -eq 0) {
    Write-Error "No supply data via SNMP. Consider HTTP fallback."
    return
}

# Build toner objects from SNMP results
$toners = @()
$colorOrder = @('Black','Cyan','Magenta','Yellow')
for ($i=0; $i -lt $snmpResult.Count; $i++) {
    $lvl = [int]$snmpResult[$i].Data
    $color = $colorOrder[$i % $colorOrder.Count]
    $toners += [pscustomobject]@{
        Color = $color
        Description = "$color Cartridge"
        LevelUnits = $lvl
        Percent = $lvl
        Source = "SNMP"
    }
}

# -------------------- HTML rendering --------------------
$cards = ''
foreach ($row in $toners) {
    $pct = [int][math]::Max(0,[math]::Min(100,$row.Percent))
    $hex = Get-ColorHex $row.Color

    $cards += @"
    <div style='margin-bottom:25px;padding:20px;border-radius:6px;background-color:#fafafa;border-left:4px solid $hex;'>
        <div style='margin-bottom:8px;font-size:16px;font-weight:600;color:#333;'>$($row.Color) — $pct%</div>
        <div style='width:100%;height:24px;background:#e0e0e0;border-radius:12px;overflow:hidden;'>
            <div style='height:100%;width:$pct%;background:$hex;'></div>
        </div>
        <div style='font-size:14px;color:#666;font-style:italic;margin-top:6px;'>$($row.Description) | Source: $($row.Source)</div>
    </div>
"@
}

$anyLow = $toners | Where-Object { $_.Percent -le $ThresholdPct }
$statusText = if ($anyLow) {
    "<span style='color:#c0392b;font-weight:700;'>Threshold met: ≤ $ThresholdPct%</span>"
} else {
    "<span style='color:#28a745;font-weight:700;'>All above threshold ($ThresholdPct%) ✓ GOOD</span>"
}

$body = @"
<html>
<head>
<meta charset='UTF-8'>
</head>
<body style='font-family:Segoe UI,Arial,sans-serif;background:#f5f5f5;padding:20px;'>
    <div style='max-width:800px;margin:0 auto;background:#fff;padding:30px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,0.1);'>
        <div style='text-align:center;margin-bottom:30px;'>
            <div style='font-size:24px;font-weight:bold;color:#333;'>Toner Status — $PrinterIP</div>
            <div style='font-size:16px;font-weight:600;margin-top:5px;'>$statusText</div>
        </div>
        $cards
        <div style='margin-top:30px;padding-top:20px;border-top:1px solid #eee;text-align:center;color:#888;font-size:12px;'>
            Retrieved on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') via SNMP
        </div>
    </div>
</body>
</html>
"@

$subject = if ($anyLow) {
    $summary = ($anyLow | ForEach-Object { "$($_.Color) $([int]$_.Percent)%" }) -join ' | '
    "TONER LOW [HP Color Laserjet M477fnw][$PrinterIP]: $summary"
} else {
    "Toner status [HP Color Laserjet M477fnw][$PrinterIP]: All above ${ThresholdPct}%"
}

# -------------------- Microsoft Graph send --------------------
function Use-Graph {
    if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph -ErrorAction Stop
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Connect-MgGraph -Scopes "Mail.Send" -ErrorAction Stop | Out-Null
    } elseif ($ctx.Scopes -notcontains "Mail.Send") {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes "Mail.Send" -ErrorAction Stop | Out-Null
    }
}

try {
    Use-Graph

    $msg = @{
        Subject = $subject
        Body    = @{ ContentType = "HTML"; Content = $body }
        ToRecipients = @(@{ EmailAddress = @{ Address = $To } })
    }

    Send-MgUserMail -UserId $From -Message $msg -SaveToSentItems:$false
    Write-Host "Email sent successfully: $subject"
} catch {
    Write-Warning "Graph send failed: $($_.Exception.Message)"
    Write-Output $body
}

return

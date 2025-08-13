<#
.SYNOPSIS
  Toner monitor for SNMP-capable printers with HTML email via Microsoft Graph.

.DESCRIPTION
  Uses SNMP first (working OID) and renders an HTML email with inline CSS.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PrinterIP = '10.14.0.99',
    [Parameter()][string]$Community = 'public',
    [Parameter()][int]$SnmpPort = 161,
    [Parameter()][int]$ThresholdPct = 20,
    [Parameter(Mandatory)][string]$To = 'edaniels@openapproach.com',
    [Parameter(Mandatory)][string]$From = 'cloudadmin@simonpearce.com'
)

# -------------------- SNMP toner levels --------------------
$snmpOid = "1.3.6.1.2.1.43.11.1.1.9"
$snmpResult = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $snmpOid -Community $Community -UDPport $SnmpPort

if (-not $snmpResult -or $snmpResult.Count -eq 0) {
    Write-Error "No SNMP data retrieved from $PrinterIP"
    return
}

Write-Output $snmpResult
Write-Host ""

# -------------------- Map SNMP results to toner objects --------------------
$colors = @('Black','Cyan','Magenta','Yellow')
$toners = @()
for ($i = 0; $i -lt $snmpResult.Count; $i++) {
    $percent = [int]$snmpResult[$i].Data
    if ($percent -gt 100) { $percent = 100 }
    $toners += [pscustomobject]@{
        Source      = 'SNMP'
        Color       = $colors[$i]
        Description = "$($colors[$i]) Cartridge"
        Percent     = $percent
        LevelUnits  = $percent
    }
}

# -------------------- Inline CSS helpers --------------------
function Get-ColorHex($name) {
    switch ($name.ToLower()) {
        'black'   { '#000000' }
        'cyan'    { '#00bcd4' }
        'magenta' { '#e91e63' }
        'yellow'  { '#ffeb3b' }
        default   { '#777777' }
    }
}

# -------------------- Build HTML --------------------
$cards = ''
foreach ($row in $toners) {
    $pct = [int][math]::Max(0,[math]::Min(100,$row.Percent))
    $hex = Get-ColorHex $row.Color
    $colorClass = $row.Color.ToLower()
    $cards += @"
    <div style='margin-bottom:25px;padding:20px;border-radius:6px;background-color:#fafafa;border-left:4px solid $hex;'>
        <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;'>
            <div style='font-size:18px;font-weight:600;color:#333;'>$($row.Color)</div>
            <div style='font-size:20px;font-weight:bold;color:#333;'>$pct%</div>
        </div>
        <div style='width:100%;height:24px;background:#e0e0e0;border-radius:12px;overflow:hidden;margin-bottom:8px;'>
            <div style='height:100%;width:$pct%;background:$hex;'></div>
        </div>
        <div style='font-size:14px;color:#666;font-style:italic;'>$($row.Description) | Source: $($row.Source)</div>
    </div>
"@
}

# -------------------- Threshold indicator --------------------
$anyLow = $toners | Where-Object { $_.Percent -le $ThresholdPct }
$thresholdText = if ($anyLow) { "Threshold met: ≤ $ThresholdPct% ❌" } else { "All above threshold ($ThresholdPct%) ✓ GOOD" }
$statusColor = if ($anyLow) { '#c0392b' } else { '#28a745' }

# -------------------- HTML body --------------------
$body = @"
<html>
  <body style='margin:0;padding:0;background:#f5f5f5;font-family:Segoe UI,Arial,sans-serif;'>
    <div style='max-width:800px;margin:0 auto;padding:20px;'>
        <div style='background:white;border-radius:8px;padding:30px;box-shadow:0 2px 10px rgba(0,0,0,0.1);'>
            <div style='text-align:center;margin-bottom:30px;padding-bottom:20px;border-bottom:2px solid #eee;'>
                <div style='font-size:24px;font-weight:bold;color:#333;margin-bottom:5px;'>Toner Status — $PrinterIP</div>
                <div style='color:$statusColor;font-weight:600;font-size:16px;'>$thresholdText</div>
            </div>
            $cards
            <div style='margin-top:30px;padding-top:20px;border-top:1px solid #eee;text-align:center;color:#888;font-size:12px;'>
                Retrieved on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') via SNMP
            </div>
        </div>
    </div>
  </body>
</html>
"@

# -------------------- Email via Microsoft Graph --------------------
function Use-Graph {
    if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    }
    Import-Module Microsoft.Graph -ErrorAction Stop
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $ctx) { Connect-MgGraph -Scopes "Mail.Send" | Out-Null }
    elseif ($ctx.Scopes -notcontains "Mail.Send") {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        Connect-MgGraph -Scopes "Mail.Send" | Out-Null
    }
}

try {
    Use-Graph
    $subject = if ($anyLow) {
        $summary = ($anyLow | ForEach-Object { "$($_.Color) $($_.Percent)%" }) -join ' | '
        "TONER LOW [$PrinterIP]: $summary"
    } else {
        "Toner status [$PrinterIP]: All above $ThresholdPct%"
    }

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

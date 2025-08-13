<#
.SYNOPSIS
  Toner monitor for SNMP-capable printers with HTML email via Microsoft Graph.

.DESCRIPTION
  1) Attempts SNMP first (Printer-MIB / RFC3805).
     - Uses Proxx.SNMP Invoke-SnmpWalk with your *known working* parameters.
  2) If SNMP yields no usable percent values, falls back to HTTP scrape of the
     HP status page (parses your provided HTML structure, including % text and
     the hpGasGaugeBorder width bars).
  3) Renders color-coded bars (inline CSS, Outlook-safe) and sends via Graph.

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
  Sender UPN used for Graph delegated send (browser auth popup).

.PARAMETER BaseUrl
  Base URL to the printer (e.g. http://10.14.0.99). We will try common supply pages:
  - /hp/device/info_suppliesStatus.html?tab=Home&menu=SupplyStatus
  - /DevMgmt/ConsumableConfigDyn.xml
  - /DevMgmt/ProductUsageDyn.xml

.NOTES
  Requires: Microsoft.Graph
  Optional: Proxx.SNMP (for SNMP; otherwise only HTTP is used)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory)][string]$PrinterIP,
  [Parameter()][string]$Community = 'public',
  [Parameter()][int]$SnmpPort = 161,
  [Parameter()][int]$ThresholdPct = 20,
  [Parameter(Mandatory)][string]$To,
  [Parameter(Mandatory)][string]$From,
  [Parameter()][string]$BaseUrl = ""
)

# -------------------- Utility --------------------
function Test-Command { param([string]$Name) [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue) }

function HtmlEncode([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  return [System.Web.HttpUtility]::HtmlEncode($s)
}

function Get-ColorHex {
  param([string]$Name)
  switch -Regex ($Name) {
    '(?i)black|k(?!ey)'   { '#111111' ; break }
    '(?i)cyan'            { '#0095ff' ; break }
    '(?i)magenta'         { '#ff00aa' ; break }
    '(?i)yellow'          { '#f2c200' ; break }
    default               { '#777777' }
  }
}

# -------------------- Optional: Proxx.SNMP availability --------------------
$HaveProxx = Test-Command 'Invoke-SnmpWalk'
if (-not $HaveProxx) {
  try {
    Install-Module Proxx.SNMP -Scope CurrentUser -Force -ErrorAction Stop
    $HaveProxx = Test-Command 'Invoke-SnmpWalk'
  } catch {
    Write-Verbose "Proxx.SNMP not available; SNMP stage will be skipped. HTTP fallback will be used."
  }
}

# -------------------- OIDs (Printer-MIB RFC3805) --------------------
$Oids = @{
  SuppliesTableBase     = '1.3.6.1.2.1.43.11.1.1'   # table root
  SuppliesColorIdxBase  = '1.3.6.1.2.1.43.11.1.1.3.1'
  SuppliesDescBase      = '1.3.6.1.2.1.43.11.1.1.6.1'
  SuppliesMaxBase       = '1.3.6.1.2.1.43.11.1.1.8.1'
  SuppliesLevelBase     = '1.3.6.1.2.1.43.11.1.1.9.1'
  ColorantValueBase     = '1.3.6.1.2.1.43.12.1.1.4.1'
}

# -------------------- SNMP helpers --------------------
function Invoke-WalkExact {
  param([Parameter(Mandatory)][string]$BaseOid)

  if (-not $HaveProxx) { return @() }

  # Match your working CLI:  Invoke-SnmpWalk -IP 10.14.0.99 -Version V2 -OIDStart "1.3.6.1..." -Community public -UDPport 161
  try {
    $res = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $BaseOid -Community $Community -UDPport $SnmpPort -Verbose:$false -Timeout 5000
    # Normalize: Oid, Value
    $res | ForEach-Object {
      [pscustomobject]@{ Oid = $_.OID; Value = $_.Value }
    }
  } catch {
    Write-Warning "SNMP Walk issue ($BaseOid): $($_.Exception.Message)"
    @()
  }
}

function Get-PrinterSuppliesFromSnmp {
  if (-not $HaveProxx) { return @() }

  $supWalk  = Invoke-WalkExact -BaseOid $Oids.SuppliesTableBase
  if (-not $supWalk -or $supWalk.Count -eq 0) { return @() }

  $colorVal = Invoke-WalkExact -BaseOid $Oids.ColorantValueBase
  $colorMap = @{}
  foreach ($row in $colorVal) {
    if ($row.Oid -match '\.43\.12\.1\.1\.4\.1\.(?<cidx>\d+)$') {
      $colorMap[$Matches['cidx']] = ($row.Value -replace '^STRING:\s*','').Trim(' "')
    }
  }

  $rows = @{}
  foreach ($row in $supWalk) {
    if ($row.Oid -match '\.43\.11\.1\.1\.(?<leaf>\d+)\.1\.(?<idx>\d+)$') {
      $leaf = [int]$Matches['leaf']; $idx = [int]$Matches['idx']
      if (-not $rows.ContainsKey($idx)) { $rows[$idx] = [ordered]@{} }
      $val = $row.Value -replace '^(?:STRING|INTEGER|Gauge32):\s*',''
      $val = $val.Trim(' "')
      switch ($leaf) {
        3 { $rows[$idx]['ColorIdx']  = $val }
        6 { $rows[$idx]['Desc']      = $val }
        8 { $rows[$idx]['Max']       = [int]($val -as [int]) }
        9 { $rows[$idx]['Level']     = [int]($val -as [int]) }
      }
    }
  }

  $result = @()
  foreach ($k in ($rows.Keys | Sort-Object)) {
    $r = $rows[$k]
    $name = $null
    if ($r['ColorIdx'] -and $colorMap.ContainsKey($r['ColorIdx'])) {
      $name = $colorMap[$r['ColorIdx']]
    }
    if (-not $name -and $r['Desc']) {
      if ($r['Desc'] -match '(?i)black')   { $name = 'Black' }
      elseif ($r['Desc'] -match '(?i)cyan'){ $name = 'Cyan' }
      elseif ($r['Desc'] -match '(?i)magenta'){ $name = 'Magenta' }
      elseif ($r['Desc'] -match '(?i)yellow'){ $name = 'Yellow' }
      else { $name = $r['Desc'] }
    }

    $pct = $null
    $lvl = $r['Level']
    $max = $r['Max']
    if ($lvl -ge 0 -and $max -gt 0) {
      $pct = [int][math]::Min(100,[math]::Round(($lvl*100.0)/$max,0))
    }

    $result += [pscustomobject]@{
      Source      = 'SNMP'
      Color       = $name
      Description = $r['Desc']
      MaxCapacity = $max
      LevelUnits  = $lvl
      Percent     = $pct
    }
  }
  $result
}

# -------------------- HTTP fallback (parses your HTML) --------------------
function Get-PrinterSuppliesFromWeb {
  param([string]$Base)

  if ([string]::IsNullOrWhiteSpace($Base)) { return @() }

  $tryUrls = @(
    "$Base/hp/device/info_suppliesStatus.html?tab=Home&menu=SupplyStatus",
    "$Base/DevMgmt/ConsumableConfigDyn.xml",
    "$Base/DevMgmt/ProductUsageDyn.xml"
  )

  foreach ($u in $tryUrls) {
    try {
      $resp = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
      $text = $resp.Content
      if (-not [string]::IsNullOrWhiteSpace($text)) {
        $parsed = Convert-HPConsumablesHtml -Html $text
        if ($parsed.Count -gt 0) { return $parsed }
      }
    } catch { continue }
  }
  @()
}

function Convert-HPConsumablesHtml {
  param([Parameter(Mandatory)][string]$Html)

  $results = New-Object System.Collections.Generic.List[object]

  # Split on the cartridge "mainContentArea" sections
  $blocks = [regex]::Split($Html, '(?is)<table[^>]*class\s*=\s*"[^"]*\bmainContentArea\b[^"]*"[^>]*>')
  if ($blocks.Count -le 1) {
    # Not the expected page, try to sniff simple % by color words as last resort
    $colors = 'Black','Cyan','Magenta','Yellow'
    foreach ($c in $colors) {
      $m = [regex]::Match($Html, "(?is)($c)[^%]{0,200}?(\d{1,3})\s*%")
      if ($m.Success) {
        $results.Add([pscustomobject]@{
          Source='HTTP'; Color=$m.Groups[1].Value; Description="$($m.Groups[1].Value) Cartridge (web)";
          MaxCapacity=100; LevelUnits=[int]$m.Groups[2].Value; Percent=[int]$m.Groups[2].Value
        })
      }
    }
    return ,$results
  }

  # Each block contains: names, percentage in alignRight, and optional gauge widths
  for ($i=1; $i -lt $blocks.Count; $i++) {
    $b = $blocks[$i]

    # Cartridge Name
    $name = $null
    $nameMatch = [regex]::Match($b, '(?is)<td[^>]*class="[^"]*\bSupplyName\b[^"]*"[^>]*>(.*?)</td>')
    if ($nameMatch.Success) {
      $nameText = ($nameMatch.Groups[1].Value -replace '<br\s*/?>', "`n" -replace '<.*?>','' ).Trim()
      $name = ($nameText -split "`n")[0].Trim()
    }

    # Percentage text (e.g., "90%")
    $pct = $null
    $pctMatch = [regex]::Match($b, '(?is)<td[^>]*class="[^"]*\balignRight\b[^"]*"[^>]*>.*?(\d{1,3})\s*%.*?</td>')
    if ($pctMatch.Success) {
      $pct = [int]$pctMatch.Groups[1].Value
      if ($pct -gt 100) { $pct = 100 }
    } else {
      # Fallback: parse hpGasGaugeBorder widths -> WIDTH:90% black bar
      $w1 = [regex]::Match($b, '(?is)hpGasGaugeBorder.*?WIDTH\s*:\s*(\d{1,3})\s*%[^%]*?BACKGROUND-COLOR', 'IgnoreCase')
      if ($w1.Success) {
        $pct = [int][math]::Min(100,[int]$w1.Groups[1].Value)
      }
    }

    if ($null -eq $pct) { continue }

    # Attempt color from name
    $color = $name
    if ($name -match '(?i)black')   { $color = 'Black' }
    elseif ($name -match '(?i)cyan'){ $color = 'Cyan' }
    elseif ($name -match '(?i)magenta'){ $color = 'Magenta' }
    elseif ($name -match '(?i)yellow'){ $color = 'Yellow' }

    $results.Add([pscustomobject]@{
      Source='HTTP'; Color=$color; Description=$name; MaxCapacity=100; LevelUnits=$pct; Percent=$pct
    })
  }

  $results
}

# -------------------- Collect data --------------------
$toners = @()

# 1) SNMP stage
$snmpRows = Get-PrinterSuppliesFromSnmp
if ($snmpRows.Count -gt 0) {
  # Keep toner-like only if possible; else keep all
  $toners = $snmpRows | Where-Object {
    $_.Percent -ne $null -and ($_.Description -match '(?i)toner|black|cyan|magenta|yellow' -or $_.Color -match '(?i)black|cyan|magenta|yellow')
  }
}

# 2) HTTP fallback if no usable percents
if ($toners.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($BaseUrl)) {
  $webRows = Get-PrinterSuppliesFromWeb -Base $BaseUrl
  if ($webRows.Count -gt 0) { $toners = $webRows }
}

if ($toners.Count -eq 0) {
  Write-Error "No supply data retrieved via SNMP or HTTP. Verify SNMP v2 accessibility and/or BaseUrl."
  return
}

# -------------------- Render HTML --------------------
$cards = ''
foreach ($row in $toners) {
  $pct = [int][math]::Max(0,[math]::Min(100, ($row.Percent -as [int])))
  $hex = Get-ColorHex $row.Color
  $lbl = HtmlEncode(($row.Color) ? $row.Color : $row.Description)
  $sub = HtmlEncode(($row.Description) ? $row.Description : '')
  $cards += @"
  <div style="flex:1 1 220px; max-width:260px; min-width:220px; margin:8px; padding:12px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.08); border:1px solid #eee; background:#fff;">
    <div style="font-weight:700; font-size:14px; color:#333; margin-bottom:4px;">$lbl</div>
    <div style="font-size:12px; color:#666; margin-bottom:10px;">$sub</div>
    <div style="height:18px; width:220px; background:#f1f3f5; border:1px solid #e5e7eb; border-radius:9px; overflow:hidden;">
      <div style="height:100%; width:${pct}%; background:$hex;"></div>
    </div>
    <div style="margin-top:6px; font-size:13px; color:#333;">$pct%</div>
    <div style="margin-top:2px; font-size:11px; color:#777;">Source: $($row.Source)</div>
  </div>
"@
}

$anyLow = $toners | Where-Object { $_.Percent -le $ThresholdPct }
$alertHeader = ($anyLow) ? "<div style='color:#c0392b;font-weight:700;'>Threshold met: ≤ $ThresholdPct%</div>" : "<div style='color:#2e7d32;font-weight:700;'>All above threshold ($ThresholdPct%)</div>"

$body = @"
<html>
  <body style="margin:0;padding:0;background:#f8fafc;font-family:Segoe UI,Arial,Helvetica,sans-serif;">
    <div style="max-width:960px;margin:16px auto;padding:16px;">
      <div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;padding:16px 16px 8px 16px;box-shadow:0 2px 10px rgba(0,0,0,.05);">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
          <div style="font-size:18px;font-weight:700;color:#111;">Toner Status — $(HtmlEncode($PrinterIP))</div>
          $alertHeader
        </div>
        <div style="display:flex;flex-wrap:wrap;align-items:stretch;">$cards</div>
        <div style="margin-top:12px;font-size:12px;color:#666;">
          Retrieved on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') via $(if ($snmpRows.Count -gt 0) {'SNMP (Printer-MIB)'} else {'HTTP scrape'}).
        </div>
      </div>
    </div>
  </body>
</html>
"@

# Email subject
$subject = if ($anyLow) {
  $summary = ($anyLow | ForEach-Object { "$($_.Color ?? $_.Description) $([int]$_.Percent)%" }) -join ' | '
  "TONER LOW [$PrinterIP]: $summary"
} else {
  "Toner status [$PrinterIP]: All above ${ThresholdPct}%"
}

# -------------------- Microsoft Graph send --------------------
function Use-Graph {
  if (-not (Get-Module -ListAvailable Microsoft.Graph)) {
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
  }
  Import-Module Microsoft.Graph -ErrorAction Stop
  # Only Mail.Send needed
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
    ToRecipients = @(
      @{ EmailAddress = @{ Address = $To } }
    )
  }

  # Send as the signed-in user ($From should match the UPN that signs in)
  # If you need to send explicitly as a different user (delegated), use -UserId $From
  Send-MgUserMail -UserId $From -Message $msg -SaveToSentItems:$false

  Write-Host "Email sent: $subject"
} catch {
  Write-Warning "Graph send failed: $($_.Exception.Message)"
  # Fallback: write HTML to console so you can copy if needed
  Write-Output $body
  return
}

return

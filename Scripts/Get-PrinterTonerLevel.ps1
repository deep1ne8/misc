<#
.SYNOPSIS
  Interactive, vendor-agnostic toner monitoring via SNMP with HTML alert.

.DESCRIPTION
  - Prompts for printer, SNMP, threshold, and mail settings.
  - Uses Printer-MIB (RFC 3805) OIDs by default.
  - Auto-detects SNMP backend: Proxx.SNMP or Net-SNMP tools.
  - Computes percent = Level / MaxCapacity (handles unknown/negative).
  - Resolves colors via Colorant table; falls back to description.
  - Sends HTML “dashboard” email when any supply is <= threshold.
  - Optional HTTP EWS fallback if SNMP fails.

.NOTES
  Works with HP Color LaserJet M477fnw and most SNMP-capable printers.
  For SNMP:
    Install-Module Proxx.SNMP -Scope CurrentUser   # if you prefer PS module
  OR install Net-SNMP utilities (snmpwalk/snmpget) and ensure they’re in PATH.
#>

# -------- Utility: prompt with defaults --------
function Read-HostDefault {
  param([Parameter(Mandatory)][string]$Prompt, [string]$Default = "")
  $sfx = $Default ? " [$Default]" : ""
  $val = Read-Host "$Prompt$sfx"
  if ([string]::IsNullOrWhiteSpace($val)) { return $Default } else { return $val }
}

# -------- Prompts --------
$PrinterIP      = Read-HostDefault 'Printer IP or hostname'
$SnmpPort       = [int](Read-HostDefault 'SNMP UDP port' '161')
$Community      = Read-HostDefault 'SNMP community (v2c)' 'public'
$ThresholdPct   = [int](Read-HostDefault 'Alert threshold percentage' '20')

$Userid         = Read-HostDefault 'Email username' 'cloudadmin@simonpearce.com'
#$From           = Read-HostDefault 'Email From address'
$To             = Read-HostDefault 'Email To address'
#$SmtpServer     = Read-HostDefault 'SMTP server address'
#$SmtpClient     = Read-HostDefault 'SMTP Client' 'STARTTLS'
#$SmtpPort       = [int](Read-HostDefault 'SMTP port' '25')
#$UseSsl         = (Read-HostDefault 'Use SSL for SMTP? (Y/N)' 'N').ToUpper() -eq 'Y'
$UseAuth        = (Read-HostDefault 'Use SMTP auth? (Y/N)' 'N').ToUpper() -eq 'Y'
$FallbackUrl    = Read-HostDefault 'Fallback status page base URL (e.g., http://printer) [optional]' 'http://10.14.0.99/hp/device/info_suppliesStatus.html?tab=Home&amp;menu=SupplyStatus'

if ($UseAuth) {
  $SmtpCred = Connect-MSGraph
}

# -------- Printer-MIB OIDs (RFC 3805) --------
$Oids = @{
  # Supplies table base
  SuppliesTableBase     = '1.3.6.1.2.1.43.11.1.1'
  SuppliesClassBase     = '1.3.6.1.2.1.43.11.1.1.4.1'  # .<idx>
  SuppliesTypeBase      = '1.3.6.1.2.1.43.11.1.1.5.1'
  SuppliesDescBase      = '1.3.6.1.2.1.43.11.1.1.6.1'
  SuppliesUnitBase      = '1.3.6.1.2.1.43.11.1.1.7.1'
  SuppliesMaxBase       = '1.3.6.1.2.1.43.11.1.1.8.1'
  SuppliesLevelBase     = '1.3.6.1.2.1.43.11.1.1.9.1'
  SuppliesColorIdxBase  = '1.3.6.1.2.1.43.11.1.1.3.1'
  # Colorant table
  ColorantValueBase     = '1.3.6.1.2.1.43.12.1.1.4.1'  # .<colorIdx>
}

# -------- Helpers: detect backends --------
function Test-Command { param([string]$Name) return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue ) }

$HaveProxx   = Test-Command -Name 'Invoke-SnmpWalk'
$HaveNetSnmp = (Test-Command -Name 'snmpwalk') -and (Test-Command -Name 'snmpget')

if (-not $HaveProxx) {
  # Try import if module installed but not loaded
  try { Install-Module Proxx.SNMP -Force -ErrorAction SilentlyContinue; $HaveProxx = Test-Command -Name 'Invoke-SnmpWalk' } catch {}
}

if (-not ($HaveProxx -or $HaveNetSnmp)) {
  Write-Warning "No SNMP backend found. Install Proxx.SNMP (PowerShell) or Net-SNMP (snmpwalk/snmpget). Will try HTTP fallback if provided."
}

# -------- SNMP wrappers --------
function Invoke-Walk {
  param(
    [Parameter(Mandatory)][string]$BaseOid
  )
  if ($HaveProxx) {
    # Proxx.SNMP
    # Example: Invoke-SnmpWalk -IpAddress 192.0.2.10 -Community public -Version 2 -OID 1.3.6.1.2.1.43.11.1.1
    $res = Invoke-SnmpWalk -IP $PrinterIP -Community $Community -Version 2 -UdpPort $SnmpPort -OID $BaseOid -Timeout 5000
    # Normalize to [pscustomobject] with Oid/Value
    return $res | ForEach-Object {
      [pscustomobject]@{ Oid = $_.OID; Value = $_.Value }
    }
  }
  elseif ($HaveNetSnmp) {
    # Net-SNMP
    # Example: snmpwalk -v2c -c public -On -t 5 -r 1 192.0.2.10 1.3.6.1.2.1.43.11.1.1
    $snmpArgs = @('-v2c', '-c', $Community, '-On', '-t', '5', '-r', '1', "${PrinterIP}:$SnmpPort", $BaseOid)
    $text = & snmpwalk @snmpArgs 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($text)) { return @() }
    $lines = $text -split "`r?`n" | Where-Object { $_ -match '^\.' }
    return $lines | ForEach-Object {
      # Format: .1.3... = TYPE: value
      if ($_ -match '^(?<oid>\.\d+(?:\.\d+)*)\s*=\s*[^:]+:\s*(?<val>.+)$') {
        [pscustomobject]@{ Oid = $Matches['oid'].Trim('.'); Value = $Matches['val'].Trim() }
      }
    }
  }
  else {
    return @()
  }
}

function Invoke-Get {
  param([Parameter(Mandatory)][string]$Oid)
  if ($HaveProxx) {
    $r = Invoke-SnmpWalk -IP $PrinterIP -Community $Community -Version 2 -UdpPort $SnmpPort -OID $Oid -Timeout 5000
    $obj = $r | Select-Object -First 1
    if ($obj) { return $obj.Value } else { return $null }
  }
  elseif ($HaveNetSnmp) {
    $snmpArgs = @('-v2c', '-c', $Community, '-On', '-t', '5', '-r', '1', "${PrinterIP}:$SnmpPort", $Oid)
    $text = & snmpget @snmpArgs 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -match '=\s*[^:]+:\s*(?<val>.+)$') { return $Matches['val'].Trim() } else { return $null }
  }
  else { return $null }
}

# -------- Parse walk results into supply rows --------
function Get-PrinterSuppliesFromSnmp {
  # Walk the supplies table and colorant values
  $supWalk = Invoke-Walk -BaseOid $Oids.SuppliesTableBase
  if (-not $supWalk -or $supWalk.Count -eq 0) { return @() }

  $colorValWalk = Invoke-Walk -BaseOid $Oids.ColorantValueBase
  $colorMap = @{}
  foreach ($row in $colorValWalk) {
    # OID: 1.3.6.1.2.1.43.12.1.1.4.1.<colorIdx>
    if ($row.Oid -match '\.43\.12\.1\.1\.4\.1\.(?<cidx>\d+)$') {
      $colorMap[$Matches['cidx']] = ($row.Value -replace '^STRING:\s*', '').Trim(' "')
    }
  }

  # Supplies rows keyed by <idx>
  $rows = @{}
  foreach ($row in $supWalk) {
    # Match leafId and row index
    if ($row.Oid -match '\.43\.11\.1\.1\.(?<leaf>\d+)\.1\.(?<idx>\d+)$') {
      $leaf = [int]$Matches['leaf']; $idx = [int]$Matches['idx']
      if (-not $rows.ContainsKey($idx)) { $rows[$idx] = [ordered]@{} }
      $val = $row.Value -replace '^STRING:\s*', '' -replace '^INTEGER:\s*', '' -replace '^Gauge32:\s*',''
      $val = $val.Trim(' "')
      switch ($leaf) {
        3 { $rows[$idx]['ColorIdx']   = $val }  # prtMarkerSuppliesColorantIndex
        4 { $rows[$idx]['Class']      = [int]$val } # Class: supplyThatIsConsumed(3), receptacleThatIsFilled(4)
        5 { $rows[$idx]['Type']       = [int]$val } # Type: toner(3), tonerCartridge(21), matteToner(35) ...
        6 { $rows[$idx]['Desc']       = $val }
        7 { $rows[$idx]['Unit']       = $val }
        8 { $rows[$idx]['Max']        = [int]$val }
        9 { $rows[$idx]['Level']      = [int]$val }
      }
    }
  }

  $result = @()
  foreach ($k in ($rows.Keys | Sort-Object)) {
    $r = $rows[$k]
    $colorName = $null
    if ($r.ContainsKey('ColorIdx') -and $r['ColorIdx'] -match '^\d+$' -and $colorMap.ContainsKey($r['ColorIdx'])) {
      $colorName = $colorMap[$r['ColorIdx']]
    }
    if (-not $colorName -and $r.ContainsKey('Desc')) {
      if ($r['Desc'] -match '(?i)black')   { $colorName = 'Black' }
      elseif ($r['Desc'] -match '(?i)cyan'){ $colorName = 'Cyan' }
      elseif ($r['Desc'] -match '(?i)magenta'){ $colorName = 'Magenta' }
      elseif ($r['Desc'] -match '(?i)yellow'){ $colorName = 'Yellow' }
      else { $colorName = $r['Desc'] }
    }

    # Negative special values per MIB: unknown/other/notAvailable -> treat as null
    $lvl  = $r['Level']
    $max  = $r['Max']
    $pct  = $null
    if ($lvl -ge 0 -and $max -gt 0) {
      $pct = [int][math]::Round(($lvl * 100.0) / $max, 0)
      if ($pct -gt 100) { $pct = 100 }
    }

    # Build object
    $result += [pscustomobject]@{
      Index       = $k
      Class       = $r['Class']      # 3=consumable, 4=receptacle
      Type        = $r['Type']       # 3=toner, 21=tonerCartridge, 35=matteToner etc.
      Color       = $colorName
      Description = $r['Desc']
      Unit        = $r['Unit']
      MaxCapacity = $max
      LevelUnits  = $lvl
      Percent     = $pct
    }
  }
  return $result
}

# -------- HTTP fallback (best-effort) --------
function Get-PrinterSuppliesFromWeb {
  param([string]$BaseUrl)
  if ([string]::IsNullOrWhiteSpace($BaseUrl)) { return @() }
  try {
    $urls = @(
      "$BaseUrl/DevMgmt/ConsumableConfigDyn.xml",
      "$BaseUrl/DevMgmt/ProductUsageDyn.xml",
      "$BaseUrl/hp/device/ConfigPage/Index"
    )
    foreach ($u in $urls) {
      try {
        $resp = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 8 -ErrorAction Stop
        $text = $resp.Content
        $pairs = @{}
        foreach ($color in 'Black','Cyan','Magenta','Yellow') {
          if ($text -match ("(?is)" + [regex]::Escape($color) + ".*?(\d{1,3})\s*%")) {
            $pairs[$color] = [int][math]::Min(100,[int]$matches[1])
          }
        }
        if ($pairs.Count -gt 0) {
          return $pairs.GetEnumerator() | ForEach-Object {
            [pscustomobject]@{
              Index       = -1
              Class       = 3
              Type        = 3
              Color       = $_.Key
              Description = $_.Key + ' (web)'
              Unit        = 'percent'
              MaxCapacity = 100
              LevelUnits  = $_.Value
              Percent     = $_.Value
            }
          }
        }
      } catch {}
    }
  } catch {}
  return @()
}

# -------- Collect data --------
$supplies = @()
if ($HaveProxx -or $HaveNetSnmp) {
  $supplies = Get-PrinterSuppliesFromSnmp
}
if (($supplies | Where-Object { $_.Percent -ne $null }).Count -eq 0 -and -not [string]::IsNullOrWhiteSpace($FallbackUrl)) {
  $supplies = Get-PrinterSuppliesFromWeb -BaseUrl $FallbackUrl
}

if ($supplies.Count -eq 0) {
  Write-Error "No supply data retrieved via SNMP or HTTP. Verify SNMP v2 is enabled and accessible from this host (UDP $SnmpPort), or provide a working status URL."
  Invoke-Get -Oid $Oids.PrinterStatus
  return
}

# Only toner-like items (type = toner/tonerCartridge/matteToner OR description contains 'toner')
$tonerTypes = @(3,21,35)
$toners = $supplies | Where-Object {
  ($_.Type -in $tonerTypes) -or ($_.Description -match '(?i)toner')
}

if ($toners.Count -eq 0) {
  # Fall back to everything if we couldn't classify
  $toners = $supplies
}

# Fill missing percents with 0 to render, but keep null for alert logic
$displayRows = $toners | ForEach-Object {
  $pct = $_.Percent
  if ($null -eq $pct -and $_.MaxCapacity -gt 0 -and $_.LevelUnits -ge 0) {
    $pct = [int][math]::Round(($_.LevelUnits * 100.0) / $_.MaxCapacity, 0)
  }
  if ($null -eq $pct) { $pct = 0 }
  $_ | Add-Member -NotePropertyName DisplayPct -NotePropertyValue $pct -Force
  $_
}

$low = $displayRows | Where-Object { $_.Percent -ne $null -and $_.Percent -le $ThresholdPct }

# -------- HTML email --------
function Get-ColorHex {
  param([string]$Name)
  switch -Regex ($Name) {
    '(?i)black'   { '#111111' ; break }
    '(?i)cyan'    { '#0095ff' ; break }
    '(?i)magenta' { '#ff00aa' ; break }
    '(?i)yellow'  { '#f2c200' ; break }
    default       { '#777777' }
  }
}

$cards = ''
foreach ($row in $displayRows) {
  $pct = [int]$row.DisplayPct
  $barPx = [int]([math]::Max(0,[math]::Min(100,$pct)) * 2) # 200px wide track
  $hex = Get-ColorHex $row.Color
  $lbl = [System.Web.HttpUtility]::HtmlEncode(($row.Color) ? $row.Color : $row.Description)
  $sub = [System.Web.HttpUtility]::HtmlEncode(($row.Description) ? $row.Description : '')
  $cards += @"
  <div style="flex:1 1 220px; max-width:260px; min-width:220px; margin:8px; padding:12px; border-radius:12px; box-shadow:0 2px 8px rgba(0,0,0,.08); border:1px solid #eee; background:#fff;">
    <div style="font-weight:700; font-size:14px; color:#333; margin-bottom:4px;">$lbl</div>
    <div style="font-size:12px; color:#666; margin-bottom:10px;">$sub</div>
    <div style="height:18px; width:200px; background:#f1f3f5; border:1px solid #e5e7eb; border-radius:9px; overflow:hidden;">
      <div style="height:100%; width:${barPx}px; background:linear-gradient(90deg, $hex 0%, #222 100%); opacity:0.9;"></div>
    </div>
    <div style="margin-top:6px; font-size:13px; color:#333;">$pct%</div>
  </div>
"@
}

$alertHeader = ($low.Count -gt 0) ? "<div style='color:#c0392b;font-weight:700;'>Threshold met: ≤ $ThresholdPct%</div>" : "<div style='color:#2e7d32;font-weight:700;'>All above threshold ($ThresholdPct%)</div>"

$body = @"
<html>
<body style="margin:0;padding:0;background:#f8fafc;font-family:Segoe UI,Arial,Helvetica,sans-serif;">
  <div style="max-width:940px;margin:16px auto;padding:16px;">
    <div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;padding:16px 16px 8px 16px;box-shadow:0 2px 10px rgba(0,0,0,.05);">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
        <div style="font-size:18px;font-weight:700;color:#111;">Toner Status — $([System.Web.HttpUtility]::HtmlEncode($PrinterIP))</div>
        $alertHeader
      </div>
      <div style="display:flex;flex-wrap:wrap;align-items:stretch;">$cards</div>
      <div style="margin-top:12px;font-size:12px;color:#666;">
        Retrieved via $(if ($HaveProxx -or $HaveNetSnmp) { 'SNMP (Printer-MIB)' } else { 'HTTP fallback' }) on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss').
      </div>
    </div>
  </div>
</body>
</html>
"@

# -------- Send email if threshold hit --------
$subject =
  if ($low.Count -gt 0) {
    "TONER LOW [$PrinterIP]: " + ($low | ForEach-Object { ($_.Color ? $_.Color : $_.Description) + " " + ($_.Percent -as [int]) + "%" }) -join ' | '
  } else {
    "Toner status [$PrinterIP]: All above ${ThresholdPct}%"
  }


<#
  $mailParams = @{
    From       = $From
    To         = $To
    Subject    = $subject
    Body       = $body
    BodyAsHtml = $true
    SmtpServer = $SmtpServer
    Port       = $SmtpPort
    SmtpClient = $SmtpClient
  }
  if ($UseSsl) { $mailParams['UseSsl'] = $true }
  if ($UseAuth) { $mailParams['Credential'] = $SmtpCred }
  Send-MailMessage @mailParams
  Write-Host "Email sent: $subject"
} catch {
  Write-Warning "Failed to send email: $($_.Exception.Message)"
  # Still output to console for visibility
  Write-Output $body
}
#>

Import-Module Microsoft.Graph.Users.Actions

$mailparams = @{
	message = @{
		subject = "$subject"
		body = @{
			contentType = "html"
			content = "$body"
		}
		toRecipients = @(
			@{
				emailAddress = @{
					address = "$To"
				}
			}
		)
		ccRecipients = @(
			@{
				emailAddress = @{
					address = ""
				}
			}
		)
	}
	saveToSentItems = "false"
}

try {
    # A UPN can also be used as -UserId.
    $SmtpCred
    $ForwardEmail = Send-MgUserMail -UserId $UserId -BodyParameter $mailparams
    
    if ($ForwardEmail) {
    Write-Host "Email sent: $subject"
    }
} catch {
  Write-Warning "Failed to send email: $($_.Exception.Message)"
  # Still output to console for visibility
  Write-Output $body
}

Write-Host "Done."
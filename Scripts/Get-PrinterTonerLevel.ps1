<#
.SYNOPSIS
  Interactive toner-level monitor via SNMPv2 with HTML email alert.

.DESCRIPTION
  Prompts for printer details and thresholds.
  Defaults to HP toner OIDs but allows custom mapping.
#>

# Prompt for configuration
$PrinterIP = Read-Host 'Printer IP or hostname'
$Community = Read-Host 'SNMP community string'
$ThresholdPercent = [int](Read-Host 'Alert threshold percentage (e.g., 20)')
$From = Read-Host 'Email From address'
$To = Read-Host 'Email To address'
$SmtpServer = Read-Host 'SMTP server address'
$SmtpPort = [int](Read-Host 'SMTP port (e.g., 587)')
$FallbackUrl = Read-Host 'Fallback HTML status URL (leave blank if none)'

Write-Host "`nUse default HP toner OIDs? (Y/n)"
if ((Read-Host).ToLower() -in @('n','no')) {
  $Oids = @{}
  Write-Host "Enter custom color name and OID pair; leave color blank to finish."
  while ($true) {
    $color = Read-Host 'Color name (e.g., Black)'
    if ([string]::IsNullOrWhiteSpace($color)) { break }
    $oid = Read-Host "OID for $color"
    $Oids[$color] = $oid
  }
}
else {
  $Oids = @{
    "Black"   = "1.3.6.1.2.1.43.11.1.1.9.1.1"
    "Cyan"    = "1.3.6.1.2.1.43.11.1.1.9.1.2"
    "Magenta" = "1.3.6.1.2.1.43.11.1.1.9.1.3"
    "Yellow"  = "1.3.6.1.2.1.43.11.1.1.9.1.4"
    
  }
}

function Get-TonerLevel {
  param($Color, $Oid)
  try {
    #$out = snmpget -v2c -c $Community $PrinterIP $Oid 2>&1
    #Invoke-SnmpWalk -IP 10.14.0.99 -Version V2 -OIDStart "1.3.6.1.2.1.43.11.1.1.9" -Community public -UDPport 161 -Verbose
    $out = Invoke-SnmpWalk -IP 10.14.0.99 -Version V2 -OIDStart "1.3.6.1.2.1.43.11.1.1.9" -Community public -UDPport 161 -Verbose
    $out

    if ($LASTEXITCODE -ne 0) { throw $out }
    if ($out -match 'INTEGER: (\d+)') {
      return [int]$Matches[1]
    }
    throw "Parse fail"
  }
  catch {
    Write-Host "SNMP failed for ${Color}: $_"
    return $null
  }
}

function Get-ViaWeb {
  param($Color, $Url)
  try {
    $html = Invoke-WebRequest $Url -ErrorAction Stop
    if ($html.Content -match "$Color.*?(\d+)%") {
      return [int]$Matches[1]
    }
    throw "Parse fail"
  }
  catch {
    Write-Host "Fallback failed for ${Color}: $_"
    return $null
  }
}

$results = @{}
foreach ($color in $Oids.Keys) {
  $results[$color] = Get-TonerLevel $color $Oids[$color]
  if ($null -eq $results[$color] -and -not [string]::IsNullOrWhiteSpace($FallbackUrl)) {
    $results[$color] = Get-ViaWeb $color $FallbackUrl
  }
}

$low = $results.GetEnumerator() | Where-Object { $_.Value -ne $null -and $_.Value -le $ThresholdPercent }
if ($low.Count -gt 0) {
  $bars = ''
  foreach ($c in $results.Keys) {
    $lvl = $null -ne $results[$c] ? $results[$c] : 0
    $hex = switch ($c.ToLower()) {
      'black' { '#000' }
      'cyan' { '#00f' }
      'magenta' { '#f0f' }
      'yellow' { '#ff0' }
      default { '#888' }
    }
    $bars += "<div style='margin:4px 0;'><strong>$c</strong>: <div style='background:#ddd;width:200px;'><div style='background:$hex;width:${lvl}px;height:18px;'></div></div> $lvl%</div>"
  }

  $body = @"
<html>
<body style='font-family:sans-serif'>
<h2>Toner Alert for $PrinterIP</h2>
<p>Threshold: $ThresholdPercent %</p>
$bars
</body>
</html>
"@

  $subject = "TONER LOW: " + ($low | ForEach-Object { "$($_.Key) $($_.Value)%" }) -join '; '
  Send-MailMessage -From $From -To $To -Subject $subject -Body $body -BodyAsHtml -SmtpServer $SmtpServer -Port $SmtpPort
}

param(
    [string]$PrinterIP = "10.14.0.99",
    [string]$Community = "public"
)

# OIDs
$OID_SupplyLevel = "1.3.6.1.2.1.43.11.1.1.9"
$OID_SupplyType  = "1.3.6.1.2.1.43.11.1.1.6"

Write-Host "`nFetching printer supplies from $PrinterIP..." -ForegroundColor Cyan

# Walk Supply Types
$types = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $OID_SupplyType -Community $Community -UDPport 161 |
    ForEach-Object { $_.Data }

# Walk Supply Levels
$levels = Invoke-SnmpWalk -IP $PrinterIP -Version V2 -OIDStart $OID_SupplyLevel -Community $Community -UDPport 161 |
    ForEach-Object { $_.Data }

# Combine and display
for ($i = 0; $i -lt $levels.Count; $i++) {
    Write-Host ("{0,-20} {1}" -f $types[$i], $levels[$i]) -ForegroundColor Green
}

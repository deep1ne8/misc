param (
    [string]$PrinterIP = "10.14.0.99"
)

# Function to get toner percentage from printer via SNMP
function Get-TonerLevel {
    param (
        [string]$IP,
        [string]$OID
    )
    try {
        $value = (snmpget -v 2c -c public $IP $OID 2>$null) -replace ".*INTEGER: ", ""
        if ($value -match '^\d+$') { return [int]$value }
        else { return 0 }
    }
    catch {
        return 0
    }
}

# HP M477fnw Toner Level OIDs (percent remaining)
$OIDs = @{
    Black   = "1.3.6.1.2.1.43.11.1.1.9.1.1"
    Cyan    = "1.3.6.1.2.1.43.11.1.1.9.1.2"
    Magenta = "1.3.6.1.2.1.43.11.1.1.9.1.3"
    Yellow  = "1.3.6.1.2.1.43.11.1.1.9.1.4"
}

# Get each toner level
$Black   = Get-TonerLevel -IP $PrinterIP -OID $OIDs.Black
$Cyan    = Get-TonerLevel -IP $PrinterIP -OID $OIDs.Cyan
$Magenta = Get-TonerLevel -IP $PrinterIP -OID $OIDs.Magenta
$Yellow  = Get-TonerLevel -IP $PrinterIP -OID $OIDs.Yellow

# Output in PRTG XML format
@"
<?xml version="1.0" encoding="UTF-8"?>
<prtg>
    <result>
        <channel>Black Cartridge</channel>
        <value>$Black</value>
        <unit>Percent</unit>
        <limitmode>1</limitmode>
        <limitminerror>10</limitminerror>
    </result>
    <result>
        <channel>Cyan Cartridge</channel>
        <value>$Cyan</value>
        <unit>Percent</unit>
        <limitmode>1</limitmode>
        <limitminerror>10</limitminerror>
    </result>
    <result>
        <channel>Magenta Cartridge</channel>
        <value>$Magenta</value>
        <unit>Percent</unit>
        <limitmode>1</limitmode>
        <limitminerror>10</limitminerror>
    </result>
    <result>
        <channel>Yellow Cartridge</channel>
        <value>$Yellow</value>
        <unit>Percent</unit>
        <limitmode>1</limitmode>
        <limitminerror>10</limitminerror>
    </result>
</prtg>
"@

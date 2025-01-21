# Initialize arrays for storing printer information
$ModelInfo = @()
$SerialInfo = @()
$Final = @()

# Retrieve USB printer-related models and serials
$USBPrinterModels = Get-WmiObject Win32_PnPEntity | Where-Object { $_.DeviceID -match "USBPRINT" } | Select-Object -ExpandProperty DeviceID -ErrorAction SilentlyContinue
$USBPrinterSerials = Get-WmiObject Win32_PnPEntity | Where-Object { 
    $_.Description -match "USB 列印支援" -or 
    $_.Description -match "USB Printing Support" -or
    $_.Description -match "USB Composite Device" 
} | Select-Object -ExpandProperty DeviceID -ErrorAction SilentlyContinue

# Handle null values
if (-not $USBPrinterModels -and -not $USBPrinterSerials) {
    Write-Host "No USB printers detected." -ForegroundColor Yellow
    return
}

# Extract model names from DeviceID
foreach ($ModelDevice in $USBPrinterModels) {
    if ($ModelDevice) {
        $ModelName = $ModelDevice.Split("\")[1]
        $ModelInfo += $ModelName
    }
}

# Extract serial numbers from DeviceID
foreach ($SerialDevice in $USBPrinterSerials) {
    if ($SerialDevice) {
        $SerialNumber = $SerialDevice.Split("\")[2]
        if ($SerialNumber -notmatch "&") {
            $SerialInfo += $SerialNumber
        }
    }
}

# Pair models and serials
$MaxLength = [Math]::Max($ModelInfo.Count, $SerialInfo.Count)
for ($i = 0; $i -lt $MaxLength; $i++) {
    $Final += [PSCustomObject]@{
        Model  = $ModelInfo[$i] -as [string]
        Serial = $SerialInfo[$i] -as [string]
    }
}

# Create a WMI class for USB Printer Details (if it doesn't exist)
$ClassExists = Get-WmiObject Win32_USBPrinterDetails -ErrorAction SilentlyContinue
if ($ClassExists) {
    Remove-WmiObject -Class Win32_USBPrinterDetails
}

$WMIClass = New-Object System.Management.ManagementClass("root\cimv2", [String]::Empty, $null)
$WMIClass["__CLASS"] = "Win32_USBPrinterDetails"
$WMIClass.Qualifiers.Add("Static", $true)
$WMIClass.Properties.Add("Model", [System.Management.CimType]::String, $false).Qualifiers.Add("read", $true)
$WMIClass.Properties.Add("Serial", [System.Management.CimType]::String, $false).Qualifiers.Add("key", $true).Qualifiers.Add("read", $true)
$WMIClass.Put()

# Populate the WMI class with printer details
foreach ($Printer in $Final) {
    if ($Printer.Model -and $Printer.Serial) {
        Set-WmiInstance -Path "\\.\root\cimv2:Win32_USBPrinterDetails" -Arguments @{
            Model  = $Printer.Model
            Serial = $Printer.Serial
        }
    }
}

# Output final printer details in a formatted table
if ($Final.Count -gt 0) {
    $Final | Format-Table -AutoSize
} else {
    Write-Host "No USB printer information available to display." -ForegroundColor Yellow
}

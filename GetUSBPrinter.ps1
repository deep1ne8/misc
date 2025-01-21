# Initialize arrays for printer details
$ModelInfo = @()
$SerialInfo = @()
$Final = @()

# Retrieve USB printer-related data
$USBPrinterModels = Get-WmiObject Win32_PnPEntity | Where-Object { $_.DeviceID -match "USBPRINT" } | Select-Object -ExpandProperty DeviceID -ErrorAction SilentlyContinue
$USBPrinterSerials = Get-WmiObject Win32_PnPEntity | Where-Object { 
    $_.Description -match "USB 列印支援" -or 
    $_.Description -match "USB Printing Support" -or 
    $_.Description -match "USB Composite Device" 
} | Select-Object -ExpandProperty DeviceID -ErrorAction SilentlyContinue

# Handle null results
if (-not $USBPrinterModels -and -not $USBPrinterSerials) {
    Write-Host "No USB printers detected." -ForegroundColor Yellow
    return
}

# Extract model information
foreach ($ModelDevice in $USBPrinterModels) {
    if ($ModelDevice) {
        try {
            $ModelName = $ModelDevice.Split("\")[1]
            $ModelInfo += $ModelName
        } catch {
            Write-Warning "Unable to extract model from DeviceID: $ModelDevice"
        }
    }
}

# Extract serial numbers
foreach ($SerialDevice in $USBPrinterSerials) {
    if ($SerialDevice) {
        try {
            $SerialNumber = $SerialDevice.Split("\")[2]
            if ($SerialNumber -notmatch "&") {
                $SerialInfo += $SerialNumber
            }
        } catch {
            Write-Warning "Unable to extract serial from DeviceID: $SerialDevice"
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

# Check if WMIClass creation is successful
try {
    # Remove existing class if it exists
    $ClassExists = Get-WmiObject Win32_USBPrinterDetails -ErrorAction SilentlyContinue
    if ($ClassExists) {
        Remove-WmiObject -Class Win32_USBPrinterDetails
    }

    # Create WMI class for USB Printer Details
    $WMIClass = New-Object System.Management.ManagementClass("root\cimv2", [String]::Empty, $null)
    $WMIClass["__CLASS"] = "Win32_USBPrinterDetails"
    $WMIClass.Qualifiers.Add("Static", $true)
    $WMIClass.Properties.Add("Model", [System.Management.CimType]::String, $false).Qualifiers.Add("read", $true)
    $WMIClass.Properties.Add("Serial", [System.Management.CimType]::String, $false).Qualifiers.Add("key", $true).Qualifiers.Add("read", $true)
    $WMIClass.Put()
} catch {
    Write-Error "Failed to create or modify the WMI class: $_"
    return
}

# Populate WMI class with data
foreach ($Printer in $Final) {
    if ($Printer.Model -and $Printer.Serial) {
        try {
            Set

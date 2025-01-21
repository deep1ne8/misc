$ModelInfo = @()
$SerialInfo = @()
$FullInfo = @{}
$Final=@()
$USBPrinterModels =  Get-WmiObject Win32_PnPEntity | Where-Object {$_.DeviceID -Match "USBPRINT"}|select DeviceID
$USBPrinterSerials2 = Get-WmiObject Win32_PnPEntity | Where-Object {$_.Description -Match "USB 列印支援" -or $_.Description -Match "USB Printing Support"}|select DeviceID
$USBPrinterSerials = Get-WmiObject Win32_PnPEntity | Where-Object {$_.Description -Match "USB Composite Device"}|select DeviceID

Foreach ($USBPrinterModel in $USBPrinterModels)

{
    $ModelFull = $USBPrinterModel.DeviceID
    $Model = @{}
    $Model.model += ($ModelFull.Split("\"))[1]
    $ModelInfo += $Model
}

Foreach ($USBPrinterSerial in $USBPrinterSerials)
{
    $SerialFull = $USBPrinterSerial.DeviceID
    $Serial = @{}
    $Serial.serial += $SerialFull.Split("\")[2]
    If($Serial.serial -notmatch "&")
    {
    $SerialInfo += $Serial
    }
    
}
Foreach ($USBPrinterSerial2 in $USBPrinterSerials2)
{
    $SerialFull2 = $USBPrinterSerial2.DeviceID
    $Serial2 = @{}
    $Serial2.serial += $SerialFull2.Split("\")[2]
    If($Serial2.serial -notmatch "&")
    {
    $SerialInfo += $Serial2
    }
}
If ($ModelInfo.model.GetType().name -eq "String") {
 $Final += new-object psobject -Property @{
                   Model=$ModelInfo.model
                   Serial=$SerialInfo.serial
                   }
}
ElseIf ($ModelInfo.model.GetType().name -ne "String"){                   
$MaxLength = [Math]::Max($ModelInfo.Length, $SerialInfo.Length)
for ($loop_index = 0; $loop_index -lt $MaxLength; $loop_index++)
{ 
 $Final += new-object psobject -Property @{
                   Model=$ModelInfo.model[$loop_index]
                   Serial=$SerialInfo.serial[$loop_index]
                         }
    }
}
$Class = Get-WmiObject Win32_USBPrinterDetails -ErrorAction SilentlyContinue
If ($Class) {Remove-WmiObject -Class Win32_USBPrinterDetails}

$WMIClass = New-Object System.Management.ManagementClass("root\cimv2", [String]::Empty, $null);
$WMIClass["__CLASS"] = "Win32_USBPrinterDetails";
$WMIClass.Qualifiers.Add("Static", $true)
$WMIClass.Properties.Add("Model", [System.Management.CimType]::String, $false)
$WMIClass.Properties["Model"].Qualifiers.Add("read", $true)
$WMIClass.Properties.Add("Serial", [System.Management.CimType]::String, $false)
$WMIClass.Properties["Serial"].Qualifiers.Add("key", $true)
$WMIClass.Properties["Serial"].Qualifiers.Add("read", $true)
$WMIClass.Put()

ForEach ($FInfo in $Final) {
    [void](Set-WmiInstance -Path \\.\root\cimv2:Win32_USBPrinterDetails -Arguments @{Model=$FInfo.model; Serial=$FInfo.serial})
}
$final|FT

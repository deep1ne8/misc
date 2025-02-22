# Prompt for ZIP file upload
$UploadedZipDriverFile = Read-Host "Please enter the full path of the ZIP file (e.g., C:\Path\To\Driver.zip)"

# Validate the uploaded file
if (-not (Test-Path -Path $UploadedZipDriverFile)) {
    Write-Error "The file path provided does not exist. Please check and try again."
    return
}

Write-Output "ZIP file path provided: $UploadedZipDriverFile"

# Variables
$ExtractedFolderPath = "C:\Temp\ExtractedPrinterDriver"  # Path where files will be extracted
$PrinterIP = Read-Host "Enter the Printer IP Address (e.g., 192.168.1.100)"
$PrinterName = Read-Host "Enter the Printer Name (e.g., MyPrinter)"
$PrinterDriver = Read-Host "Enter the Printer Driver Name (e.g., PrinterDriverName)"

# Ensure the destination folder exists
if (-not (Test-Path -Path $ExtractedFolderPath)) {
    New-Item -ItemType Directory -Path $ExtractedFolderPath -Force | Out-Null
}

# Unzip the uploaded ZIP file
try {
    Expand-Archive -Path $UploadedZipDriverFile -DestinationPath $ExtractedFolderPath -Force
    Write-Output "Successfully extracted the ZIP file to $ExtractedFolderPath."
} catch {
    Write-Error "Failed to extract the ZIP file. Error: $_"
    return
}

# Locate the INF file within the extracted folder
$InfFile = Get-ChildItem -Path $ExtractedFolderPath -Recurse -Filter *.inf | Select-Object -First 1
if (-not $InfFile) {
    Write-Error "No INF file found in the extracted folder."
    return
}

Write-Output "Found INF file: $($InfFile.FullName)"

# Add Printer Port if it doesn't already exist
$PortName = $PrinterIP
$CheckPortExists = Get-PrinterPort -Name $PortName -ErrorAction SilentlyContinue

if (-not $CheckPortExists) {
    try {
        Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP
        Write-Output "Printer port $PortName added successfully."
    } catch {
        Write-Error "Failed to add printer port. Error: $_"
        return
    }
} else {
    Write-Output "Printer port $PortName already exists."
}

# Add the Printer
try {
    Add-Printer -Name $PrinterName -DriverName $PrinterDriver -PortName $PortName
    Write-Output "Printer $PrinterName added successfully."
} catch {
    Write-Error "Failed to add printer. Error: $_"
    return
}

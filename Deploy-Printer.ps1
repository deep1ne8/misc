# Script to deploy printers

# Prompt user for inputs
$PrinterModel = Read-Host "Enter the printer model (e.g., 'Brother MFC-L3770CDW series')"
$PrinterIP = Read-Host "Enter the printer IP address (e.g., '10.0.1.254')"
$DriverZipPath = Read-Host "Enter the full path to the driver ZIP file (e.g., 'C:\Drivers\PrinterDriver.zip')"

# Validate input file existence
if (!(Test-Path -Path $DriverZipPath)) {
    Write-Error "Driver ZIP file not found at '$DriverZipPath'. Please check the path and try again."
    exit
}

# Define working directories
$TempDirectory = "C:\TEMP\$PrinterModel"
$ExecutingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Prepare the working directory
if (!(Test-Path -Path $TempDirectory)) {
    New-Item -Path $TempDirectory -ItemType Directory | Out-Null
} else {
    Write-Host "Cleaning up the temporary directory..."
    Remove-Item -Path "$TempDirectory\*" -Force -Recurse
}

# Extract the driver ZIP file
Write-Host "Extracting driver ZIP file to $TempDirectory..."
Expand-Archive -LiteralPath $DriverZipPath -DestinationPath $TempDirectory -Force

# Search for the INF file
$InfFile = Get-ChildItem -Path $TempDirectory -Recurse -Filter "*.inf" | Select-Object -First 1
if (!$InfFile) {
    Write-Error "No .inf file found in the extracted driver folder. Please ensure the ZIP file contains valid drivers."
    exit
}

# Display the detected INF file
Write-Host "Using INF file: $($InfFile.FullName)"

# Add the printer driver
Write-Host "Installing printer driver..."
PnPUtil /Add-Driver "$($InfFile.FullName)" /Install | Out-Null

# Add the printer port
$PortName = "${PrinterIP}_IP"
Write-Host "Creating printer port: $PortName"
Add-PrinterPort -Name $PortName -PrinterHostAddress $PrinterIP

# Install the printer driver
Write-Host "Installing printer driver to the system..."
Add-PrinterDriver -Name $PrinterModel -InfPath $InfFile.DirectoryName

# Verify driver installation
$InstalledDriver = Get-PrinterDriver | Where-Object { $_.Name -eq $PrinterModel }
if (!$InstalledDriver) {
    Write-Error "Failed to install the printer driver. Exiting script."
    exit
}

# Add the printer
Write-Host "Adding the printer..."
Add-Printer -DriverName $PrinterModel -Name $PrinterModel -PortName $PortName

# Final verification
$InstalledPrinter = Get-Printer | Where-Object { $_.Name -eq $PrinterModel }
if ($InstalledPrinter) {
    Write-Host "Printer '$PrinterModel' successfully added and ready to use!"
} else {
    Write-Error "Failed to add the printer. Please verify the inputs and try again."
}

# Cleanup
Write-Host "Cleaning up temporary files..."
Remove-Item -Path $TempDirectory -Force -Recurse

Write-Host "Script execution completed."

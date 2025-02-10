<#
.SYNOPSIS
Installs a printer driver from a zip file and configures a printer port.

.DESCRIPTION
This script installs a printer driver from a zip file which contains the printer driver and the inf file for the model, and creates a printer port with the specified IP address.

.PARAMETER ZipFilePath
The path to the zip file containing the printer driver and the inf file for the model.

.EXAMPLE
.\InstallPrinter.ps1 -ZipFilePath "C:\Path\To\PrinterDriver.zip"
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$true)]
    [string]$ZipFilePath
)

# Validate the input file
if (!(Test-Path -Path $ZipFilePath)) {
    Write-Host "The file path provided does not exist. Please check and try again." -ForegroundColor Red
    exit
}

# Prompt for the printer IP address
$PrinterIP = Read-Host "Enter the printer IP address (e.g., '10.0.1.254')"

# Define working directories
$TempDirectory = "C:\TEMP\PrinterDriver"
$ExecutingScriptDirectory = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Prepare the working directory
if (!(Test-Path -Path $TempDirectory)) {
    New-Item -Path $TempDirectory -ItemType Directory | Out-Null
} else {
    Write-Host "Cleaning up the temporary directory..." -ForegroundColor Yellow
    Remove-Item -Path "$TempDirectory\*" -Force -Recurse
}

# Extract the driver ZIP file
Write-Host "Extracting driver ZIP file to $TempDirectory..." -ForegroundColor Cyan
Expand-Archive -LiteralPath $ZipFilePath -DestinationPath $TempDirectory -Force

# Search for the INF file
$InfFile = Get-ChildItem -Path $TempDirectory -Recurse -Filter "*.inf" | Select-Object -First 1
if (!$InfFile) {
    Write-Host "No .inf file found in the extracted driver folder. Please ensure the ZIP file contains valid drivers." -ForegroundColor Red
    exit
}

# Display the detected INF file
Write-Host "Using INF file: $($InfFile.FullName)" -ForegroundColor Green

# Add the printer driver
Write-Host "Installing printer driver..." -ForegroundColor Cyan
PnPUtil /Add-Driver "$($InfFile.FullName)" /Install | Out-Null

# Create printer port
Write-Host "Creating printer port for IP: $PrinterIP" -ForegroundColor Cyan
Add-PrinterPort -Name "IP_$PrinterIP" -PrinterHostAddress $PrinterIP

# Verify IP connectivity
Write-Host "Pinging the printer IP to verify connectivity..." -ForegroundColor Cyan
if (Test-Connection -ComputerName $PrinterIP -Count 2 -Quiet) {
    Write-Host "Successfully pinged the printer IP: $PrinterIP" -ForegroundColor Green
} else {
    Write-Host "Failed to ping the printer IP: $PrinterIP. Please check the network connection." -ForegroundColor Red
}


# Define variables
$odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe"
$odtInstallerPath = "$env:TEMP\odt_setup.exe"
$odtExtractPath = "$env:TEMP\ODT"
$teamsConfigPath = "$odtExtractPath\teams.xml"

# Step 1: Download the Office Deployment Tool (ODT)
Write-Host "Downloading the Office Deployment Tool..." -ForegroundColor Cyan
Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $odtInstallerPath

# Step 2: Extract the ODT
Write-Host "Extracting the Office Deployment Tool..." -ForegroundColor Cyan
Start-Process -FilePath $odtInstallerPath -ArgumentList "/extract:$odtExtractPath /quiet" -Wait

# Step 3: Create the XML configuration file for Teams
Write-Host "Creating Teams installation configuration..." -ForegroundColor Cyan
$teamsConfig = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="Teams">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Property Name="PinIconsToTaskbar" Value="TRUE" />
</Configuration>
"@
Set-Content -Path $teamsConfigPath -Value $teamsConfig

# Step 4: Install Teams using the ODT
Write-Host "Installing Microsoft Teams..." -ForegroundColor Cyan
Start-Process -FilePath "$odtExtractPath\setup.exe" -ArgumentList "/configure $teamsConfigPath" -Wait

# Step 5: Verify Teams installation
Write-Host "Verifying Teams installation..." -ForegroundColor Cyan
$teamsExePath = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"
if (Test-Path $teamsExePath) {
    Write-Host "Microsoft Teams has been installed successfully." -ForegroundColor Green
} else {
    Write-Host "Microsoft Teams installation failed." -ForegroundColor Red
}
return
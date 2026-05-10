# Define variables
$odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe"
$odtInstallerPath = "$env:TEMP\odt_setup.exe"
$odtExtractPath = "$env:TEMP\ODT"
$teamsConfigPath = "$odtExtractPath\teams.xml"
$teamsExePath = "$env:LOCALAPPDATA\Microsoft\Teams\Update.exe"

# Step 1: Download the Office Deployment Tool (ODT)
Write-Host "Downloading the Office Deployment Tool..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $odtInstallerPath
} catch {
    Write-Host "Failed to download Office Deployment Tool: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if (-not(Test-Path $odtExtractPath)) {
    try {
        mkdir $odtExtractPath
    } catch {
        Write-Host "Failed to create directory for ODT extraction: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# Step 2: Extract the ODT
Write-Host "Extracting the Office Deployment Tool..." -ForegroundColor Cyan
try {
    Start-Process -FilePath $odtInstallerPath -ArgumentList "/extract:$odtExtractPath /quiet" -Wait
} catch {
    Write-Host "Failed to extract Office Deployment Tool: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 3: Create the XML configuration file for Teams
Write-Host "Creating Teams installation configuration..." -ForegroundColor Cyan
$teamsConfig = @"
<Configuration>
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="O365ProPlusRetail">
            <Language ID="en-us" />
            <ExcludeApp ID="Access" />
            <ExcludeApp ID="Excel" />
            <ExcludeApp ID="OneDrive" />
            <ExcludeApp ID="OneNote" />
            <ExcludeApp ID="Outlook" />
            <ExcludeApp ID="PowerPoint" />
            <ExcludeApp ID="Publisher" />
            <ExcludeApp ID="Word" />
        </Product>
    </Add>
    <Display Level="None" AcceptEULA="TRUE" />
    <Property Name="AUTOACTIVATE" Value="1" />
    <Property Name="SharedComputerLicensing" Value="1" />
    <Property Name="DeviceBasedLicensing" Value="1" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Property Name="PinIconsToTaskbar" Value="TRUE" />
</Configuration>
"@
try {
    Set-Content -Path $teamsConfigPath -Value $teamsConfig
} catch {
    Write-Host "Failed to create Teams configuration file: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 4: Install Teams using the ODT
Write-Host "Installing Microsoft Teams..." -ForegroundColor Cyan
try {
    Start-Process -FilePath "$odtExtractPath\setup.exe" -ArgumentList "/configure $teamsConfigPath" -Wait -Verbose
} catch {
    Write-Host "Failed to install Microsoft Teams: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 5: Verify Teams installation
Write-Host "Verifying Teams installation..." -ForegroundColor Cyan
if (Test-Path $teamsExePath) {
    Write-Host "Microsoft Teams has been installed successfully." -ForegroundColor Green
} else {
    Write-Host "Microsoft Teams installation failed." -ForegroundColor Red
}
return

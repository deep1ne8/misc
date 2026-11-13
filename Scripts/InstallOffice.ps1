# Define variables
#$odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe"
# Official evergreen Office Deployment Tool bootstrapper (signature verified below).
$odtDownloadUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
$GitHubConfigUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/ODTTool/MSOffice.xml"
$odtExtractPath = "$env:TEMP\ODT"
$odtInstallerPath = "$odtExtractPath\setup.exe"
$MSOfficeConfigFile = "$odtExtractPath\MSOffice.xml"
$MSOfficeExePath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"

# Check if Microsoft Office is already installed
if (Test-Path $MSOfficeExePath) {
    Write-Host "Microsoft Office is already installed." -ForegroundColor Green
    return
}

Write-Host "Installing Microsoft Office..." -ForegroundColor Cyan

# Step 1: Create Directory
if (-not (Test-Path $odtExtractPath)) {
    try {
        New-Item -ItemType Directory -Path $odtExtractPath -Force | Out-Null
    } catch {
        Write-Host "Failed to create directory for ODT extraction: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}


# Step 2: Download the Office Deployment Tool (ODT)
Write-Host "Downloading the Office Deployment Tool..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $odtInstallerPath -UseBasicParsing
    $odtSig = Get-AuthenticodeSignature -FilePath $odtInstallerPath
    if ($odtSig.Status -ne 'Valid' -or $odtSig.SignerCertificate.Subject -notmatch 'O=Microsoft Corporation') {
        Remove-Item $odtInstallerPath -Force -ErrorAction SilentlyContinue
        throw "Office Deployment Tool signature check failed ($($odtSig.Status))"
    }
} catch {
    Write-Host "Failed to download Office Deployment Tool: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 3: Download MSOffices.xml configuration file
Write-Host "Downloading Microsoft Office configuration file..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $GitHubConfigUrl -OutFile $MSOfficeConfigFile -UseBasicParsing
} catch {
    Write-Host "Failed to download Microsoft Office configuration file: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 5: Install Office using the ODT
if (Test-Path $odtInstallerPath) {
    Write-Host "Installing Microsoft Office..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath $odtInstallerPath -ArgumentList "/configure $MSOfficeConfigFile" -Wait -NoNewWindow
    } catch {
        Write-Host "Failed to install Microsoft Office: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
} else {
    Write-Host "Setup.exe not found. Extraction may have failed." -ForegroundColor Red
    return
}

# Step 6: Verify Office installation
Write-Host "Verifying Microsoft Office installation..." -ForegroundColor Cyan
if (Test-Path $MSOfficeExePath) {
    Write-Host "Microsoft Office is installed at: $MSOfficeExePath" -ForegroundColor Green
} else {
    Write-Host "Microsoft Office installation failed." -ForegroundColor Red
}
return

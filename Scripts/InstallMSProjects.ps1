# Define variables
#$odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe"
$odtDownloadUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/ODTTool/setup.exe"
$GitHubConfigUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/ODTTool/MSProjects.xml"
$odtExtractPath = "$env:TEMP\ODT"
$odtInstallerPath = "$odtExtractPath\setup.exe"
$MSProjectConfigFile = "$odtExtractPath\MSProjects.xml"
$MSProjectExePath = "C:\Program Files\Microsoft Office\root\Office16\WINPROJ.EXE"

# Check if Microsoft Project is already installed
if (Test-Path $MSProjectExePath) {
    Write-Host "Microsoft Project is already installed." -ForegroundColor Green
    return
}

Write-Host "Installing Microsoft Project..." -ForegroundColor Cyan

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
    Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $odtInstallerPath
} catch {
    Write-Host "Failed to download Office Deployment Tool: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 3: Download MSProjects.xml configuration file
Write-Host "Downloading Microsoft Project configuration file..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $GitHubConfigUrl -OutFile $MSProjectConfigFile -UseBasicParsing
} catch {
    Write-Host "Failed to download Microsoft Project configuration file: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 5: Install Project using the ODT
if (Test-Path $odtInstallerPath) {
    Write-Host "Installing Microsoft Project..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath $odtInstallerPath -ArgumentList "/configure $MSProjectConfigFile" -Wait -NoNewWindow
    } catch {
        Write-Host "Failed to install Microsoft Project: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
} else {
    Write-Host "Setup.exe not found. Extraction may have failed." -ForegroundColor Red
    return
}

# Step 6: Verify Project installation
Write-Host "Verifying Microsoft Project installation..." -ForegroundColor Cyan
if (Test-Path $MSProjectExePath) {
    Write-Host "Microsoft Project is installed at: $MSProjectExePath" -ForegroundColor Green
} else {
    Write-Host "Microsoft Project installation failed." -ForegroundColor Red
}
return

# Define variables
$odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe"
$GitHubConfigUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/ODTTool/MSProjects.xml"
$odtInstallerPath = "$env:TEMP\odt_setup.exe"
$odtExtractPath = "$env:TEMP\ODT"
$MSProjectConfigFile = "$odtExtractPath\MSProjects.xml"
$MSProjectExePath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PROJECT.EXE" -ErrorAction SilentlyContinue)."(default)"

# Step 1: Download the Office Deployment Tool (ODT)
Write-Host "Downloading the Office Deployment Tool..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $odtInstallerPath
} catch {
    Write-Host "Failed to download Office Deployment Tool: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 2: Create Extraction Directory
if (-not (Test-Path $odtExtractPath)) {
    try {
        New-Item -ItemType Directory -Path $odtExtractPath -Force | Out-Null
    } catch {
        Write-Host "Failed to create directory for ODT extraction: $($_.Exception.Message)" -ForegroundColor Red
        return
    }
}

# Step 3: Extract the ODT
Write-Host "Extracting the Office Deployment Tool..." -ForegroundColor Cyan
try {
    Start-Process -FilePath $odtInstallerPath -ArgumentList "/extract:$odtExtractPath /quiet" -Wait
} catch {
    Write-Host "Failed to extract Office Deployment Tool: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 4: Download MSProjects.xml configuration file
Write-Host "Downloading Microsoft Project configuration file..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $GitHubConfigUrl -OutFile $MSProjectConfigFile
} catch {
    Write-Host "Failed to download Microsoft Project configuration file: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 5: Install Project using the ODT
$setupPath = "$odtExtractPath\setup.exe"
if (Test-Path $setupPath) {
    Write-Host "Installing Microsoft Project..." -ForegroundColor Cyan
    try {
        Start-Process -FilePath $setupPath -ArgumentList "/configure $MSProjectConfigFile" -Wait
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
    Write-Host "MSProjectExePath: $MSProjectExePath" -ForegroundColor Green
    Write-Host "Microsoft Project has been installed successfully." -ForegroundColor Green
} else {
    Write-Host "Microsoft Project installation failed." -ForegroundColor Red
}
return

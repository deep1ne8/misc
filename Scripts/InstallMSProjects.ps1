# Define variables
$odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_18324-20194.exe"
$GitHubParentUrl = "https://github.com/deep1ne/misc/refs/heads/main/"
$MSProjectConfigPath = $GitHubParentUrl + "ODTTool"
$odtInstallerPath = "$env:TEMP\odt_setup.exe"
$odtExtractPath = "$env:TEMP\ODT"
$MSProjectConfigFile = $MSProjectConfigPath + "MSProjects.xml"
$MSProjectExePath = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PROJECT.EXE" -ErrorAction SilentlyContinue)."(default)"


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


# Step 4: Install Project using the ODT
Write-Host "Installing Microsoft Project..." -ForegroundColor Cyan
try {
    Start-Process -FilePath "$odtExtractPath\setup.exe" -ArgumentList "/configure $MSProjectConfigFile" -Wait -Verbose
} catch {
    Write-Host "Failed to install Microsoft Project: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# Step 5: Verify Project installation
Write-Host "Verifying Teams installation..." -ForegroundColor Cyan
if (Test-Path $MSProjectExePath) {
    Write-Host "MSProjectExePath: $MSProjectExePath" -foregroundColor Green
    Write-Host "Microsoft Project has been installed successfully." -ForegroundColor Green
} else {
    Write-Host "Microsoft Project installation failed." -ForegroundColor Red
}
return
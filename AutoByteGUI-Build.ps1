# Step 1: Download Python Installer
Write-Host "Step 1: Downloading Python installer..." -ForegroundColor Cyan
$pythonUrl = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.exe"  # Update to latest if needed
$pythonInstaller = "$env:TEMP\python-latest.exe"
Invoke-WebRequest -Uri $pythonUrl -OutFile $pythonInstaller -Verbose

if (Test-Path $pythonInstaller) {
    Write-Host "Python installer downloaded successfully." -ForegroundColor Green
} else {
    Write-Host "Failed to download Python installer." -ForegroundColor Red
    exit 1
}

# Step 2: Install Python Silently
Write-Host "Step 2: Installing Python silently..." -ForegroundColor Cyan
Start-Process -FilePath $pythonInstaller -ArgumentList "/quiet InstallAllUsers=1 PrependPath=1" -Wait -NoNewWindow

# Step 3: Verify Python Installation
Write-Host "Step 3: Verifying Python installation..." -ForegroundColor Cyan
$pythonVersion = & python --version 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "Python installed: $pythonVersion" -ForegroundColor Green
} else {
    Write-Host "Python installation failed or not found in PATH." -ForegroundColor Red
    exit 1
}

# Step 4: Verify pip Installation
Write-Host "Step 4: Verifying pip installation..." -ForegroundColor Cyan
$pipVersion = & pip --version 2>$null
if ($LASTEXITCODE -eq 0) {
    Write-Host "pip installed: $pipVersion" -ForegroundColor Green
} else {
    Write-Host "pip not found. Attempting to repair with ensurepip..." -ForegroundColor Yellow
    & python -m ensurepip --upgrade
    & python -m pip install --upgrade pip
    $pipVersion = & pip --version 2>$null
    if ($LASTEXITCODE -eq 0) {
        Write-Host "pip installed: $pipVersion" -ForegroundColor Green
    } else {
        Write-Host "pip installation failed." -ForegroundColor Red
        exit 1
    }
}

# Step 5: Download requirements.txt
Write-Host "Step 5: Downloading requirements.txt..." -ForegroundColor Cyan
$requirementsUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/requirements.txt"
$requirementsFile = "$env:TEMP\requirements.txt"
Invoke-WebRequest -Uri $requirementsUrl -OutFile $requirementsFile -Verbose

if (Test-Path $requirementsFile) {
    Write-Host "requirements.txt downloaded successfully." -ForegroundColor Green
} else {
    Write-Host "Failed to download requirements.txt." -ForegroundColor Red
    exit 1
}

# Step 6: Install Python Requirements
Write-Host "Step 6: Installing Python requirements with pip (verbose)..." -ForegroundColor Cyan
Start-Sleep -Seconds 2
& pip install -r $requirementsFile --verbose
if ($LASTEXITCODE -eq 0) {
    Write-Host "All requirements installed successfully." -ForegroundColor Green
} else {
    Write-Host "Some requirements failed to install. Please check the output above." -ForegroundColor Red
    exit 1
}

Write-Host "All steps completed!" -ForegroundColor Green
Start-Sleep -Seconds 2

Write-Host "Running AutoByte GUI..." -ForegroundColor Cyan
Start-Sleep -Seconds 2
& python AutoByteGUI.py
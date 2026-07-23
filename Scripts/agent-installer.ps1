# Enable TLS 1.2 for secure downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Check for administrative privileges
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "`nERROR: This script must be run as an administrator!`n" -ForegroundColor Red
    Write-Host "Right-click the script and select 'Run as administrator'.`n" -ForegroundColor Yellow
    exit 1
}

# Configuration Variables
$liongardUrl = "us5.app.liongard.com"
$accessKeyId = "1608c91df79db5986e44"
$accessKeySecret = "4120668aa5f761ee2ae23dd2ea7866d75a7855f141c466bb7179bf83a7b54ee9"
$environment = "Empire Architectural Metal & Glass"
$installerPath = Join-Path $env:TEMP "LiongardAgent-lts.msi"
$logPath = Join-Path $env:TEMP "liongard_install.log"

Write-Host "Starting Liongard Agent installation..."
try {
    # Download the installer
    Write-Host "Downloading Liongard Agent installer..."
    Invoke-WebRequest -Uri "https://agents.static.liongard.com/LiongardAgent-lts.msi" -OutFile $installerPath
    if (!(Test-Path $installerPath)) {
        throw "Failed to download installer"
    }

    # Prepare installation arguments
    $arguments = @(
        "/i",
        $installerPath,
        "LIONGARDURL=$liongardUrl",
        "LIONGARDACCESSKEY=$accessKeyId",
        "LIONGARDACCESSSECRET=$accessKeySecret",
        "LIONGARDENVIRONMENT=""$($environment)""",
        "/qn",
        "/L*v",
        $logPath
    )

    # Execute installation
    Write-Host "Installing Liongard Agent..."
    $process = Start-Process "msiexec.exe" -ArgumentList $arguments -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -eq 0) {
        Write-Host "Installation completed successfully!"
    } else {
        throw "Installation failed with exit code: $($process.ExitCode)"
    }
} catch {
    Write-Host "Error during installation: $_" -ForegroundColor Red
    if (Test-Path $logPath) {
        Write-Host "Installation log contents:"
        Get-Content $logPath
    }
    exit 1
} finally {
    # Cleanup
    if (Test-Path $installerPath) {
        Remove-Item $installerPath -Force
    }
}

Write-Host "Installation process completed."
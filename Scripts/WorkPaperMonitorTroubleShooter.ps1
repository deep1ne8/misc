using namespace System
using namespace System.Diagnostics
using namespace System.Security.Principal

# Thomson Reuters Workpapers Troubleshooting Script
function Resolve-ThomsonReutersWorkpapers {
    param (
        [switch]$Verbose
    )
    
    # Elevated privilege check
    $currentPrincipal = [WindowsPrincipal]::new([WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([WindowsBuiltInRole]::Administrator)) {
        throw "Requires administrative privileges. Please run as administrator."
    }
    
    # Logging
    $logPath = Join-Path -Path $env:TEMP -ChildPath "ThomsonReutersDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Host "Logging to: $logPath" -ForegroundColor Cyan

    # Clean browser cache
    Write-Host "Clearing browser caches..." -ForegroundColor Yellow
    $browserPaths = @(
        "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
        "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*.default\cache2",
        "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
    )
    foreach ($path in $browserPaths) {
        if (Test-Path $path) {
            Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 2

    # Terminate Thomson Reuters related processes
    Write-Host "Stopping related processes..." -ForegroundColor Yellow
    $trProcesses = @('EXCEL', 'WINWORD', 'OUTLOOK', 'chrome', 'firefox', 'msedge', 'WorkPaperMonitor')
    foreach ($proc in $trProcesses) {
        Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2

    # Restart WorkPaper Monitor service
    Write-Host "Restarting Thomson Reuters Workpapers services..." -ForegroundColor Yellow
    $services = Get-Service -Name "*Workpapers*" -ErrorAction SilentlyContinue
    foreach ($service in $services) {
        Restart-Service -Name $service.Name -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 3

    # Clear local WorkPaper cache
    Write-Host "Clearing WorkPaper Monitor local cache..." -ForegroundColor Yellow
    $wpCachePath = "$env:LOCALAPPDATA\ThomsonReuters\Workpapers"
    if (Test-Path $wpCachePath) {
        Remove-Item -Path "$wpCachePath\*" -Recurse -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Seconds 2

    # Flush DNS & Renew Network
    Write-Host "Flushing DNS and renewing network..." -ForegroundColor Yellow
    ipconfig /flushdns | Out-Null
    ipconfig /release | Out-Null
    ipconfig /renew | Out-Null
    Start-Sleep -Seconds 3

    # Check system dependencies
    Write-Host "Checking required system dependencies..." -ForegroundColor Yellow
    $dotNetInstalled = (Get-WindowsFeature -Name NET-Framework-Core).Installed
    if (-not $dotNetInstalled) {
        Write-Host ".NET Framework missing. Installing..." -ForegroundColor Red
        Install-WindowsFeature -Name NET-Framework-Core -IncludeAllSubFeature
    }
    Start-Sleep -Seconds 2

    # Run DNS resolution check
    Write-Host "Checking DNS resolution..." -ForegroundColor Yellow
    try {
        Resolve-DnsName 'www.thomsonreuters.com' -ErrorAction Stop | Out-File -FilePath $logPath -Append
    } catch {
        Write-Host "DNS resolution failed." -ForegroundColor Red
    }

    Write-Host "Diagnostics complete. Log saved to: $logPath" -ForegroundColor Green
    #Get-Content -Path $logPath
}

# Execute troubleshooting
try {
    Resolve-ThomsonReutersWorkpapers -Verbose
} catch {
    Write-Error "Unhandled exception: $_"
}

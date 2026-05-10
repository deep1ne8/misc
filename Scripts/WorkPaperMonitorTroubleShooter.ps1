using namespace System
using namespace System.Diagnostics
using namespace System.Security.Principal

# Thomson Reuters Workpapers Troubleshooting Script
function Resolve-ThomsonReutersWorkpapers {
    # Elevated privilege check
    $currentPrincipal = [WindowsPrincipal]::new([WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([WindowsBuiltInRole]::Administrator)) {
        throw "Requires administrative privileges. Please run as administrator."
    }

    # Logging mechanism
    $logPath = Join-Path -Path $env:TEMP -ChildPath "ThomsonReutersDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Host "Logging to: $logPath" -ForegroundColor Cyan

    try {
        # Check .NET Framework Installation
        Write-Host "Checking .NET Framework installation..." -ForegroundColor Yellow
        $dotNetInstalled = Get-ItemProperty -Path "HKLM:\\SOFTWARE\\Microsoft\\NET Framework Setup\\NDP\\v4\\Full" -Name Release -ErrorAction SilentlyContinue
        if ($dotNetInstalled -and $dotNetInstalled.Release -ge 528040) {
            Write-Host ".NET Framework 4.8 or later is installed." -ForegroundColor Green
        } else {
            Write-Host ".NET Framework 4.8 is NOT installed. Please install it." -ForegroundColor Red
        }

        # Clean browser cache
        Write-Host "Clearing browser cache..." -ForegroundColor Yellow
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

        # Terminate Thomson Reuters related processes
        Write-Host "Terminating Thomson Reuters related processes..." -ForegroundColor Yellow
        $trProcesses = @('EXCEL', 'WINWORD', 'OUTLOOK', 'chrome', 'firefox', 'msedge')
        foreach ($proc in $trProcesses) {
            Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
        }

        # Restart WorkPaper Monitor Services
        Write-Host "Restarting WorkPaper Monitor services..." -ForegroundColor Yellow
        Get-Service -Name "WorkPaperMonitor" -ErrorAction SilentlyContinue | Restart-Service -Force

        # Clear Local WorkPaper Cache
        Write-Host "Clearing WorkPaper cache..." -ForegroundColor Yellow
        $workpaperCache = "$env:LOCALAPPDATA\Thomson Reuters\WorkPapers\Cache"
        if (Test-Path $workpaperCache) {
            Remove-Item -Path $workpaperCache -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Flush DNS & Renew Network
        Write-Host "Flushing DNS & Renewing Network..." -ForegroundColor Yellow
        ipconfig /flushdns | Out-Null
        Start-Process -FilePath "ipconfig" -ArgumentList "/release" -NoNewWindow -Wait
        Start-Process -FilePath "ipconfig" -ArgumentList "/renew" -NoNewWindow -Wait

        # System diagnostics
        Write-Host "Running system diagnostics..." -ForegroundColor Yellow
        $dnsResult = $null
        try {
            $dnsResult = Resolve-DnsName 'www.thomsonreuters.com' -ErrorAction Stop
        } catch {
            $dnsResult = "DNS resolution failed"
        }

        # Collect diagnostic results
        $diagnosticResults = @{
            NetworkStatus  = (Test-NetConnection -ComputerName "www.thomsonreuters.com").PingSucceeded
            DNSResolution  = $dnsResult
            FirewallStatus = (Get-NetFirewallRule -Enabled True).Count
        }
        $diagnosticResults | ConvertTo-Json | Out-File -FilePath $logPath -Append
    }
    catch {
        $errorDetails = @{
            Message    = $_.Exception.Message
            Position   = $_.InvocationInfo.PositionMessage
            Time       = Get-Date
        }
        $errorDetails | ConvertTo-Json | Out-File -FilePath $logPath -Append
        Write-Error "Critical error during Thomson Reuters Workpapers resolution: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)"
    }

    Write-Host "Diagnostic log saved to: $logPath" -ForegroundColor Green
    Write-Host "Please restart your computer and retry opening Workpapers." -ForegroundColor Yellow
    #Get-Content -Path $logPath
}

# Execution wrapper
try {
    Resolve-ThomsonReutersWorkpapers -Verbose
} catch {
    Write-Error "Unhandled exception in Thomson Reuters Workpapers resolution: $_"
}

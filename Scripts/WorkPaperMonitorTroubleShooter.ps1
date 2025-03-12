using namespace System
using namespace System.Diagnostics
using namespace System.Security.Principal

# Thomson Reuters Workpapers troubleshooting script
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
        Write-Host "Cleaning browser cache..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
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
        
        Write-Host "Clearing Internet Explorer/Edge temporary files..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 255" -NoNewWindow -Wait
        
        Write-Host "Terminating Thomson Reuters related processes..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        $trProcesses = @('EXCEL', 'WINWORD', 'OUTLOOK', 'chrome', 'firefox', 'msedge')
        foreach ($proc in $trProcesses) {
            Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
        }
        
        Write-Host "Flushing DNS & Renewing Network..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        ipconfig /flushdns | Out-Null
        Start-Process -FilePath "ipconfig" -ArgumentList "/release" -NoNewWindow -Wait
        Start-Process -FilePath "ipconfig" -ArgumentList "/renew" -NoNewWindow -Wait
        
        Write-Host "Resetting browser settings..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        $resetCommands = @{
            'Chrome'  = { Remove-Item -Path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Preferences" -Force -ErrorAction SilentlyContinue }
            'Firefox' = { Remove-Item -Path "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*.default\prefs.js" -Force -ErrorAction SilentlyContinue }
            'Edge'    = { Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Preferences" -Force -ErrorAction SilentlyContinue }
        }
        foreach ($cmd in $resetCommands.Values) {
            Invoke-Command $cmd -ErrorAction SilentlyContinue
        }
        
        Write-Host "Resetting registry settings..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        $registryPaths = @(
            'HKCU:\Software\Microsoft\Office\16.0\Common\Internet',
            'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
        )
        foreach ($reg in $registryPaths) {
            Set-ItemProperty -Path $reg -Name 'ProxyEnable' -Value 0 -ErrorAction SilentlyContinue
        }
        
        Write-Host "Running system diagnostics..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        $dnsResult = $null
        try {
            $dnsResult = Resolve-DnsName 'www.thomsonreuters.com' -ErrorAction Stop
        } catch {
            $dnsResult = "DNS resolution failed"
        }
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
        Write-Error "Critical error: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)"
    }
    
    Write-Host "Diagnostic log saved to: $logPath" -ForegroundColor Green
    Write-Host "Please restart your computer and retry opening workpapers." -ForegroundColor Yellow
    Write-Host "Displaying log content..." -ForegroundColor Cyan
    Get-Content -Path $logPath
}

# Execution wrapper
try {
    Resolve-ThomsonReutersWorkpapers -Verbose
}
catch {
    Write-Error "Unhandled exception: $_"
}

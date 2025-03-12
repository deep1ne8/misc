using namespace System
using namespace System.Diagnostics
using namespace System.Security.Principal

# Thomson Reuters workpapers troubleshooting script
function Resolve-ThomsonReutersWorkpapers {
        # Elevated privilege check
        $currentPrincipal = [WindowsPrincipal]::new([WindowsIdentity]::GetCurrent())
        if (-not $currentPrincipal.IsInRole([WindowsBuiltInRole]::Administrator)) {
            # Throw error if not running as administrator
            throw "Requires administrative privileges. Please run as administrator."
        }

        # Logging mechanism
        $logPath = Join-Path -Path $env:TEMP -ChildPath "ThomsonReutersDiagnostics_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

        # Verbose logging
        Write-Verbose "Logging to: $logPath"
    }

        try {
            # Clean browser cache
            $browserPaths = @(
                "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Cache",
                "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*.default\cache2",
                "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Cache"
            )

            # Iterate through browser cache paths
            foreach ($path in $browserPaths) {
                if (Test-Path $path) {
                    # Remove browser cache
                    Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
                }
            }

            # Clear Internet Explorer/Edge temporary internet files
            if (Get-Command -Name Clear-WebBrowserCache -ErrorAction SilentlyContinue) {
                Clear-WebBrowserCache -BrowserName "InternetExplorer", "Edge"
            } else {
                # Use RunDll32 to clear temporary internet files
                Start-Process -FilePath "RunDll32.exe" -ArgumentList "InetCpl.cpl,ClearMyTracksByProcess 255" -NoNewWindow -Wait
            }

            # Terminate Thomson Reuters related processes
            Write-Verbose "Terminating Thomson Reuters related processes..."
            $trProcesses = @('EXCEL', 'WINWORD', 'OUTLOOK', 'chrome', 'firefox', 'msedge')
            foreach ($proc in $trProcesses) {
                # Stop Thomson Reuters processes
                Stop-Process -Name $proc -Force -ErrorAction SilentlyContinue
            }

            # Flush DNS & Renew Network
            Write-Verbose "Flushing DNS & Renewing Network..."
            ipconfig /flushdns | Out-Null
            Start-Process -FilePath "ipconfig" -ArgumentList "/release" -NoNewWindow -Wait
            Start-Process -FilePath "ipconfig" -ArgumentList "/renew" -NoNewWindow -Wait
            ipconfig /release | Out-Null
            ipconfig /renew | Out-Null

            # Reset browser settings
            $resetCommands = @{
                'Chrome'  = { Remove-Item -Path "$env:LOCALAPPDATA\Google\Chrome\User Data\Default\Preferences" -Force -ErrorAction SilentlyContinue }
                'Firefox' = { Remove-Item -Path "$env:LOCALAPPDATA\Mozilla\Firefox\Profiles\*.default\prefs.js" -Force -ErrorAction SilentlyContinue }
                'Edge'    = { Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Edge\User Data\Default\Preferences" -Force -ErrorAction SilentlyContinue }
            }

            # Iterate through browser reset commands
            foreach ($cmd in $resetCommands.Values) {
                # Execute reset command
                Invoke-Command $cmd -ErrorAction SilentlyContinue
            }

            # Reset registry settings
            $registryPaths = @(
                'HKCU:\Software\Microsoft\Office\16.0\Common\Internet',
                'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
            )

            # Iterate through registry paths
            foreach ($reg in $registryPaths) {
                # Set registry value
                Set-ItemProperty -Path $reg -Name 'ProxyEnable' -Value 0 -ErrorAction SilentlyContinue
            }

            # System diagnostics
            Write-Verbose "Running system diagnostics..."
            $dnsResult = $null
            try {
                # Resolve DNS name
                $dnsResult = Resolve-DnsName 'www.thomsonreuters.com' -ErrorAction Stop
            } catch {
                # Set error message
                $dnsResult = "DNS resolution failed"
            }

            # Collect diagnostic results
            $diagnosticResults = @{
                NetworkStatus  = (Test-NetConnection -ComputerName "www.thomsonreuters.com").PingSucceeded
                DNSResolution  = $dnsResult
                FirewallStatus = (Get-NetFirewallRule -Enabled True).Count
            }

            # Log diagnostic results
            $diagnosticResults | ConvertTo-Json | Out-File -FilePath $logPath -Append
        }
        catch {
            # Handle errors
            $errorDetails = @{
                Message    = $_.Exception.Message
                Position   = $_.InvocationInfo.PositionMessage
                Time       = Get-Date
            }

            # Log error details
            $errorDetails | ConvertTo-Json | Out-File -FilePath $logPath -Append
            Write-Error "Critical error during Thomson Reuters workpapers resolution: $($_.Exception.Message) at $($_.InvocationInfo.PositionMessage)"
        }

    end {
        # Final output
        }
    
        # Final output
        Write-Host "Diagnostic log saved to: $logPath" -ForegroundColor Green
        Write-Host "Please restart your computer and retry opening workpapers." -ForegroundColor Yellow
        Write-Host ""
        Get-Content -Path $logPath

        # Verbose output
        if ($Verbose) {
            Get-Content -Path $logPath
}


# Execution wrapper
try {
    Resolve-ThomsonReutersWorkpapers -Verbose
}
catch {
    Write-Error "Unhandled exception in Thomson Reuters workpapers resolution: $_"
    return
}


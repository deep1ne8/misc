#Requires -Version 5.1
<#
.SYNOPSIS
    AutoByte - MSP Automation Toolkit
.DESCRIPTION
    Production-grade PowerShell automation for MSP engineers.
    Covers system health, M365, software deployment, disk, and diagnostics.
.NOTES
    Author  : Earl Quinn
    GitHub  : https://github.com/deep1ne8/misc
    Version : 2.0.0
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Config ──────────────────────────────────────────────────────────────
$Script:Config = @{
    LogPath   = "$env:ProgramData\AutoByte\Logs\AutoByte_$(Get-Date -f 'yyyyMMdd').log"
    TempPath  = "$env:TEMP\AutoByte"
    Version   = '2.0.0'
}
#endregion

#region ── Helpers ─────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK')]$Level = 'INFO')
    $ts   = Get-Date -f 'yyyy-MM-dd HH:mm:ss'
    $line = "[$ts] [$Level] $Message"
    $null = New-Item -ItemType Directory -Force -Path (Split-Path $Script:Config.LogPath)
    Add-Content -Path $Script:Config.LogPath -Value $line
    $colour = @{ INFO='Cyan'; WARN='Yellow'; ERROR='Red'; OK='Green' }[$Level]
    Write-Host $line -ForegroundColor $colour
}

function Invoke-WithLog {
    param([string]$Label, [scriptblock]$Action)
    Write-Log "START: $Label"
    try   { & $Action; Write-Log "DONE:  $Label" -Level OK }
    catch { Write-Log "FAIL:  $Label — $_" -Level ERROR }
}

function Assert-Admin {
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'AutoByte must run as Administrator.'
    }
}

function Show-Banner {
    Clear-Host
    Write-Host @'

  ___         _       ___      _       
 / _ \ _  _ | |_  __| _ ) ___| |_  ___
| (_) | || ||  _|/ _ | _ \/ -_)  _|/ -_)
 \___/ \_,_| \__|\___/___/\___|\__|\___|

  MSP Automation Toolkit  v2.0.0
  github.com/deep1ne8/misc
'@ -ForegroundColor Cyan
}

function Show-Menu {
    param([string]$Title, [string[]]$Options)
    Write-Host "`n  ── $Title ──" -ForegroundColor DarkCyan
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "  [$($i+1)] $($Options[$i])"
    }
    Write-Host "  [0] Back / Exit`n"
    $choice = Read-Host '  Select'
    return $choice
}
#endregion

#region ── 1. System Health ────────────────────────────────────────────────────
function Get-SystemHealth {
    Write-Log 'Running system health check'

    $cpu  = (Get-CimInstance Win32_Processor | Measure-Object -Property LoadPercentage -Average).Average
    $ram  = Get-CimInstance Win32_OperatingSystem
    $ramUsed = [math]::Round(($ram.TotalVisibleMemorySize - $ram.FreePhysicalMemory) / 1MB, 2)
    $ramTotal = [math]::Round($ram.TotalVisibleMemorySize / 1MB, 2)
    $uptime = (Get-Date) - $ram.LastBootUpTime

    $disks = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used -and $_.Free } | ForEach-Object {
        $usedPct = [math]::Round($_.Used / ($_.Used + $_.Free) * 100, 1)
        [pscustomobject]@{ Drive="$($_.Name):"; UsedPct=$usedPct; FreeGB=[math]::Round($_.Free/1GB,1) }
    }

    Write-Host "`n  CPU Load   : $cpu%" -ForegroundColor $(if($cpu -gt 80){'Red'}else{'Green'})
    Write-Host "  RAM Used   : $ramUsed GB / $ramTotal GB"
    Write-Host "  Uptime     : $([math]::Floor($uptime.TotalDays))d $($uptime.Hours)h $($uptime.Minutes)m"
    Write-Host "`n  Disk Usage:"
    $disks | Format-Table -AutoSize | Out-Host

    Write-Log "Health check complete — CPU $cpu%, RAM $ramUsed/$ramTotal GB"
}
#endregion

#region ── 2. Windows Update ───────────────────────────────────────────────────
function Invoke-WindowsUpdate {
    Invoke-WithLog 'Windows Update scan + install' {
        if (-not (Get-Module -ListAvailable PSWindowsUpdate)) {
            Install-Module PSWindowsUpdate -Force -Scope CurrentUser
        }
        Import-Module PSWindowsUpdate
        Get-WindowsUpdate -AcceptAll -Install -AutoReboot:$false | Out-Host
    }
}
#endregion

#region ── 3. Disk Cleanup ─────────────────────────────────────────────────────
function Invoke-DiskCleanup {
    Invoke-WithLog 'Disk cleanup' {
        $targets = @(
            "$env:TEMP\*",
            "$env:SystemRoot\Temp\*",
            "$env:SystemRoot\SoftwareDistribution\Download\*",
            "$env:LOCALAPPDATA\Microsoft\Windows\INetCache\*"
        )
        $freed = 0
        foreach ($path in $targets) {
            $size  = (Get-ChildItem $path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
            Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
            $freed += $size
            Write-Log "Cleaned: $path ($([math]::Round($size/1MB,1)) MB)" -Level OK
        }
        Write-Host "  Total freed: $([math]::Round($freed/1MB,1)) MB" -ForegroundColor Green
    }
}
#endregion

#region ── 4. Software Deployment ─────────────────────────────────────────────
function Install-WingetApp {
    param([string]$AppId, [string]$AppName = $AppId)
    Invoke-WithLog "Install $AppName via winget" {
        winget install --id $AppId --silent --accept-source-agreements --accept-package-agreements
    }
}

function Invoke-SoftwareMenu {
    $apps = [ordered]@{
        '7-Zip'             = '7zip.7zip'
        'Google Chrome'     = 'Google.Chrome'
        'Mozilla Firefox'   = 'Mozilla.Firefox'
        'Notepad++'         = 'Notepad++.Notepad++'
        'VLC'               = 'VideoLAN.VLC'
        'VS Code'           = 'Microsoft.VisualStudioCode'
        'Zoom'              = 'Zoom.Zoom'
        'Teams (new)'       = 'Microsoft.Teams'
        'Adobe Reader'      = 'Adobe.Acrobat.Reader.64-bit'
        'Greenshot'         = 'Greenshot.Greenshot'
    }
    $names = $apps.Keys | ForEach-Object { $_ }
    $sel = Show-Menu 'Software Deployment' $names
    if ($sel -match '^\d+$' -and [int]$sel -ge 1 -and [int]$sel -le $apps.Count) {
        $name = $names[[int]$sel - 1]
        Install-WingetApp -AppId $apps[$name] -AppName $name
    }
}
#endregion

#region ── 5. M365 / Azure AD Diagnostics ─────────────────────────────────────
function Get-M365Status {
    Invoke-WithLog 'M365 tenant diagnostics' {
        $modules = @('MSOnline','AzureAD','ExchangeOnlineManagement','MicrosoftTeams')
        foreach ($mod in $modules) {
            $installed = Get-Module -ListAvailable $mod
            $status    = if ($installed) { '✓ Installed' } else { '✗ Missing' }
            Write-Host "  $status  $mod" -ForegroundColor $(if($installed){'Green'}else{'Yellow'})
        }

        Write-Host "`n  Checking Microsoft Online connectivity..."
        $resp = Invoke-WebRequest -Uri 'https://login.microsoftonline.com' -UseBasicParsing -TimeoutSec 5
        Write-Host "  login.microsoftonline.com : HTTP $($resp.StatusCode)" -ForegroundColor Green
    }
}
#endregion

#region ── 6. OneDrive Files On-Demand ────────────────────────────────────────
function Set-OneDriveFilesOnDemand {
    param([ValidateSet('Enable','Disable')][string]$Action = 'Enable')
    Invoke-WithLog "OneDrive Files On-Demand: $Action" {
        $regPath = 'HKCU:\Software\Microsoft\OneDrive'
        $val     = if ($Action -eq 'Enable') { 1 } else { 0 }
        Set-ItemProperty -Path $regPath -Name 'FilesOnDemandEnabled' -Value $val -Type DWord
        Write-Log "FilesOnDemandEnabled set to $val" -Level OK
    }
}
#endregion

#region ── 7. Dell Command Update ─────────────────────────────────────────────
function Invoke-DellCommandUpdate {
    Invoke-WithLog 'Dell Command Update' {
        $dcu = Get-Command 'dcu-cli.exe' -ErrorAction SilentlyContinue
        if (-not $dcu) { $dcu = 'C:\Program Files\Dell\CommandUpdate\dcu-cli.exe' }
        if (-not (Test-Path $dcu)) { throw 'Dell Command Update not found.' }
        & $dcu /applyUpdates -reboot=disable | Out-Host
    }
}
#endregion

#region ── 8. Network Diagnostics ─────────────────────────────────────────────
function Get-NetworkDiagnostics {
    Invoke-WithLog 'Network diagnostics' {
        $targets = @('8.8.8.8','1.1.1.1','login.microsoftonline.com','outlook.office365.com')
        foreach ($t in $targets) {
            $ping = Test-Connection $t -Count 2 -ErrorAction SilentlyContinue
            if ($ping) {
                $avg = [math]::Round(($ping | Measure-Object ResponseTime -Average).Average)
                Write-Host "  ✓ $t — ${avg}ms" -ForegroundColor Green
            } else {
                Write-Host "  ✗ $t — unreachable" -ForegroundColor Red
            }
        }
        Write-Host "`n  Active adapters:"
        Get-NetAdapter | Where-Object Status -eq 'Up' |
            Select-Object Name, InterfaceDescription, LinkSpeed |
            Format-Table -AutoSize | Out-Host
    }
}
#endregion

#region ── 9. Intune Sync ─────────────────────────────────────────────────────
function Invoke-IntuneSync {
    Invoke-WithLog 'Intune policy sync' {
        $svc = Get-Service -Name 'dmwappushservice' -ErrorAction SilentlyContinue
        if ($svc) { Start-Service $svc -ErrorAction SilentlyContinue }

        $intuneAgent = "$env:ProgramFiles\Microsoft Intune Management Extension\AgentExecutor.exe"
        if (Test-Path $intuneAgent) {
            Start-Process $intuneAgent -ArgumentList '-syncApp' -NoNewWindow -Wait
            Write-Log 'Intune Management Extension sync triggered' -Level OK
        }

        # Trigger MDM enrollment sync via scheduled task
        $task = Get-ScheduledTask -TaskName 'PushLaunch' -ErrorAction SilentlyContinue
        if ($task) { Start-ScheduledTask -TaskName 'PushLaunch' }

        Write-Host '  Intune sync triggered. Check Intune portal for status.' -ForegroundColor Green
    }
}
#endregion

#region ── 10. Export Report ──────────────────────────────────────────────────
function Export-SystemReport {
    Invoke-WithLog 'Generating system report' {
        $out  = "$env:USERPROFILE\Desktop\AutoByte_Report_$(Get-Date -f 'yyyyMMdd_HHmm').html"
        $cpu  = (Get-CimInstance Win32_Processor | Measure-Object LoadPercentage -Average).Average
        $os   = Get-CimInstance Win32_OperatingSystem
        $cs   = Get-CimInstance Win32_ComputerSystem
        $bios = Get-CimInstance Win32_BIOS
        $net  = Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.PrefixOrigin -ne 'WellKnown' }
        $disks= Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Used }

        $diskRows = $disks | ForEach-Object {
            $pct = if (($_.Used + $_.Free) -gt 0) { [math]::Round($_.Used/($_.Used+$_.Free)*100) } else { 0 }
            "<tr><td>$($_.Name):</td><td>$([math]::Round($_.Used/1GB,1)) GB used</td><td>$([math]::Round($_.Free/1GB,1)) GB free</td><td>$pct%</td></tr>"
        }

        $netRows = $net | ForEach-Object {
            "<tr><td>$($_.InterfaceAlias)</td><td>$($_.IPAddress)</td><td>/$($_.PrefixLength)</td></tr>"
        }

        $html = @"
<!DOCTYPE html><html><head><meta charset='utf-8'>
<title>AutoByte System Report</title>
<style>
  body { font-family: Segoe UI, sans-serif; background: #0f1117; color: #e0e0e0; margin: 0; padding: 2rem; }
  h1   { color: #00d4ff; border-bottom: 1px solid #333; padding-bottom: .5rem; }
  h2   { color: #7dd3fc; margin-top: 2rem; }
  table{ width: 100%; border-collapse: collapse; margin-top: .5rem; }
  th   { background: #1e2535; padding: 8px 12px; text-align: left; color: #7dd3fc; }
  td   { padding: 8px 12px; border-bottom: 1px solid #1e2535; }
  .badge { display:inline-block; padding:2px 10px; border-radius:12px; font-size:12px; }
  .ok  { background:#064e3b; color:#34d399; }
  footer{ margin-top:3rem; font-size:12px; color:#555; }
</style></head><body>
<h1>⚡ AutoByte System Report</h1>
<p>Generated: $(Get-Date -f 'yyyy-MM-dd HH:mm:ss') &nbsp;|&nbsp; Host: $($env:COMPUTERNAME) &nbsp;|&nbsp; User: $($env:USERNAME)</p>

<h2>System</h2>
<table><tr><th>Property</th><th>Value</th></tr>
<tr><td>Model</td><td>$($cs.Manufacturer) $($cs.Model)</td></tr>
<tr><td>Serial</td><td>$($bios.SerialNumber)</td></tr>
<tr><td>OS</td><td>$($os.Caption) $($os.OSArchitecture)</td></tr>
<tr><td>Build</td><td>$($os.BuildNumber)</td></tr>
<tr><td>CPU Load</td><td>$cpu%</td></tr>
<tr><td>RAM</td><td>$([math]::Round($os.TotalVisibleMemorySize/1MB,1)) GB total</td></tr>
<tr><td>Last Boot</td><td>$($os.LastBootUpTime)</td></tr>
</table>

<h2>Disk</h2>
<table><tr><th>Drive</th><th>Used</th><th>Free</th><th>Usage%</th></tr>
$($diskRows -join "`n")
</table>

<h2>Network Adapters</h2>
<table><tr><th>Adapter</th><th>IP</th><th>Prefix</th></tr>
$($netRows -join "`n")
</table>

<footer>AutoByte v$($Script:Config.Version) — github.com/deep1ne8/misc</footer>
</body></html>
"@
        $html | Set-Content -Path $out -Encoding UTF8
        Write-Host "  Report saved: $out" -ForegroundColor Cyan
        Start-Process $out
    }
}
#endregion

#region ── Main Menu ──────────────────────────────────────────────────────────
function Start-AutoByte {
    Assert-Admin
    Show-Banner

    $menuItems = @(
        'System Health Check'
        'Windows Update'
        'Disk Cleanup'
        'Software Deployment (Winget)'
        'M365 Module Status'
        'OneDrive Files On-Demand'
        'Dell Command Update'
        'Network Diagnostics'
        'Intune Policy Sync'
        'Export HTML System Report'
    )

    do {
        $sel = Show-Menu 'Main Menu' $menuItems
        switch ($sel) {
            '1'  { Get-SystemHealth }
            '2'  { Invoke-WindowsUpdate }
            '3'  { Invoke-DiskCleanup }
            '4'  { Invoke-SoftwareMenu }
            '5'  { Get-M365Status }
            '6'  {
                $a = Read-Host '  [1] Enable  [2] Disable'
                Set-OneDriveFilesOnDemand -Action $(if($a -eq '2'){'Disable'}else{'Enable'})
            }
            '7'  { Invoke-DellCommandUpdate }
            '8'  { Get-NetworkDiagnostics }
            '9'  { Invoke-IntuneSync }
            '10' { Export-SystemReport }
            '0'  { Write-Host "`n  Goodbye.`n" -ForegroundColor DarkGray; return }
            default { Write-Host '  Invalid option.' -ForegroundColor Yellow }
        }
        Write-Host "`n  Press any key to continue..."
        $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
    } while ($true)
}

Start-AutoByte

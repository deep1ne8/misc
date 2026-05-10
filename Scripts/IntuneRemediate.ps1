<# 
.SYNOPSIS
  Fixes looping IME failures around Dell Command | Endpoint Configure by:
  - Uninstalling/removing remnants
  - (Optionally) clearing Intune Management Extension cache
  - Installing via provided ImmyBot package block
  - Verifying install

.NOTES
  Run as SYSTEM (Intune / Task Scheduler). No external tools required.
#>

param(
  [switch]$ClearIMECache,
  [string]$WorkRoot = "C:\ProgramData\IntuneFixes\DellEndpointConfigure"
)

$ErrorActionPreference = "Stop"
if (!(Test-Path $WorkRoot)) { New-Item -ItemType Directory -Path $WorkRoot -Force | Out-Null }
$TimeStamp = (Get-Date -Format "yyyyMMdd_HHmmss")
$LogPath   = Join-Path $WorkRoot "FixDellEndpointConfigure_$TimeStamp.log"
Start-Transcript -Path $LogPath -Force | Out-Null

function Write-Info([string]$msg){ Write-Host "[INFO] $msg" }
function Write-Warn([string]$msg){ Write-Warning $msg }
function Write-Err ([string]$msg){ Write-Error $msg }

# Find installed products by DisplayName pattern across 32/64-bit uninstall hives
function Get-InstalledProducts {
  param([string[]]$NamePatterns)
  $roots = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
  )
  $foundProducts = @()
  foreach ($r in $roots) {
    if (!(Test-Path $r)) { continue }
    Get-ChildItem $r | ForEach-Object {
      $props = Get-ItemProperty $_.PsPath -ErrorAction SilentlyContinue
      $dn = $props.DisplayName
      if ($dn) {
        foreach ($p in $NamePatterns) {
          if ($dn -like $p) {
            $foundProducts += [PSCustomObject]@{
              DisplayName     = $dn
              DisplayVersion  = $props.DisplayVersion
              UninstallString = $props.UninstallString
              KeyPath         = $_.PsPath
            }
            break
          }
        }
      }
    }
  }
  return $foundProducts
}

function Invoke-MSIUninstall {
  param([string]$UninstallString)
  # Normalize to msiexec /x {GUID}
  $cmd = $null
  if ($UninstallString -match 'MsiExec\.exe.*\{.*\}') {
    $cmd = ($UninstallString -replace '\/I','/X') + ' /qn /norestart'
  } elseif ($UninstallString -match '\{[0-9A-F\-]{36}\}') {
    $guid = ($UninstallString | Select-String -Pattern '\{[0-9A-F\-]{36}\}' -AllMatches).Matches.Value
    $cmd  = "msiexec.exe /x $guid /qn /norestart"
  } else {
    $cmd = "$UninstallString /qn /norestart"
  }
  Write-Info "Uninstall: $cmd"
  $p = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd" -Wait -PassThru
  if ($p.ExitCode -ne 0) { Write-Warn "Uninstall exit code: $($p.ExitCode)" }
}

function Stop-And-DisableService {
  param([string]$NameLike)
  Get-Service | Where-Object { $_.Name -like $NameLike } | ForEach-Object {
    Write-Info "Stopping service: $($_.Name)"
    try { Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue } catch {}
    try { Set-Service -Name $_.Name -StartupType Disabled -ErrorAction SilentlyContinue } catch {}
  }
}

# --- Step 1: Quiesce likely conflicting Dell services
'DellClientManagementService*','Dell*Update*','SupportAssist*' | ForEach-Object {
  Stop-And-DisableService -NameLike $_
}

# --- Step 2: Uninstall existing Dell Endpoint Configure variants
$namePatterns = @(
  'Dell Command*Endpoint*Configure*',
  'Dell*Endpoint*Configure*',
  'Dell Command*Configure*Intune*'
)
$installed = Get-InstalledProducts -NamePatterns $namePatterns
if ($installed.Count -gt 0) {
  Write-Info "Found existing installs:`n$($installed | Format-Table -AutoSize | Out-String)"
  foreach ($app in $installed) {
    if ($app.UninstallString) { Invoke-MSIUninstall -UninstallString $app.UninstallString }
  }
} else {
  Write-Info "No matching legacy installs found."
}

# --- Step 3: Remove leftovers
$leftoverPaths = @(
  'C:\Program Files\Dell\Command*',
  'C:\Program Files (x86)\Dell\Command*',
  'C:\ProgramData\Dell\Command*'
)
foreach ($p in $leftoverPaths) {
  Get-ChildItem $p -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Info "Removing leftover: $($_.FullName)"
    try { Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop } catch { Write-Warn $_.Exception.Message }
  }
}

# Optional: clean registry crumbs
$regCrumbs = @(
  'HKLM:\SOFTWARE\Dell\Command*',
  'HKLM:\SOFTWARE\WOW6432Node\Dell\Command*'
)
foreach ($rk in $regCrumbs) {
  Get-Item $rk -ErrorAction SilentlyContinue | ForEach-Object {
    Write-Info "Removing leftover reg key: $($_.Name)"
    try { Remove-Item -Path $_.PsPath -Recurse -Force } catch { Write-Warn $_.Exception.Message }
  }
}

# --- Step 4: (Optional) Clear IME cache to break retry loop
if ($ClearIMECache.IsPresent) {
  Write-Info "Clearing Intune Management Extension cache..."
  try { Stop-Service -Name IntuneManagementExtension -Force -ErrorAction SilentlyContinue } catch {}
  'C:\Program Files (x86)\Microsoft Intune Management Extension\Content',
  'C:\ProgramData\Microsoft\IntuneManagementExtension' | ForEach-Object {
    if (Test-Path $_) {
      Write-Info "Deleting $_"
      try { Remove-Item -Path $_ -Recurse -Force -ErrorAction Stop } catch { Write-Warn $_.Exception.Message }
    }
  }
  try { Start-Service -Name IntuneManagementExtension -ErrorAction SilentlyContinue } catch {}
}

# --- Step 5: INSTALL (Your requested block, unchanged except wrapping in transcript/logging) ---
try {
  Write-Info "Starting installer via ImmyBot endpoint…"
  $ErrorActionPreference = "Stop";$url = 'https://openapproach.immy.bot/plugins/api/v1/1/installer/latest-download';$InstallerFile = [io.path]::ChangeExtension([io.path]::GetTempFileName(), ".msi");(New-Object System.Net.WebClient).DownloadFile($url, $InstallerFile);$InstallerLogFile = [io.path]::ChangeExtension([io.path]::GetTempFileName(), ".log");$Arguments = " /c msiexec /i `"$InstallerFile`" /qn /norestart /l*v `"$InstallerLogFile`" REBOOT=REALLYSUPPRESS ID=849350a6-d80c-4a05-b5ad-b095c559fa02 ADDR=https://openapproach.immy.bot/plugins/api/v1/1 KEY=8LTD0MKFGhc5hYdER5j4g7J2OK4l/5ZOC3kjHSe7jlU=";Write-Host "InstallerLogFile: $InstallerLogFile";$Process = Start-Process -Wait cmd -ArgumentList $Arguments -Passthru;if ($Process.ExitCode -ne 0) {    Get-Content $InstallerLogFile -ErrorAction SilentlyContinue | Select-Object -Last 200;    throw "Exit Code: $($Process.ExitCode), ComputerName: $($env:ComputerName)"}else {    Write-Host "Exit Code: $($Process.ExitCode)";    Write-Host "ComputerName: $($env:ComputerName)";}
}
catch {
  Write-Err "Installer step failed: $($_.Exception.Message)"
  Write-Err "See installer log path above and transcript: $LogPath"
  Stop-Transcript | Out-Null
  exit 1
}

# --- Step 6: Verify (post-install) ---
Start-Sleep -Seconds 5
$verify = Get-InstalledProducts -NamePatterns $namePatterns
if ($verify.Count -gt 0) {
  Write-Info "✅ Installation verified:"
  $verify | Format-Table DisplayName, DisplayVersion, KeyPath -AutoSize | Out-String | Write-Host
  Write-Info "Transcript: $LogPath"
  Stop-Transcript | Out-Null
  exit 0
} else {
  Write-Err "❌ Dell Endpoint Configure not detected after install attempt."
  Write-Info "Transcript: $LogPath"
  Stop-Transcript | Out-Null
  exit 1
}

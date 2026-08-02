#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Fully removes the Liongard agent (MSI/EXE uninstall + forced remnant cleanup).
#>

$ErrorActionPreference = 'Stop'
$LogPath = Join-Path $env:TEMP "Liongard-Agent-Removal.log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO')
    $Entry = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $Entry
    try { Add-Content -Path $LogPath -Value $Entry -Encoding UTF8 -ErrorAction SilentlyContinue } catch {}
}

function Get-LiongardUninstallEntries {
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )
    Get-ItemProperty -Path $paths -ErrorAction SilentlyContinue |
        Where-Object {
            $_.DisplayName    -match '(?i)Liongard' -or
            $_.Publisher      -match '(?i)Liongard' -or
            $_.InstallLocation -match '(?i)Liongard'
        } |
        Sort-Object PSPath -Unique
}

function Stop-LiongardComponents {
    Write-Log "Stopping Liongard services and processes."

    Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)Liongard' -or $_.DisplayName -match '(?i)Liongard' } |
        ForEach-Object {
            try {
                if ($_.Status -ne 'Stopped') {
                    Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
                    Write-Log "Stopped service: $($_.Name)"
                }
            } catch { Write-Log "Unable to stop service $($_.Name): $($_.Exception.Message)" 'WARN' }
        }

    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '(?i)Liongard' } |
        ForEach-Object {
            try {
                Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
                Write-Log "Stopped process: $($_.ProcessName), PID $($_.Id)"
            } catch { Write-Log "Unable to stop process $($_.ProcessName): $($_.Exception.Message)" 'WARN' }
        }
}

function Invoke-LiongardUninstall {
    param([Parameter(Mandatory)][object]$Entry)

    $DisplayName = $Entry.DisplayName
    $ProductCode = $null
    if ($Entry.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') { $ProductCode = $Entry.PSChildName }

    if ($ProductCode) {
        Write-Log "Attempting MSI uninstall for: $DisplayName ($ProductCode)"
        try {
            $p = Start-Process -FilePath 'msiexec.exe' `
                -ArgumentList "/x `"$ProductCode`" /qn /norestart REBOOT=ReallySuppress" `
                -Wait -PassThru -NoNewWindow
            Write-Log "MSI uninstall exit code: $($p.ExitCode)"
            return $p.ExitCode
        } catch { Write-Log "MSI uninstall failed: $($_.Exception.Message)" 'ERROR'; return 1603 }
    }

    $cmd = $Entry.QuietUninstallString
    if (-not $cmd) { $cmd = $Entry.UninstallString }

    if ($cmd) {
        if ($cmd -match '(?i)msiexec') {
            $cmd = $cmd -replace '(?i)/I(?=\s|\{)', '/X'
            if ($cmd -notmatch '(?i)/q')         { $cmd += ' /qn' }
            if ($cmd -notmatch '(?i)/norestart') { $cmd += ' /norestart REBOOT=ReallySuppress' }
        }
        Write-Log "Running uninstall command for: $DisplayName"
        try {
            $p = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `"$cmd`"" -Wait -PassThru -NoNewWindow
            Write-Log "Uninstall exit code: $($p.ExitCode)"
            return $p.ExitCode
        } catch { Write-Log "Uninstall command failed: $($_.Exception.Message)" 'ERROR'; return 1603 }
    }

    Write-Log "No uninstall command found for: $DisplayName" 'WARN'
    return 1612
}

function Remove-LiongardRemnants {
    Write-Log "Starting forced Liongard cleanup."
    Stop-LiongardComponents

    Get-Service -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)Liongard' -or $_.DisplayName -match '(?i)Liongard' } |
        ForEach-Object {
            try {
                & sc.exe delete $_.Name | Out-Null
                Start-Sleep -Milliseconds 500
                Write-Log "Deleted service: $($_.Name)"
            } catch { Write-Log "Unable to delete service $($_.Name): $($_.Exception.Message)" 'WARN' }
        }

    $folders = @(
        "$env:ProgramFiles\Liongard", "$env:ProgramFiles\Liongard Agent",
        "${env:ProgramFiles(x86)}\Liongard", "${env:ProgramFiles(x86)}\Liongard Agent",
        "$env:ProgramData\Liongard", "$env:ProgramData\LionGard",
        "$env:LOCALAPPDATA\Liongard", "$env:APPDATA\Liongard"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    foreach ($f in $folders) {
        try { Remove-Item -Path $f -Recurse -Force -ErrorAction Stop; Write-Log "Removed folder: $f" }
        catch { Write-Log "Unable to remove folder $f`: $($_.Exception.Message)" 'WARN' }
    }

    @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
      'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall') |
    ForEach-Object {
        Get-ChildItem -Path $_ -ErrorAction SilentlyContinue | ForEach-Object {
            $prop = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($prop -and ($prop.DisplayName -match '(?i)Liongard' -or $prop.Publisher -match '(?i)Liongard' -or $prop.InstallLocation -match '(?i)Liongard')) {
                Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log "Removed uninstall registry entry: $($_.PSChildName)"
            }
        }
    }

    @('HKLM:\SOFTWARE\Liongard', 'HKLM:\SOFTWARE\WOW6432Node\Liongard') | ForEach-Object {
        if (Test-Path $_) {
            try { Remove-Item -Path $_ -Recurse -Force -ErrorAction Stop; Write-Log "Removed registry path: $_" }
            catch { Write-Log "Unable to remove registry path $_`: $($_.Exception.Message)" 'WARN' }
        }
    }
}

# ---- Main ----
Write-Log "Starting dynamic Liongard Agent removal."
Write-Log "Log file: $LogPath"

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) { Write-Log "Administrator privileges are required." 'ERROR'; exit 1 }

$entries = @(Get-LiongardUninstallEntries)
$uninstallFailed = $false

if ($entries.Count -eq 0) {
    Write-Log "No uninstall entries found. Checking for orphaned remnants."
    Remove-LiongardRemnants
} else {
    Write-Log "Found $($entries.Count) Liongard installation entry(ies)."
    foreach ($entry in $entries) {
        Write-Log "Detected: $($entry.DisplayName)"
        Stop-LiongardComponents
        $code = Invoke-LiongardUninstall -Entry $entry
        switch ($code) {
            0        { Write-Log "Uninstall completed successfully." }
            3010     { Write-Log "Uninstall completed. Restart required." 'WARN' }
            1641     { Write-Log "Uninstall completed; restart initiated." 'WARN' }
            1605     { Write-Log "Already removed. Cleaning remnants." 'WARN'; Remove-LiongardRemnants }
            1612     { Write-Log "Installer source unavailable. Forcing cleanup." 'WARN'; Remove-LiongardRemnants }
            default  { Write-Log "Uninstall returned exit code $code. Forcing cleanup." 'WARN'; $uninstallFailed = $true; Remove-LiongardRemnants }
        }
    }
}

$remainingEntries = @(Get-LiongardUninstallEntries)
if ($remainingEntries.Count -gt 0) {
    Write-Log "Entries remain after uninstall. Running final cleanup." 'WARN'
    Remove-LiongardRemnants
}

$remainingServices = @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '(?i)Liongard' -or $_.DisplayName -match '(?i)Liongard' })
$remainingFolders = @(
    "$env:ProgramFiles\Liongard", "$env:ProgramFiles\Liongard Agent",
    "${env:ProgramFiles(x86)}\Liongard", "${env:ProgramFiles(x86)}\Liongard Agent",
    "$env:ProgramData\Liongard", "$env:ProgramData\LionGard"
) | Where-Object { $_ -and (Test-Path $_) }

if ($remainingEntries.Count -eq 0 -and $remainingServices.Count -eq 0 -and $remainingFolders.Count -eq 0) {
    Write-Log "Liongard Agent removed successfully."
    exit 0
}

Write-Log "Cleanup completed, but some remnants may remain." 'ERROR'
exit ($(if ($uninstallFailed) { 1 } else { 0 }))

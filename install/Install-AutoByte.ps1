<#
.SYNOPSIS
    Installs (or updates) AutoByte v3 from a pinned release, verifying every file's SHA-256.

.DESCRIPTION
    Downloads src/AutoByte/files.sha256 from the release tag, then each listed file, checks each
    hash, stages everything in a temp folder and only then moves it into place, so a failed or
    tampered download never leaves a half-installed module. Re-running is safe: an identical
    installed version is left untouched.

    Installs to:
      Module : <Program Files>\WindowsPowerShell\Modules\AutoByte\<version>  (Windows PowerShell 5.1 and 7)
      CLI    : <Program Files>\AutoByte\abyte.ps1  (+ abyte.cmd launcher)
    Without admin rights it installs for the current user instead (Documents\WindowsPowerShell\Modules,
    %LOCALAPPDATA%\AutoByte).

.PARAMETER Version
    Release to install (tag v<Version>). Default: 3.0.0.

.PARAMETER SourcePath
    Install from a local copy of the repository instead of GitHub (offline / USB / testing).

.PARAMETER AddToPath
    Add the CLI folder to PATH (machine PATH when elevated, else user PATH) so 'abyte' works anywhere.

.PARAMETER CreateShortcut
    Add a Start menu shortcut "AutoByte PC check-up" for end users.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Install-AutoByte.ps1 -AddToPath -CreateShortcut
.EXAMPLE
    .\Install-AutoByte.ps1 -Version 3.0.0 -WhatIf

.NOTES
    Version: 1.0.0
    Exit codes: 0 installed or already current, 1 failure.
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidatePattern('^\d+\.\d+\.\d+$')][string]$Version = '3.0.0',
    [string]$SourcePath,
    [ValidatePattern('^[A-Za-z0-9_.\-]+/[A-Za-z0-9_.\-]+$')][string]$Repository = 'deep1ne8/misc',
    [string]$InstallRoot,
    [switch]$AddToPath,
    [switch]$CreateShortcut
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Get-NormalizedSha256([byte[]]$Bytes) {
    $ms = New-Object IO.MemoryStream
    for ($i = 0; $i -lt $Bytes.Length; $i++) {
        if ($Bytes[$i] -eq 13 -and ($i + 1) -lt $Bytes.Length -and $Bytes[$i + 1] -eq 10) { continue }
        $ms.WriteByte($Bytes[$i])
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { ([BitConverter]::ToString($sha.ComputeHash($ms.ToArray())) -replace '-', '').ToLowerInvariant() } finally { $sha.Dispose(); $ms.Dispose() }
}

function Get-SourceByte([string]$RelativePath) {
    if ($SourcePath) { return [IO.File]::ReadAllBytes([IO.Path]::Combine($SourcePath, $RelativePath)) }
    $url = "https://raw.githubusercontent.com/$Repository/v$Version/" + (($RelativePath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/')
    $wc = New-Object Net.WebClient
    try { $wc.DownloadData($url) } finally { $wc.Dispose() }
}

function Test-IsAdmin {
    if ($env:OS -ne 'Windows_NT') { return $false }
    (New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

try {
    $admin = Test-IsAdmin
    if ($InstallRoot) { $moduleBase = [IO.Path]::Combine($InstallRoot, 'Modules'); $binDir = [IO.Path]::Combine($InstallRoot, 'bin') }
    elseif ($admin) { $moduleBase = [IO.Path]::Combine($env:ProgramFiles, 'WindowsPowerShell', 'Modules'); $binDir = [IO.Path]::Combine($env:ProgramFiles, 'AutoByte') }
    else { $moduleBase = [IO.Path]::Combine([Environment]::GetFolderPath('MyDocuments'), 'WindowsPowerShell', 'Modules'); $binDir = [IO.Path]::Combine($env:LOCALAPPDATA, 'AutoByte') }
    $moduleDir = [IO.Path]::Combine($moduleBase, 'AutoByte', $Version)

    Write-Output "AutoByte $Version from $(if ($SourcePath) { $SourcePath } else { "$Repository@v$Version" })"
    $listText = [Text.Encoding]::UTF8.GetString((Get-SourceByte 'src/AutoByte/files.sha256'))
    $entries = @(foreach ($line in ($listText -split "`r?`n")) {
            if ($line -match '^([0-9a-f]{64})  (\S+)$') { [pscustomobject]@{ Sha = $Matches[1]; Path = $Matches[2] } }
        })
    if (-not $entries) { throw 'File list is empty or unreadable.' }
    foreach ($e in $entries) { if ($e.Path -match '(^|/)\.\.(/|$)' -or $e.Path -match '^[\\/]' -or $e.Path -notmatch '^(src/AutoByte/|bin/|datto/)') { throw "Refusing unexpected path in file list: $($e.Path)" } }
    $listVersion = if ($listText -match 'AutoByte (\d+\.\d+\.\d+)') { $Matches[1] } else { '' }
    if ($listVersion -ne $Version) { throw "File list is for version '$listVersion', expected '$Version'." }

    # Already installed and identical?
    $current = $true
    foreach ($e in $entries) {
        $target = if ($e.Path -like 'src/AutoByte/*') { [IO.Path]::Combine($moduleDir, $e.Path.Substring(13)) } else { [IO.Path]::Combine($binDir, (Split-Path -Leaf $e.Path)) }
        if (-not (Test-Path -LiteralPath $target) -or (Get-NormalizedSha256 ([IO.File]::ReadAllBytes($target))) -ne $e.Sha) { $current = $false; break }
    }
    if ($current -and -not $AddToPath -and -not $CreateShortcut) { Write-Output "AutoByte $Version is already installed and verified ($moduleDir)."; exit 0 }

    if (-not $current) {
        $stage = [IO.Path]::Combine([IO.Path]::GetTempPath(), "AutoByte-install-$([guid]::NewGuid().ToString('N'))")
        try {
            $i = 0
            foreach ($e in $entries) {
                $i++
                Write-Progress -Activity "Installing AutoByte $Version" -Status "Verifying $($e.Path)" -PercentComplete ([int](100 * $i / $entries.Count))
                $bytes = Get-SourceByte $e.Path
                $got = Get-NormalizedSha256 $bytes
                if ($got -ne $e.Sha) { throw "SHA-256 mismatch for $($e.Path) (expected $($e.Sha), got $got). Nothing was installed." }
                $dest = [IO.Path]::Combine($stage, ($e.Path -replace '/', [IO.Path]::DirectorySeparatorChar))
                # Staging happens in a private temp folder, so it is exempt from -WhatIf.
                $null = New-Item -ItemType Directory -Path (Split-Path -Parent $dest) -Force -WhatIf:$false
                [IO.File]::WriteAllBytes($dest, $bytes)
            }
            Write-Progress -Activity "Installing AutoByte $Version" -Completed
            if ($PSCmdlet.ShouldProcess($moduleDir, "Install AutoByte $Version")) {
                if (Test-Path -LiteralPath $moduleDir) { Remove-Item -LiteralPath $moduleDir -Recurse -Force }
                $null = New-Item -ItemType Directory -Path $moduleDir -Force
                Copy-Item -Path ([IO.Path]::Combine($stage, 'src', 'AutoByte', '*')) -Destination $moduleDir -Recurse -Force
                $null = New-Item -ItemType Directory -Path $binDir -Force
                Copy-Item -Path ([IO.Path]::Combine($stage, 'bin', '*')) -Destination $binDir -Force
                $dattoStage = [IO.Path]::Combine($stage, 'datto')
                if (Test-Path -LiteralPath $dattoStage) { Copy-Item -Path ([IO.Path]::Combine($dattoStage, '*')) -Destination $binDir -Force }
                $cmd = "@echo off`r`npowershell.exe -NoProfile -ExecutionPolicy Bypass -File `"%~dp0abyte.ps1`" %*`r`nexit /b %ERRORLEVEL%`r`n"
                [IO.File]::WriteAllText([IO.Path]::Combine($binDir, 'abyte.cmd'), $cmd, (New-Object Text.ASCIIEncoding))
                Write-Output "Installed module to $moduleDir"
                Write-Output "Installed CLI to $binDir"
            }
        }
        finally { if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue -WhatIf:$false } }
    }

    if ($AddToPath) {
        $scope = if ($admin) { 'Machine' } else { 'User' }
        $path = [Environment]::GetEnvironmentVariable('Path', $scope)
        if (($path -split ';') -notcontains $binDir) {
            if ($PSCmdlet.ShouldProcess("$scope PATH", "Add $binDir")) {
                [Environment]::SetEnvironmentVariable('Path', ($path.TrimEnd(';') + ';' + $binDir), $scope)
                Write-Output "Added $binDir to the $scope PATH (open a new window to use 'abyte')."
            }
        }
    }
    if ($CreateShortcut) {
        $menu = if ($admin) { [Environment]::GetFolderPath('CommonPrograms') } else { [Environment]::GetFolderPath('Programs') }
        $lnk = [IO.Path]::Combine($menu, 'AutoByte PC check-up.lnk')
        if ($PSCmdlet.ShouldProcess($lnk, 'Create Start menu shortcut')) {
            $shell = New-Object -ComObject WScript.Shell
            $s = $shell.CreateShortcut($lnk)
            $s.TargetPath = 'powershell.exe'
            $s.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$([IO.Path]::Combine($binDir, 'abyte.ps1'))`""
            $s.Description = 'Check this PC for common problems and fix the safe ones'
            $s.Save()
            Write-Output "Created shortcut: $lnk"
        }
    }
    exit 0
}
catch {
    Write-Error "AutoByte install failed: $($_.Exception.Message)" -ErrorAction Continue
    exit 1
}

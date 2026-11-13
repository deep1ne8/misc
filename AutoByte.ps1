<#
.SYNOPSIS
    AutoByte launcher: menu of curated MSP scripts, downloaded from a pinned release and hash-verified.

.DESCRIPTION
    Fetches scripts.manifest.json from the pinned -Ref of the repository, downloads the
    chosen script to a temp file, verifies its SHA-256 against the manifest, runs it with
    the call operator, and deletes the temp file. Nothing is executed with Invoke-Expression,
    and nothing is fetched from the moving 'main' branch unless -AllowUnpinnedRef is given.

    TLS 1.2+ is enabled for this session only. The machine-wide .NET SchUseStrongCrypto
    registry change is applied only with -PersistStrongCrypto (supports -WhatIf).

.PARAMETER Ref
    Release tag (vX.Y.Z) or full 40-character commit SHA to fetch scripts from.

.PARAMETER AllowUnpinnedRef
    Permit a branch name such as 'main' for -Ref (development only; prints a warning).

.PARAMETER Script
    Run one script by manifest name without showing the menu (RMM use). The launcher
    exits with that script's exit code.

.PARAMETER PersistStrongCrypto
    Also set SchUseStrongCrypto=1 under HKLM for .NET 4.x (machine-wide, persistent).

.PARAMETER AllowUnsupported
    Allow -Script to run a script the manifest flags unsupported (e.g. TPM bypass).

.EXAMPLE
    .\AutoByte.ps1
.EXAMPLE
    .\AutoByte.ps1 -Script CheckDriveSpace
.EXAMPLE
    .\AutoByte.ps1 -Ref 0123456789abcdef0123456789abcdef01234567

.NOTES
    Version: 2.0.0
    Exit codes: 0 success, 1 failure (download/verification/script error), otherwise the
    executed script's exit code (e.g. 3010 reboot required) when -Script is used.
#>
#Requires -Version 5.1
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Parameters are read by script-scope functions.')]
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateNotNullOrEmpty()][string]$Ref = 'v2.0.0',
    [switch]$AllowUnpinnedRef,
    [ValidatePattern('^[A-Za-z0-9_.\-]+$')][string]$Script,
    [switch]$PersistStrongCrypto,
    [switch]$AllowUnsupported,
    [ValidatePattern('^[A-Za-z0-9_.\-]+/[A-Za-z0-9_.\-]+$')][string]$Repository = 'deep1ne8/misc',
    # Mirror/testing only. Must be https unless it is localhost.
    [ValidatePattern('^(https://|http://(localhost|127\.0\.0\.1)(:\d+)?)')][string]$BaseUri = 'https://raw.githubusercontent.com'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Curated menu (order and descriptions preserved from v1). Names must exist in the manifest.
$MenuItems = @(
    @{ Name = 'DiskCleaner'; Description = 'Windows C Drive Cleaner Tool' }
    @{ Name = 'NetworkScan'; Description = 'Windows Network Scanner' }
    @{ Name = 'Datto_VM_Deployment_Remediation'; Description = 'Datto VM Deployment Remediation Probe' }
    @{ Name = 'WindowsOnlineRepair'; Description = 'WindowsOnlineRepair' }
    @{ Name = 'GetWindowsEvents'; Description = 'Get Windows Events' }
    @{ Name = 'DellCommandUpdate'; Description = 'Dell Command Update' }
    @{ Name = 'InstallWindowsUpdate'; Description = 'ResetandInstallWindowsUpdate' }
    @{ Name = 'WindowsSystemRepair'; Description = 'WindowsSystemRepair' }
    @{ Name = 'ResetandClearWindowsSearchDB'; Description = 'ResetandClearWindowsSearchdb' }
    @{ Name = 'M365_Office32bitRepair'; Description = 'M365_DeployOffice32bit' }
    @{ Name = 'CheckDriveSpace'; Description = 'CheckDriveSpace' }
    @{ Name = 'InternetSpeedTest'; Description = 'InternetSpeedTest' }
    @{ Name = 'InternetLatencyTest'; Description = 'InternetLatencyTest' }
    @{ Name = 'M365_Office64bitRepair'; Description = 'M365_DeployOffice64bit' }
)

function Enable-SessionTlsSupport {
    $proto = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    if ([enum]::GetNames([Net.SecurityProtocolType]) -contains 'Tls13') { $proto = $proto -bor [Net.SecurityProtocolType]::Tls13 }
    [Net.ServicePointManager]::SecurityProtocol = $proto
}

function Set-StrongCryptoRegistry {
    [CmdletBinding(SupportsShouldProcess)]
    param()
    foreach ($path in 'HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319') {
        $current = Get-ItemProperty -Path $path -Name SchUseStrongCrypto -ErrorAction SilentlyContinue
        if ($current -and $current.SchUseStrongCrypto -eq 1) { continue }
        if ($PSCmdlet.ShouldProcess($path, 'Set SchUseStrongCrypto = 1 (machine-wide)')) {
            if (-not (Test-Path -LiteralPath $path)) { New-Item -Path $path -Force | Out-Null }
            New-ItemProperty -Path $path -Name SchUseStrongCrypto -Value 1 -PropertyType DWord -Force | Out-Null
        }
    }
}

function Get-NormalizedSha256 {
    # sha256 over CRLF->LF-normalised bytes; must match build/Build-ScriptManifest.ps1.
    param([Parameter(Mandatory)][string]$Path)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $ms = New-Object System.IO.MemoryStream
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -eq 13 -and ($i + 1) -lt $bytes.Length -and $bytes[$i + 1] -eq 10) { continue }
        $ms.WriteByte($bytes[$i])
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($ms.ToArray())) -replace '-', '').ToLowerInvariant() }
    finally { $sha.Dispose(); $ms.Dispose() }
}

function Get-RawUrl {
    param([Parameter(Mandatory)][string]$RelativePath)
    $encoded = ($RelativePath -split '/' | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
    return "$($BaseUri.TrimEnd('/'))/$Repository/$Ref/$encoded"
}

function Get-Manifest {
    $url = Get-RawUrl -RelativePath 'scripts.manifest.json'
    try { $m = Invoke-RestMethod -Uri $url -UseBasicParsing -TimeoutSec 30 }
    catch { throw "Could not download the script manifest for ref '$Ref' ($url): $($_.Exception.Message)" }
    if (-not $m.PSObject.Properties['scripts']) { throw "Manifest at $url is not valid." }
    $index = @{}
    foreach ($s in $m.scripts) { $index[$s.name.ToLowerInvariant()] = $s }
    return $index
}

function Invoke-VerifiedScript {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][hashtable]$Manifest,
        [Parameter(Mandatory)][string]$Name
    )
    $entry = $Manifest[$Name.ToLowerInvariant()]
    if (-not $entry) { throw "Script '$Name' is not in the manifest for ref '$Ref'." }
    if ($entry.platform -ne 'windows' -or $entry.path -notlike '*.ps1') { throw "'$Name' is not a Windows PowerShell script." }
    if ($entry.unsupported -and -not $AllowUnsupported) { throw "'$Name' is flagged unsupported. Re-run with -AllowUnsupported if you really intend this." }

    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("AutoByte_{0}_{1}.ps1" -f $entry.name, [guid]::NewGuid().ToString('N'))
    try {
        Invoke-WebRequest -Uri (Get-RawUrl -RelativePath $entry.path) -OutFile $tmp -UseBasicParsing -TimeoutSec 120
        $actual = Get-NormalizedSha256 -Path $tmp
        if ($actual -ne $entry.sha256) {
            throw "SHA-256 mismatch for $($entry.name): expected $($entry.sha256), got $actual. Not executing."
        }
        # Windows PowerShell 5.1 reads BOM-less files as ANSI; add a BOM after verification.
        $bytes = [System.IO.File]::ReadAllBytes($tmp)
        if (-not ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)) {
            [System.IO.File]::WriteAllBytes($tmp, [byte[]](0xEF, 0xBB, 0xBF) + $bytes)
        }
        Write-Host "Verified $($entry.name) ($($entry.sha256.Substring(0, 12))...)." -ForegroundColor Green
        if (-not $PSCmdlet.ShouldProcess($entry.name, 'Run verified script')) { return 0 }
        $global:LASTEXITCODE = 0
        # Run in a child scope with v1 semantics: no StrictMode, 'Continue' error preference,
        # so legacy scripts behave exactly as they did under the old launcher.
        & {
            Set-StrictMode -Off
            $ErrorActionPreference = 'Continue'
            & $args[0]
        } $tmp | Out-Host
        $code = $global:LASTEXITCODE
        if ($null -eq $code) { $code = 0 }
        return [int]$code
    }
    finally {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue -WhatIf:$false }
    }
}

function Show-Banner {
    Clear-Host
    $separator = '=' * 64
    Write-Host "`n$separator" -ForegroundColor Cyan
    Write-Host '                        AutoByte Menu                           ' -BackgroundColor Blue -ForegroundColor White
    Write-Host $separator -ForegroundColor Cyan
    Write-Host "   Source: $Repository @ $Ref`n" -ForegroundColor DarkGray
    Write-Host 'Choose a script to run from the following options:' -BackgroundColor Blue -ForegroundColor White
    Write-Host "$separator`n" -ForegroundColor Cyan
}

# ---- main --------------------------------------------------------------------------
# Dot-sourcing (". .\AutoByte.ps1") loads the functions only; used by the Pester tests.
if ($MyInvocation.InvocationName -eq '.') { return }

$isPinned = $Ref -match '^v\d+\.\d+\.\d+(-[0-9A-Za-z.]+)?$' -or $Ref -match '^[0-9a-f]{40}$'
if (-not $isPinned) {
    if (-not $AllowUnpinnedRef) {
        Write-Error "Ref '$Ref' is not a release tag (vX.Y.Z) or full commit SHA. Use -AllowUnpinnedRef for development." -ErrorAction Continue
        exit 1
    }
    Write-Warning "Running scripts from unpinned ref '$Ref'. Hashes are checked against the manifest at the same ref only."
}

Enable-SessionTlsSupport
if ($PersistStrongCrypto) { Set-StrongCryptoRegistry }

try { $manifest = Get-Manifest }
catch { Write-Error $_.Exception.Message -ErrorAction Continue; exit 1 }

if ($Script) {
    try { exit (Invoke-VerifiedScript -Manifest $manifest -Name $Script) }
    catch { Write-Error $_.Exception.Message -ErrorAction Continue; exit 1 }
}

$exitIndex = $MenuItems.Count + 1
while ($true) {
    Show-Banner
    for ($i = 0; $i -lt $MenuItems.Count; $i++) { Write-Host ("{0}. {1}" -f ($i + 1), $MenuItems[$i].Description) -ForegroundColor Green }
    Write-Host "$exitIndex. Exit`n" -ForegroundColor Red
    $choice = Read-Host "Enter your choice (1-$exitIndex)"
    $n = 0
    if (-not [int]::TryParse($choice, [ref]$n) -or $n -lt 1 -or $n -gt $exitIndex) {
        Write-Host 'Invalid choice, please try again.' -ForegroundColor Red
        Start-Sleep -Seconds 1
        continue
    }
    if ($n -eq $exitIndex) { Write-Host 'Exiting. Goodbye!' -ForegroundColor Yellow; break }
    try {
        $code = Invoke-VerifiedScript -Manifest $manifest -Name $MenuItems[$n - 1].Name
        Write-Host "`nScript finished with exit code $code." -ForegroundColor Cyan
    }
    catch { Write-Host "Failed: $($_.Exception.Message)" -ForegroundColor Red }
    $again = Read-Host "`nReturn to the menu? [1] Yes  [2] Exit"
    if ($again -ne '1') { Write-Host 'Exiting. Goodbye!' -ForegroundColor Yellow; break }
}
exit 0

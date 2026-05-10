<#
.SYNOPSIS
    Download Windows 11 ISO via Fido, mount it, run DISM repair, then unmount.
    Updated for Windows 11 25H2 and latest builds (December 2024).

.DESCRIPTION
    - Auto-detects current Windows version and downloads matching ISO
    - Supports Windows 11 21H2, 22H2, 23H2, 24H2, and 25H2
    - Downloads Fido.ps1 and invokes it to grab the correct ISO
    - Mounts ISO, runs DISM /RestoreHealth, then unmounts

.PARAMETER DestinationDirectory
    Where to store Fido.ps1 and the downloaded ISO. Default: C:\WindowsSetup

.PARAMETER TargetWin
    Always "11" for Windows 11. Auto-detected if omitted.

.PARAMETER Release
    Release build (22H2, 23H2, 24H2, 25H2). Auto-detected if omitted.

.PARAMETER Edition
    Edition (Pro, Home, etc.). Default: Pro

.PARAMETER Arch
    Architecture (x64, arm64). Auto-detected if omitted.

.PARAMETER Language
    Language code. Default: Eng

.EXAMPLE
    .\Win11-DISM-Repair.ps1
    (Auto-detects everything)

.EXAMPLE
    .\Win11-DISM-Repair.ps1 -Release 24H2 -Edition Pro
    (Force 24H2 Pro download)
#>

Function Win11-DISM-Repair {

[CmdletBinding()]
param(
    [string]$DestinationDirectory = "C:\WindowsSetup",
    [ValidateSet("11")][string]$TargetWin = "11",
    [ValidateSet("21H2","22H2","23H2","24H2","25H2")][string]$Release,
    [string]$Edition = "Pro",
    [ValidateSet("x64","arm64")][string]$Arch,
    [string]$Language = "Eng"
)

#Requires -RunAsAdministrator

Write-Host "`n=== Windows 11 DISM Repair Tool ===`n" -ForegroundColor Cyan
Write-Host "Updated for Windows 11 25H2 (December 2024)`n" -ForegroundColor Green

# --- TLS 1.2 Configuration ---
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Host "✓ TLS 1.2 enabled for secure connections" -ForegroundColor Green

$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"
)
foreach ($path in $regPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    Set-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -Type DWord -Force -ErrorAction SilentlyContinue
}
Write-Host "✓ Strong crypto/TLS 1.2 configured in registry`n" -ForegroundColor Green

# --- Helper Functions ---
function Get-SystemInfo {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        
        return @{
            Caption = $os.Caption
            Version = $os.Version
            BuildNumber = [int]$os.BuildNumber
            Architecture = $cs.SystemType
            OSArchitecture = $os.OSArchitecture
        }
    }
    catch {
        throw "Failed to retrieve system information: $_"
    }
}

function Get-DiskSpace {
    param([string]$Drive = 'C')
    try {
        $driveInfo = Get-PSDrive -Name $Drive -ErrorAction Stop
        return @{
            FreeGB = [math]::Round($driveInfo.Free/1GB, 2)
            TotalGB = [math]::Round(($driveInfo.Used + $driveInfo.Free)/1GB, 2)
        }
    }
    catch {
        Write-Warning "Could not retrieve disk space for drive $Drive"
        return $null
    }
}

function Download-WindowsISO {
    param(
        [string]$DestinationDirectory,
        [string]$Win,
        [string]$Rel,
        [string]$Ed,
        [string]$Arch,
        [string]$Lang
    )

    if (-not (Test-Path $DestinationDirectory)) {
        New-Item $DestinationDirectory -ItemType Directory -Force | Out-Null
        Write-Host "✓ Created directory: $DestinationDirectory" -ForegroundColor Green
    }

    $fidoScript = Join-Path $DestinationDirectory 'Fido.ps1'
    $isoName = "Win${Win}_${Rel}_${Lang}_${Arch}.iso"
    $isoPath = Join-Path $DestinationDirectory $isoName

    # Check existing ISO
    if (Test-Path $isoPath) {
        $fileSize = (Get-Item $isoPath).Length / 1GB
        Write-Host "Found existing ISO: $isoName ($('{0:N2}' -f $fileSize) GB)" -ForegroundColor Yellow
        
        $response = Read-Host "Use existing ISO? (Y/N)"
        if ($response -match '^[Yy]') {
            return $isoPath
        }
        Remove-Item $isoPath -Force
    }

    # Download Fido
    Write-Host "Downloading Fido.ps1..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' `
            -OutFile $fidoScript -UseBasicParsing -ErrorAction Stop
        Write-Host "✓ Fido.ps1 downloaded" -ForegroundColor Green
    }
    catch {
        throw "Failed to download Fido.ps1: $_"
    }

    # Download ISO
    Write-Host "`nDownloading Windows $Win $Rel ISO..." -ForegroundColor Cyan
    Write-Host "This may take 15-30 minutes depending on connection speed..." -ForegroundColor Yellow
    
    try {
        $fidoParams = @{
            Win = $Win
            Rel = $Rel
            Ed = $Ed
            Lang = $Lang
            Arch = $Arch
            GetUrl = $false
        }
        
        & $fidoScript @fidoParams -OutFile $isoPath -ErrorAction Stop

        if ($LASTEXITCODE -ne 0 -and $LASTEXITCODE -ne $null) {
            throw "Fido exited with code $LASTEXITCODE"
        }
    }
    catch {
        throw "Fido execution failed: $_"
    }

    if (-not (Test-Path $isoPath)) {
        throw "ISO download failed; file not found at $isoPath"
    }
    
    $fileSize = (Get-Item $isoPath).Length / 1GB
    Write-Host "✓ ISO downloaded: $isoName ($('{0:N2}' -f $fileSize) GB)" -ForegroundColor Green
    return $isoPath
}

function Repair-Windows {
    param([string]$MountPath)

    $dismExe = Join-Path $env:SystemRoot 'System32\Dism.exe'
    $sourcesPath = Join-Path $MountPath 'sources'
    
    if (-not (Test-Path $sourcesPath)) {
        throw "Sources directory not found at $sourcesPath"
    }

    # Find source file
    $wim = Join-Path $sourcesPath 'install.wim'
    $esd = Join-Path $sourcesPath 'install.esd'

    if (Test-Path $wim) {
        $sourceArg = $wim
        Write-Host "Using install.wim as repair source" -ForegroundColor Green
    }
    elseif (Test-Path $esd) {
        $sourceArg = "esd:$esd:1"
        Write-Host "Using install.esd as repair source" -ForegroundColor Green
    }
    else {
        throw "No install.wim or install.esd found in $sourcesPath"
    }

    # Run DISM repair
    Write-Host "`nRunning DISM /RestoreHealth..." -ForegroundColor Cyan
    Write-Host "This may take 15-30 minutes. Please wait..." -ForegroundColor Yellow
    
    try {
        $arguments = @(
            '/Online'
            '/Cleanup-Image'
            '/RestoreHealth'
            "/Source:$sourceArg"
            '/LimitAccess'
        )
        
        $process = Start-Process -FilePath $dismExe -ArgumentList $arguments `
            -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -ne 0) {
            throw "DISM failed with exit code: $($process.ExitCode)"
        }
        
        Write-Host "✓ DISM repair completed successfully" -ForegroundColor Green
    }
    catch {
        throw "DISM repair failed: $_"
    }
}


# --- Main Execution ---
Write-Host "Detecting system information..." -ForegroundColor Cyan
$sysInfo = Get-SystemInfo

Write-Host "`nSystem Information:" -ForegroundColor White
Write-Host "  OS: $($sysInfo.Caption)"
Write-Host "  Version: $($sysInfo.Version)"
Write-Host "  Build: $($sysInfo.BuildNumber)"
Write-Host "  Architecture: $($sysInfo.OSArchitecture)"

# Validate Windows 11
if ($sysInfo.Caption -notmatch 'Windows\s+11') {
    throw "This script only supports Windows 11. Detected: $($sysInfo.Caption)"
}

# Auto-detect release version based on build number
if (-not $Release) {
    $buildMap = @{
        26200 = '25H2'  # Windows 11 25H2
        26100 = '24H2'  # Windows 11 24H2
        22631 = '23H2'  # Windows 11 23H2
        22621 = '22H2'  # Windows 11 22H2
        22000 = '21H2'  # Windows 11 21H2
    }
    
    # Match closest build
    $detectedRelease = $null
    foreach ($build in $buildMap.Keys | Sort-Object -Descending) {
        if ($sysInfo.BuildNumber -ge $build) {
            $detectedRelease = $buildMap[$build]
            break
        }
    }
    
    if ($detectedRelease) {
        $Release = $detectedRelease
        Write-Host "  Target Release: $Release (auto-detected)" -ForegroundColor Green
    }
    else {
        # Default to latest if unknown
        $Release = '25H2'
        Write-Host "  Target Release: $Release (defaulting to latest)" -ForegroundColor Yellow
    }
}
else {
    Write-Host "  Target Release: $Release (user-specified)" -ForegroundColor Green
}

# Auto-detect architecture
if (-not $Arch) {
    $Arch = if ($sysInfo.OSArchitecture -match 'ARM64') { 'arm64' } else { 'x64' }
    Write-Host "  Target Architecture: $Arch (auto-detected)" -ForegroundColor Green
}

# Check disk space
Write-Host "`nChecking disk space..." -ForegroundColor Cyan
$diskInfo = Get-DiskSpace -Drive 'C'
if ($diskInfo) {
    Write-Host "  Drive C: - Free: $($diskInfo.FreeGB) GB / Total: $($diskInfo.TotalGB) GB"
    
    if ($diskInfo.FreeGB -lt 8) {
        Write-Warning "Low disk space! At least 8GB free recommended."
        $continue = Read-Host "Continue anyway? (Y/N)"
        if ($continue -notmatch '^[Yy]') {
            Write-Host "Operation cancelled by user." -ForegroundColor Yellow
            exit 0
        }
    }
}

# Build mapping for version comparison
$buildTargets = @{
    '25H2' = 26200
    '24H2' = 26100
    '23H2' = 22631
    '22H2' = 22621
    '21H2' = 22000
}

$targetBuild = $buildTargets[$Release]
if ($targetBuild) {
    Write-Host "`nCurrent build: $($sysInfo.BuildNumber) | Target build: $targetBuild"
    
    if ($sysInfo.BuildNumber -ge $targetBuild) {
        Write-Host "Note: Your system is already at or above the target version." -ForegroundColor Yellow
    }
}

# Download ISO
Write-Host "`nPreparing to download Windows 11 $Release ISO..." -ForegroundColor Cyan
try {
    $isoPath = Download-WindowsISO -DestinationDirectory $DestinationDirectory `
        -Win $TargetWin -Rel $Release -Ed $Edition -Arch $Arch -Lang $Language
}
catch {
    Write-Error "Failed to download ISO: $_"
    exit 1
}

# Mount ISO
Write-Host "`nMounting ISO: $isoPath" -ForegroundColor Cyan
try {
    $mount = Mount-DiskImage -ImagePath $isoPath -PassThru -ErrorAction Stop
    $vol = Get-Volume -DiskImage $mount -ErrorAction Stop
    $drive = $vol.DriveLetter + ':\'
    Write-Host "✓ ISO mounted at: $drive" -ForegroundColor Green
}
catch {
    Write-Error "Failed to mount ISO: $_"
    exit 1
}

# Perform DISM repair
try {
    Repair-Windows -MountPath $drive
}
catch {
    Write-Error "DISM repair failed: $_"
    Dismount-DiskImage -ImagePath $isoPath -ErrorAction SilentlyContinue
    exit 1
}
finally {
    # Cleanup
    Write-Host "`nUnmounting ISO..." -ForegroundColor Cyan
    try {
        Dismount-DiskImage -ImagePath $isoPath -ErrorAction Stop
        Write-Host "✓ ISO unmounted" -ForegroundColor Green
    }
    catch {
        Write-Warning "Failed to unmount ISO: $_"
    }
}

Write-Host "`n=== Script Completed Successfully ===`n" -ForegroundColor Green
Write-Host "System repair using Windows 11 $Release has been completed." -ForegroundColor White
Write-Host "Recommended: Restart your computer to apply all changes.`n" -ForegroundColor Yellow

}

Win11-DISM-Repair -Release "25H2" -Arch "x64" -Edition "Pro"

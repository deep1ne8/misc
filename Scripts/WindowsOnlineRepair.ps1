<#
.SYNOPSIS
    Download a Windows 10 or 11 ISO via Fido, mount it, run DISM repair, then unmount.

.DESCRIPTION
    - Auto-detects or lets you specify target OS (10 or 11), release (e.g. 22H2), edition, arch, lang.
    - Downloads Fido.ps1 and invokes it to grab the correct ISO.
    - Mounts the ISO, runs DISM /RestoreHealth against install.wim/.esd, then unmounts.

.PARAMETER DestinationDirectory
    Where to store Fido.ps1 and the downloaded ISO. Default: C:\WindowsSetup

.PARAMETER TargetWin
    “10” or “11”. Auto-detected if omitted.

.PARAMETER Release
    Release build (e.g. 22H2). Default: 22H2

.PARAMETER Edition
    Edition (Pro, Home, etc.). Default: Pro

.PARAMETER Arch
    Architecture (x64, x86, arm64). Default: x64

.PARAMETER Language
    Language code (Eng). Default: Eng
#>
[CmdletBinding()]
param(
    [string] $DestinationDirectory = "C:\WindowsSetup",
    [ValidateSet("10","11")] [string] $TargetWin,
    [string] $Release  = "22H2",
    [string] $Edition  = "Pro",
    [ValidateSet("x64","x86","arm64")] [string] $Arch = "x64",
    [string] $Language = "Eng"
)

# ── 1) Elevation check
if (-not ([Security.Principal.WindowsPrincipal] `
           [Security.Principal.WindowsIdentity]::GetCurrent() `
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Please re-run this script **as Administrator**."
    exit 1
}

# ── 2) Helper functions ─────────────────────────────────────────────────────

function Get-DiskSpace {
    param([string] $Drive = 'C')
    Get-PSDrive -Name $Drive -ErrorAction Stop |
      Select-Object Root,
        @{Name='Used(GB)';Expression={[math]::Round($_.Used/1GB,2)}},
        @{Name='Free(GB)';Expression={[math]::Round($_.Free/1GB,2)}}
}

function Download-WindowsISO {
    param(
      [Parameter(Mandatory)][string] $DestinationDirectory,
      [Parameter(Mandatory)][ValidateSet("10","11")][string] $Win,
      [Parameter(Mandatory)][string] $Rel,
      [Parameter(Mandatory)][string] $Ed,
      [Parameter(Mandatory)][ValidateSet("x64","x86","arm64")][string] $Arch,
      [Parameter(Mandatory)][string] $Lang
    )

    # ensure folder
    if (-not (Test-Path $DestinationDirectory)) {
        New-Item $DestinationDirectory -ItemType Directory -Force | Out-Null
    }

    # fetch Fido
    $fidoScript = Join-Path $DestinationDirectory 'Fido.ps1'
    Write-Host "🔽 Downloading Fido.ps1 to $DestinationDirectory..."
    Invoke-WebRequest `
      -Uri 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' `
      -OutFile $fidoScript -UseBasicParsing -ErrorAction Stop

    # build ISO name/path
    $isoName = "Win${Win}_${Rel}_${Lang}_${Arch}.iso"
    $isoPath = Join-Path $DestinationDirectory $isoName

    # run Fido
    Write-Host "🪟 Running Fido to get Windows $Win $Rel ISO..."
    & $fidoScript `
      -Win  $Win `
      -Rel  $Rel `
      -Ed   $Ed  `
      -Lang $Lang `
      -Arch $Arch `
      -OutFile $isoPath `
      -ErrorAction Stop

    if (-not (Test-Path $isoPath)) {
        throw "ISO download failed; file not found at $isoPath"
    }
    Write-Host "✅ ISO is at $isoPath"
    return $isoPath
}

function Repair-Windows {
    param([Parameter(Mandatory)][string] $MountPath)  # e.g. "E:\"

    $dismExe = Join-Path $env:SystemRoot 'System32\Dism.exe'
    $wim     = Join-Path $MountPath    'sources\install.wim'
    $esd     = Join-Path $MountPath    'sources\install.esd'

    if (Test-Path $wim) {
        $sourceArg = $wim
    }
    elseif (Test-Path $esd) {
        $sourceArg = "esd:$esd:1"
    }
    else {
        throw "No install.wim or install.esd under $MountPath\sources"
    }

    Write-Host "🔧 Running DISM /RestoreHealth with source $sourceArg..."
    & $dismExe /Online /Cleanup-Image /RestoreHealth "/Source:$sourceArg" /LimitAccess `
      -ErrorAction Stop

    Write-Host "✅ DISM completed successfully."
}

# ── 3) Auto-detect TargetWin if omitted ────────────────────────────────────
if (-not $TargetWin) {
    $caption = (Get-CimInstance Win32_OperatingSystem).Caption
    if    ($caption -match 'Windows\s+11') { $TargetWin = '11' }
    elseif ($caption -match 'Windows\s+10') { $TargetWin = '10' }
    else {
        Write-Warning "Could not detect OS from '$caption'; defaulting to Win10."
        $TargetWin = '10'
    }
    Write-Host "ℹ️  Target OS set to Windows $TargetWin"
}

# ── 4) Show C: disk space ─────────────────────────────────────────────────
Write-Host "ℹ️  Disk space on C: drive:"
Get-DiskSpace -Drive 'C' | Format-Table -AutoSize

# ── 5) Ensure module loaded ───────────────────────────────────────────────
if (-not (Get-Module Microsoft.PowerShell.Management)) {
    Import-Module Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
}

# ── 6) Decide on ISO ─────────────────────────────────────────────────────
$os           = Get-CimInstance Win32_OperatingSystem
$currentBuild = [int]$os.BuildNumber
$targets      = @{ '10' = 19045; '11' = 22621 }
$targetBuild  = $targets[$TargetWin]

if ($currentBuild -lt $targetBuild) {
    Write-Host "⚙️  Current build $currentBuild < $targetBuild. Downloading ISO..."
    $IsoPath = Download-WindowsISO `
      -DestinationDirectory $DestinationDirectory `
      -Win  $TargetWin -Rel $Release -Ed $Edition -Arch $Arch -Lang $Language
}
else {
    $isoName = "Win${TargetWin}_${Release}_${Language}_${Arch}.iso"
    $IsoPath  = Join-Path $DestinationDirectory $isoName

    if (-not (Test-Path $IsoPath)) {
        Write-Warning "ISO not found at $IsoPath. Downloading automatically..."
        $IsoPath = Download-WindowsISO `
          -DestinationDirectory $DestinationDirectory `
          -Win  $TargetWin -Rel $Release -Ed $Edition -Arch $Arch -Lang $Language
    }
    else {
        Write-Host "ℹ️  Build $currentBuild >= $targetBuild; using existing ISO at $IsoPath"
    }
}

# ── 7) Mount the ISO ─────────────────────────────────────────────────────
Write-Host "📀 Mounting ISO..."
try {
    $mount = Mount-DiskImage -ImagePath $IsoPath -PassThru -ErrorAction Stop
    $vol   = Get-Volume -DiskImage $mount -ErrorAction Stop
    $drive = $vol.DriveLetter + ':\'
    Write-Host "✅ ISO mounted at $drive"
}
catch {
    Write-Error "Failed to mount ISO: $_"
    exit 1
}

# ── 8) Repair via DISM ───────────────────────────────────────────────────
Repair-Windows -MountPath $drive

# ── 9) Clean up ───────────────────────────────────────────────────────────
Write-Host "📤 Unmounting ISO..."
Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue

# ── Done!
Write-Host "All done!"

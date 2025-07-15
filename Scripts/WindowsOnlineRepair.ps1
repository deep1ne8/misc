<#
.SYNOPSIS
    Download a Windows 10 or 11 ISO via Fido, mount it, run DISM repair, then unmount.

.PARAMETER DestinationDirectory
    Path to store Fido.ps1 and the downloaded ISO. Defaults to C:\WindowsSetup.

.PARAMETER TargetWin
    “10” or “11”. If omitted, auto-detected from the running OS.

.PARAMETER Release
    Release build, e.g. 22H2. Default is 22H2.

.PARAMETER Edition
    Edition (Pro, Home, etc.). Default is Pro.

.PARAMETER Arch
    Architecture: x64, x86, or arm64. Default is x64.

.PARAMETER Language
    Language code (Eng). Default is Eng.
#>
[CmdletBinding()]
param(
    [string] $DestinationDirectory = "C:\WindowsSetup",
    [ValidateSet("10","11")]
    [string] $TargetWin,
    [string] $Release      = "22H2",
    [string] $Edition      = "Pro",
    [ValidateSet("x64","x86","arm64")]
    [string] $Arch         = "x64",
    [string] $Language     = "Eng"
)

# ── 1) Elevation check
if (-not ([Security.Principal.WindowsPrincipal] `
          [Security.Principal.WindowsIdentity]::GetCurrent() `
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Please re-run this script as Administrator."
    exit 1
}

# ── 2) Helpers

function Get-DiskSpace {
    param(
        [string] $Drive = 'C'
    )
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

    # Ensure folder exists
    if (-not (Test-Path $DestinationDirectory)) {
        New-Item -Path $DestinationDirectory -ItemType Directory -Force | Out-Null
    }

    # Download Fido.ps1
    $fidoScript = Join-Path $DestinationDirectory 'Fido.ps1'
    Write-Host "Downloading Fido.ps1..."
    Invoke-WebRequest `
        -Uri 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' `
        -OutFile $fidoScript -UseBasicParsing -ErrorAction Stop

    # Build ISO name & path
    $isoName = "Win${Win}_${Rel}_${Lang}_${Arch}.iso"
    $isoPath = Join-Path $DestinationDirectory $isoName

    # Run Fido to fetch ISO
    Write-Host "Running Fido to download Windows $Win $Rel..."
    & $fidoScript `
        -Win  $Win `
        -Rel  $Rel `
        -Ed   $Ed  `
        -Lang $Lang `
        -Arch $Arch `
        -OutFile $isoPath `
        -ErrorAction Stop

    if (-not (Test-Path $isoPath)) {
        throw "ISO download failed; not found at $isoPath"
    }

    Write-Host "ISO downloaded to $isoPath"
    return $isoPath
}

function Repair-Windows {
    param(
        [Parameter(Mandatory)][string] $MountPath
    )

    $dismExe = Join-Path $env:SystemRoot 'System32\Dism.exe'
    $wim     = Join-Path $MountPath 'sources\install.wim'
    $esd     = Join-Path $MountPath 'sources\install.esd'

    if (Test-Path $wim) {
        $sourceArg = $wim
    }
    elseif (Test-Path $esd) {
        $sourceArg = "esd:$esd:1"
    }
    else {
        throw "No install.wim or install.esd under $MountPath\sources"
    }

    Write-Host "Running DISM /RestoreHealth..."
    & $dismExe `
        /Online /Cleanup-Image /RestoreHealth `
        "/Source:$sourceArg" /LimitAccess `
        -ErrorAction Stop

    Write-Host "DISM completed successfully."
}

# ── 3) Auto-detect TargetWin if not supplied
if (-not $TargetWin) {
    $caption = (Get-CimInstance Win32_OperatingSystem).Caption
    if ($caption -match 'Windows\s+11')   { $TargetWin = '11' }
    elseif ($caption -match 'Windows\s+10') { $TargetWin = '10' }
    else {
        Write-Warning "Cannot detect OS from '$caption'; defaulting to 10."
        $TargetWin = '10'
    }
    Write-Host "Target OS: Windows $TargetWin"
}

# ── 4) Show disk space
Write-Host "Disk space on C: drive:"
Get-DiskSpace -Drive 'C' | Format-Table -AutoSize

# ── 5) Ensure module loaded
if (-not (Get-Module Microsoft.PowerShell.Management)) {
    Import-Module Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
}

# ── 6) Decide whether to download ISO
$os          = Get-CimInstance Win32_OperatingSystem
$currentBuild= [int]$os.BuildNumber
$targets     = @{ '10' = 19045; '11' = 22621 }
$targetBuild = $targets[$TargetWin]

if ($currentBuild -lt $targetBuild) {
    Write-Host "Current build $currentBuild is less than target $targetBuild. Downloading ISO..."
    $IsoPath = Download-WindowsISO `
        -DestinationDirectory $DestinationDirectory `
        -Win  $TargetWin `
        -Rel  $Release `
        -Ed   $Edition `
        -Arch $Arch `
        -Lang $Language
}
else {
    $isoName = "Win${TargetWin}_${Release}_${Language}_${Arch}.iso"
    $IsoPath = Join-Path $DestinationDirectory $isoName
    Write-Host "Build $currentBuild ≥ $targetBuild; looking for existing ISO at $IsoPath"
}

if (-not (Test-Path $IsoPath)) {
    throw "ISO not found at $IsoPath; cannot continue."
}

# ── 7) Mount the ISO
Write-Host "Mounting ISO..."
try {
    $mount = Mount-DiskImage -ImagePath $IsoPath -PassThru -ErrorAction Stop
    $vol   = Get-Volume -DiskImage $mount -ErrorAction Stop
    $drive = $vol.DriveLetter + ':\'
    Write-Host "ISO mounted at $drive"
}
catch {
    throw "Failed to mount ISO: $_"
}

# ── 8) Repair via DISM
Repair-Windows -MountPath $drive

# ── 9) Unmount
Write-Host "Dismounting ISO..."
Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue

# ── Done
Write-Host "All done!"

<#
.SYNOPSIS
    Download a Windows 10/11 ISO and repair your current OS with DISM.

.DESCRIPTION
    - Auto-detects or lets you specify target OS (10 or 11), release (e.g. 22H2), edition, arch, lang.  
    - Downloads Fido.ps1 and invokes it to grab the correct ISO.  
    - Mounts the ISO, runs DISM /RestoreHealth against install.wim/.esd, then unmounts.  

.PARAMETER DestinationDirectory
    Where to store Fido.ps1 and the downloaded ISO. Default: C:\WindowsSetup

.PARAMETER TargetWin
    Which Windows to download (10 or 11). Auto-detected if omitted.

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
    [Parameter(Position=0)]
    [string] $DestinationDirectory = "C:\WindowsSetup",

    [Parameter(Position=1)]
    [ValidateSet("10","11")]
    [string] $TargetWin,

    [Parameter(Position=2)]
    [string] $Release = "22H2",

    [Parameter(Position=3)]
    [string] $Edition = "Pro",

    [Parameter(Position=4)]
    [ValidateSet("x64","x86","arm64")]
    [string] $Arch = "x64",

    [Parameter(Position=5)]
    [string] $Language = "Eng"
)

# ── 1) Ensure we're running as Admin
if (-not ([Security.Principal.WindowsPrincipal] `
        [Security.Principal.WindowsIdentity]::GetCurrent() `
        ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "❗ Please re-run this script **as Administrator**."
    exit 1
}

# ── 2) Functions

function Get-DiskSpace {
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string] $Drive = 'C'
    )
    Get-PSDrive -Name $Drive -ErrorAction Stop |
        Select-Object Root,
            @{Name='Used(GB)';Expression={[math]::Round($_.Used/1GB,2)}},
            @{Name='Free(GB)';Expression={[math]::Round($_.Free/1GB,2)}}
}

function Download-WindowsISO {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $DestinationDirectory,
        [Parameter(Mandatory)][ValidateSet("10","11")] [string] $Win,
        [Parameter(Mandatory)][string] $Rel,
        [Parameter(Mandatory)][string] $Ed,
        [Parameter(Mandatory)][ValidateSet("x64","x86","arm64")] [string] $Arch,
        [Parameter(Mandatory)][string] $Lang
    )

    # Prepare folder
    if (-not (Test-Path $DestinationDirectory)) {
        Write-Verbose "Creating folder $DestinationDirectory"
        New-Item -Path $DestinationDirectory -ItemType Directory -Force | Out-Null
    }

    # Download Fido
    $fidoScript = Join-Path $DestinationDirectory 'Fido.ps1'
    Write-Host "🔽 Downloading Fido script..."
    try {
        Invoke-WebRequest `
            -Uri 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' `
            -OutFile $fidoScript -UseBasicParsing -ErrorAction Stop
    } catch {
        throw "Failed to download Fido.ps1: $_"
    }

    # Build ISO filename
    $isoName = "Win${Win}_${Rel}_${Lang}_${Arch}.iso"
    $isoPath = Join-Path $DestinationDirectory $isoName

    # Invoke Fido to download ISO
    Write-Host "🪟 Running Fido to get Windows $Win $Rel ISO..."
    & $fidoScript `
        -Win $Win `
        -Rel $Rel `
        -Ed  $Ed  `
        -Lang $Lang `
        -Arch $Arch `
        -OutFile $isoPath `
        -ErrorAction Stop

    if (-not (Test-Path $isoPath)) {
        throw "ISO download failed; file not found at $isoPath"
    }

    Write-Host "✅ ISO downloaded to $isoPath"
    return $isoPath
}

function Repair-Windows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $MountPath  # e.g. 'E:\'
    )

    $dismExe = Join-Path $env:SystemRoot 'System32\Dism.exe'
    $wim      = Join-Path $MountPath    'sources\install.wim'
    $esd      = Join-Path $MountPath    'sources\install.esd'

    if (Test-Path $wim) {
        $sourceArg = $wim
    } elseif (Test-Path $esd) {
        # ESD: use image index 1 by default
        $sourceArg = "esd:$esd:1"
    } else {
        throw "No install.wim or install.esd found under $MountPath\sources"
    }

    Write-Host "🔧 Repairing Windows (DISM source = $sourceArg)..."
    try {
        & $dismExe `
            /Online /Cleanup-Image /RestoreHealth `
            "/Source:$sourceArg" /LimitAccess `
            -ErrorAction Stop

        Write-Host "✅ DISM RestoreHealth completed successfully."
    } catch {
        throw "DISM repair failed: $_"
    }
}

# ── 3) Auto-detect TargetWin if not specified
if (-not $TargetWin) {
    $caption = (Get-CimInstance Win32_OperatingSystem).Caption
    if ($caption -match 'Windows\s+11') { $TargetWin = '11' }
    elseif ($caption -match 'Windows\s+10') { $TargetWin = '10' }
    else {
        Write-Warning "Could not auto-detect OS from '$caption'. Defaulting to Windows 10."
        $TargetWin = '10'
    }
    Write-Host "ℹ️  Target OS set to Windows $TargetWin"
}

# ── 4) Gather disk-space info
Write-Host "ℹ️  Disk space on C: drive:"
Get-DiskSpace -Drive 'C' | Format-Table -AutoSize

# ── 5) Ensure Microsoft.PowerShell.Management is loaded
if (-not (Get-Module Microsoft.PowerShell.Management)) {
    Import-Module Microsoft.PowerShell.Management -ErrorAction SilentlyContinue
}

# ── 6) Check current build vs. target build
$os        = Get-CimInstance Win32_OperatingSystem
$build     = [int]$os.BuildNumber
$targetMap = @{ '10' = 19045; '11' = 22621 }
$targetBuild = $targetMap[$TargetWin]

if ($build -lt $targetBuild) {
    Write-Host "⚙️  Current build $build < target $targetBuild. Downloading ISO..."
    $IsoPath = Download-WindowsISO `
        -DestinationDirectory $DestinationDirectory `
        -Win  $TargetWin `
        -Rel  $Release `
        -Ed   $Edition `
        -Arch $Arch `
        -Lang $Language
} else {
    # Assume ISO already in place
    $isoName = "Win${TargetWin}_${Release}_${Language}_${Arch}.iso"
    $IsoPath = Join-Path $DestinationDirectory $isoName
    Write-Host "ℹ️  Build $build ≥ $targetBuild. Looking for existing ISO: $IsoPath"
}

if (-not (Test-Path $IsoPath)) {
    throw "ISO not found at $IsoPath. Cannot proceed."
}

# ── 7) Mount the ISO
Write-Host "📀 Mounting ISO..."
try {
    $mountResult = Mount-DiskImage -ImagePath $IsoPath -PassThru -ErrorAction Stop
    # Grab the drive letter
    $vol = $mountResult | Get-Volume -ErrorAction Stop
    $driveLetter = $vol.DriveLetter + ':\'
    Write-Host "✅ ISO mounted at $driveLetter"
} catch {
    throw "Failed to mount ISO: $_"
}

# ── 8) Repair Windows
Repair-Windows -MountPath $driveLetter

# ── 9) Clean up: unmount
Write-Host "📤 Unmounting ISO..."
Dismount-DiskImage -ImagePath $IsoPath -ErrorAction SilentlyContinue
Write-Host "All done! 🎉"

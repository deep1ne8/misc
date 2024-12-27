# Function to retrieve disk space information for the specified path
function Get-DiskSpace {
    $diskSpace = Get-PSDrive -Name C | Select-Object Root,
        @{Name = "Used(GB)"; Expression = {[math]::Round(($_.Used / 1GB), 2)}},
        @{Name = "Free(GB)"; Expression = {[math]::Round(($_.Free / 1GB), 2)}}
    return $diskSpace
}

# Function to download Windows 10 ISO using Fido
function Download-Win10ISO {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $DestinationDirectory
    )

    # Ensure the destination directory exists
    if (-not (Test-Path -Path $DestinationDirectory)) {
        mkdir -Path $DestinationDirectory | Out-Null
    }

    # Download Fido
    Write-Host "Downloading Fido..."
    $FidoPath = Join-Path -Path $DestinationDirectory -ChildPath "Fido.ps1"
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1" -OutFile $FidoPath -Verbose

    # Validate Fido script existence
    if (-not (Test-Path -Path $FidoPath)) {
        Write-Error "Failed to download Fido.ps1. Script exiting..."
        exit 1
    }

    # Simulate Fido script execution (modify as needed based on Fido's actual usage)
    Write-Host "Running Fido to download Windows 10 ISO..."
    $IsoPath = Join-Path -Path $DestinationDirectory -ChildPath "Win10_22H2_English_x64v1.iso"
    & $FidoPath -Win "10" -Rel "22H2" -Ed "Pro" -Lang "Eng" -Arch "x64" -OutFile $IsoPath

    if (-not (Test-Path -Path $IsoPath)) {
        Write-Error "Failed to download the Windows 10 ISO. Script exiting..."
        exit 1
    }
    return $IsoPath
}

# Function to repair Windows using DISM
function Repair-Windows {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $MountPath
    )

    $dismPath = Join-Path $env:SystemRoot -ChildPath "System32\Dism.exe" -Resolve
    $sourcePath = Join-Path $MountPath -ChildPath "sources\install.wim" -Resolve

    if (Test-Path -Path $sourcePath) {
        Write-Verbose "Repairing Windows using DISM..."
        & $dismPath /Online /Cleanup-Image /RestoreHealth /Source:$sourcePath /LimitAccess
        Write-Verbose "Repair completed."
    } else {
        Write-Error "Source path not found. Repair aborted."
        exit 1
    }
}

# Main script
$DestinationDirectory = "C:\WindowsSetup"

# Get disk space information
Write-Verbose "Getting disk space on C: drive"
Get-DiskSpace

# Ensure required module
if (-not (Get-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue)) {
    Import-Module Microsoft.PowerShell.Management
}

# Check Windows version and download ISO if necessary
$windowsVersion = (Get-ComputerInfo).OsVersion
if ($windowsVersion -ne '10.0.19045') {
    Write-Host "Current Windows OS version: $windowsVersion"
    Write-Host "Downloading Windows 10 ISO..."
    $IsoPath = Download-Win10ISO -DestinationDirectory $DestinationDirectory
} else {
    $IsoPath = Join-Path -Path $DestinationDirectory -ChildPath "Win10_22H2_English_x64v1.iso"
}

# Verify ISO existence
if (-not (Test-Path -Path $IsoPath)) {
    Write-Error "ISO file not found. Script exiting..."
    exit 1
}

# Mount the ISO
$mountJob = Mount-DiskImage -ImagePath $IsoPath -PassThru -Verbose
$mountedDrive = (Get-Volume | Where-Object { $_.DriveType -eq "CD-ROM" }).DriveLetter

if (-not $mountedDrive) {
    Write-Error "Failed to mount the ISO file. Script exiting..."
    Dismount-DiskImage -ImagePath $mountJob.ImagePath
    exit 1
}

# Repair Windows
Write-Verbose "Repairing Windows using DISM..."
Repair-Windows -MountPath $mountedDrive

# Unmount the ISO
Dismount-DiskImage -ImagePath $mountJob.ImagePath
Write-Host "ISO unmounted successfully."

# Function to retrieve disk space information for the specified path
function Get-DiskSpace {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true, Position=0)]
        [ValidateScript({Test-Path $_})]
        [string]
        $Path
    )

    $diskSpace = Get-PSDrive -Name C | Select-Object Root,
        @{Name="Used(GB)"; Expression={[math]::Round(($_.Used / 1GB), 2)}},
        @{Name="Free(GB)"; Expression={[math]::Round(($_.Free / 1GB), 2)}}
    return $diskSpace
}

# Function to download Windows 10 ISO using Fido
function Download-Win10ISO {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $DestinationDirectory
    )

    $FidoCommand = Join-Path -Path $DestinationDirectory -ChildPath "Fido.ps1" -Resolve
    $IsoPath = Join-Path -Path $DestinationDirectory -ChildPath "Win10_22H2_English_x64v1.iso" -Resolve
    if (-not (Test-Path $IsoPath)) {
        Write-Verbose -Message "Windows 10 ISO not found. Downloading..."
    }
    $userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    &$FidoCommand -Win "10" -Rel "22H2" -Ed "Pro" -Lang "Eng" -Arch "x64" -UserAgent $userAgent
}

# Function to repair Windows using DISM
function Repair-Windows {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]
        $MountPath
    )

    $dismPath = Join-Path $env:SystemRoot -ChildPath "System32\Dism.exe" -Resolve
    $sourcePath = Join-Path $MountPath -ChildPath "sources\install.wim" -Resolve

    if (Test-Path $sourcePath) {
        Write-Verbose -Message "Repairing Windows using DISM..."
        & $dismPath /Online /Cleanup-Image /RestoreHealth /Source:$sourcePath /LimitAccess
        Write-Verbose -Message "Repair completed."
    } else {
        Write-Error -Message "Source path not found. Repair aborted."
    }
}

# Main script

# Load required modules
Get-Module -Name Microsoft.PowerShell.Management -ErrorAction SilentlyContinue | Import-Module

# Set strict mode
Set-StrictMode -Version Latest

# Define paths
$DestinationDirectory = "C:\WindowsSetup"
$IsoPath = Join-Path -Path $DestinationDirectory -ChildPath "Win10_22H2_English_x64v1.iso" -Resolve

# Check if the operating system is Windows 10 version 22H2 (2009)
$windowsVersion = (Get-ComputerInfo).OsVersion

if ($windowsVersion -ne '10.0.19045') {
    Write-Error -Message "This script is intended for Windows 10 version 22H2 (19045). Exiting..."
    exit
}

# Check if destination directory exists, if not create it
if (-not (Test-Path -Path $DestinationDirectory)) {
    New-Item -ItemType Directory -Path $DestinationDirectory | Out-Null
}

# Copy required files to the destination directory
Write-Verbose -Message "Copying required files to the destination directory"
Copy-Item -Path $EditorConfigFile -Destination $DestinationDirectory -Verbose -Force
Copy-Item -Path $Fido -Destination $DestinationDirectory -Verbose -Force

# Get disk space information
Write-Verbose -Message "Getting disk space on C: drive"
$DiskSpaceInfo = Get-DiskSpace -Path "C:\"
$DiskSpaceInfo | Format-List

Download-Win10ISO -DestinationDirectory $DestinationDirectory

# Mount the ISO file
$mountJob = Mount-DiskImage -ImagePath $IsoPath -PassThru -Verbose
$mountedDrive = Get-Volume | Where-Object { $_.DriveType -eq "CD-ROM" } | Select-Object -ExpandProperty DriveLetter

# Check if the ISO is mounted successfully
if (-not $mountedDrive) {
    Write-Error -Message "Failed to mount the ISO file. Repair process aborted."
    return
}

# Repair Windows using DISM
Write-Verbose -Message "Continuing with repair on Windows Filesystem Components"
Repair-Windows -MountPath $mountedDrive

#Unmount ISO file
Dismount-DiskImage -ImagePath $mountJob.ImagePath
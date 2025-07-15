<#
.SYNOPSIS
    Download a Windows 10 or 11 ISO via Fido, mount it, run DISM repair, then unmount.

.DESCRIPTION
    - Auto-detects current Windows version and downloads matching ISO
    - Downloads Fido.ps1 and invokes it to grab the correct ISO
    - Mounts the ISO, runs DISM /RestoreHealth against install.wim/.esd, then unmounts

.PARAMETER DestinationDirectory
    Where to store Fido.ps1 and the downloaded ISO. Default: C:\WindowsSetup

.PARAMETER TargetWin
    "10" or "11". Auto-detected if omitted based on current OS

.PARAMETER Release
    Release build. Auto-detected if omitted (24H2 for Win11, 22H2 for Win10)

.PARAMETER Edition
    Edition (Pro, Home, etc.). Default: Pro

.PARAMETER Arch
    Architecture (x64, x86, arm64). Auto-detected if omitted

.PARAMETER Language
    Language code. Default: Eng
#>
[CmdletBinding()]
param(
    [string] $DestinationDirectory = "C:\WindowsSetup",
    [ValidateSet("10","11")] [string] $TargetWin,
    [string] $Release,
    [string] $Edition = "Pro",
    [ValidateSet("x64","x86","arm64")] [string] $Arch,
    [string] $Language = "Eng"
)

# Check for administrator privileges
if (-not ([Security.Principal.WindowsPrincipal] `
           [Security.Principal.WindowsIdentity]::GetCurrent() `
         ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script requires administrator privileges. Please run as Administrator."
    exit 1
}

# Helper function to get system information
function Get-SystemInfo {
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        $cs = Get-CimInstance Win32_ComputerSystem
        
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

# Helper function to get disk space information
function Get-DiskSpace {
    param([string] $Drive = 'C')
    try {
        $driveInfo = Get-PSDrive -Name $Drive -ErrorAction Stop
        return @{
            Root = $driveInfo.Root
            UsedGB = [math]::Round($driveInfo.Used/1GB, 2)
            FreeGB = [math]::Round($driveInfo.Free/1GB, 2)
            TotalGB = [math]::Round(($driveInfo.Used + $driveInfo.Free)/1GB, 2)
        }
    }
    catch {
        Write-Warning "Could not retrieve disk space information for drive $Drive"
        return $null
    }
}

# Function to download Windows ISO using Fido
function Download-WindowsISO {
    param(
        [Parameter(Mandatory)][string] $DestinationDirectory,
        [Parameter(Mandatory)][ValidateSet("10","11")][string] $Win,
        [Parameter(Mandatory)][string] $Rel,
        [Parameter(Mandatory)][string] $Ed,
        [Parameter(Mandatory)][ValidateSet("x64","x86","arm64")][string] $Arch,
        [Parameter(Mandatory)][string] $Lang
    )

    # Ensure destination directory exists
    if (-not (Test-Path $DestinationDirectory)) {
        try {
            New-Item $DestinationDirectory -ItemType Directory -Force | Out-Null
            Write-Host "Created directory: $DestinationDirectory"
        }
        catch {
            throw "Failed to create directory $DestinationDirectory: $_"
        }
    }

    # Download Fido script
    $fidoScript = Join-Path $DestinationDirectory 'Fido.ps1'
    Write-Host "Downloading Fido.ps1..."
    
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest `
            -Uri 'https://raw.githubusercontent.com/pbatard/Fido/master/Fido.ps1' `
            -OutFile $fidoScript -UseBasicParsing -ErrorAction Stop
        Write-Host "Successfully downloaded Fido.ps1"
    }
    catch {
        throw "Failed to download Fido.ps1: $_"
    }

    # Build ISO name and path
    $isoName = "Win${Win}_${Rel}_${Lang}_${Arch}.iso"
    $isoPath = Join-Path $DestinationDirectory $isoName

    # Check if ISO already exists
    if (Test-Path $isoPath) {
        $fileSize = (Get-Item $isoPath).Length / 1GB
        Write-Host "Found existing ISO: $isoName ($('{0:N2}' -f $fileSize) GB)"
        
        $response = Read-Host "Do you want to use the existing ISO? (Y/N)"
        if ($response -match '^[Yy]') {
            return $isoPath
        }
        else {
            Remove-Item $isoPath -Force
        }
    }

    # Execute Fido to download ISO
    Write-Host "Downloading Windows $Win $Rel ISO (this may take a while)..."
    try {
        & $fidoScript `
            -Win $Win `
            -Rel $Rel `
            -Ed $Ed `
            -Lang $Lang `
            -Arch $Arch `
            -OutFile $isoPath `
            -ErrorAction Stop

        if ($LASTEXITCODE -ne 0) {
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
    Write-Host "ISO successfully downloaded: $isoName ($('{0:N2}' -f $fileSize) GB)"
    return $isoPath
}

# Function to perform DISM repair using mounted ISO
function Repair-Windows {
    param([Parameter(Mandatory)][string] $MountPath)

    $dismExe = Join-Path $env:SystemRoot 'System32\Dism.exe'
    $sourcesPath = Join-Path $MountPath 'sources'
    
    if (-not (Test-Path $sourcesPath)) {
        throw "Sources directory not found at $sourcesPath"
    }

    $wim = Join-Path $sourcesPath 'install.wim'
    $esd = Join-Path $sourcesPath 'install.esd'

    # Determine source file
    if (Test-Path $wim) {
        $sourceArg = $wim
        Write-Host "Using install.wim as repair source"
    }
    elseif (Test-Path $esd) {
        $sourceArg = "esd:$esd:1"
        Write-Host "Using install.esd as repair source"
    }
    else {
        throw "No install.wim or install.esd found in $sourcesPath"
    }

    # Run DISM repair
    Write-Host "Running DISM /RestoreHealth with source: $sourceArg"
    Write-Host "This operation may take 15-30 minutes depending on system condition..."
    
    try {
        $process = Start-Process -FilePath $dismExe -ArgumentList @(
            '/Online',
            '/Cleanup-Image',
            '/RestoreHealth',
            "/Source:$sourceArg",
            '/LimitAccess'
        ) -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -ne 0) {
            throw "DISM command failed with exit code: $($process.ExitCode)"
        }
        
        Write-Host "DISM repair completed successfully"
    }
    catch {
        throw "DISM repair failed: $_"
    }
}

# Main execution starts here
Write-Host "Windows ISO Download and DISM Repair Script"
Write-Host "===========================================" 

# Get system information
Write-Host "Detecting system information..."
$sysInfo = Get-SystemInfo

Write-Host "System Information:"
Write-Host "  OS: $($sysInfo.Caption)"
Write-Host "  Version: $($sysInfo.Version)"
Write-Host "  Build: $($sysInfo.BuildNumber)"
Write-Host "  Architecture: $($sysInfo.OSArchitecture)"

# Auto-detect target Windows version
if (-not $TargetWin) {
    if ($sysInfo.Caption -match 'Windows\s+11') {
        $TargetWin = '11'
    }
    elseif ($sysInfo.Caption -match 'Windows\s+10') {
        $TargetWin = '10'
    }
    else {
        throw "Unsupported Windows version detected: $($sysInfo.Caption)"
    }
    Write-Host "Target OS: Windows $TargetWin (auto-detected)"
}

# Auto-detect release version
if (-not $Release) {
    switch ($TargetWin) {
        '11' { $Release = '24H2' }
        '10' { $Release = '22H2' }
    }
    Write-Host "Target Release: $Release (auto-detected)"
}

# Auto-detect architecture
if (-not $Arch) {
    switch -Regex ($sysInfo.OSArchitecture) {
        'x64|AMD64' { $Arch = 'x64' }
        'x86|32-bit' { $Arch = 'x86' }
        'ARM64' { $Arch = 'arm64' }
        default { $Arch = 'x64' }
    }
    Write-Host "Target Architecture: $Arch (auto-detected)"
}

# Display disk space information
Write-Host "Checking disk space..."
$diskInfo = Get-DiskSpace -Drive 'C'
if ($diskInfo) {
    Write-Host "  Drive C: - Total: $($diskInfo.TotalGB) GB, Used: $($diskInfo.UsedGB) GB, Free: $($diskInfo.FreeGB) GB"
    
    if ($diskInfo.FreeGB -lt 8) {
        Write-Warning "Low disk space detected. At least 8GB free space is recommended for ISO download."
    }
}

# Build information mapping
$buildTargets = @{
    '10' = @{
        '22H2' = 19045
        '21H2' = 19044
    }
    '11' = @{
        '24H2' = 26100
        '23H2' = 22631
        '22H2' = 22621
    }
}

$targetBuild = $buildTargets[$TargetWin][$Release]
if (-not $targetBuild) {
    Write-Warning "Unknown build target for Windows $TargetWin $Release, proceeding with download..."
    $shouldDownload = $true
}
else {
    Write-Host "Current build: $($sysInfo.BuildNumber), Target build: $targetBuild"
    $shouldDownload = $sysInfo.BuildNumber -lt $targetBuild
}

# Download ISO
try {
    $isoPath = Download-WindowsISO `
        -DestinationDirectory $DestinationDirectory `
        -Win $TargetWin `
        -Rel $Release `
        -Ed $Edition `
        -Arch $Arch `
        -Lang $Language
}
catch {
    Write-Error "Failed to download ISO: $_"
    exit 1
}

# Mount the ISO
Write-Host "Mounting ISO: $isoPath"
try {
    $mount = Mount-DiskImage -ImagePath $isoPath -PassThru -ErrorAction Stop
    $vol = Get-Volume -DiskImage $mount -ErrorAction Stop
    $drive = $vol.DriveLetter + ':\'
    Write-Host "ISO mounted successfully at: $drive"
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
    # Ensure cleanup even if repair fails
    try {
        Dismount-DiskImage -ImagePath $isoPath -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Failed to unmount ISO during cleanup"
    }
    exit 1
}

# Clean up - unmount the ISO
Write-Host "Unmounting ISO..."
try {
    Dismount-DiskImage -ImagePath $isoPath -ErrorAction SilentlyContinue
    Write-Host "ISO unmounted successfully"
}
catch {
    Write-Warning "Failed to unmount ISO: $_"
}

Write-Host "Script completed successfully!"
Write-Host "System repair using Windows $TargetWin $Release has been completed."

<#
.SYNOPSIS
    Collects Windows and Liongard Agent identity information.

.DESCRIPTION
    Collects:
    - Computer name
    - Domain/workgroup
    - Windows MachineGuid
    - Windows Product ID
    - Computer System UUID
    - BIOS serial number
    - BIOS version
    - Manufacturer and model
    - Windows version/build
    - OS installation date
    - Liongard Agent MachineID
    - Liongard Agent version
    - Liongard Agent service status
    - Liongard Agent installation information
    - Network adapter information
    - IP addresses

    Output is displayed using Format-List.

.NOTES
    Read-only. This script does not modify the machine.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

function Get-RegistryValue {
    param (
        [string]$Path,
        [string]$Name
    )

    if (Test-Path $Path) {
        try {
            return (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name
        }
        catch {
            return $null
        }
    }

    return $null
}

# ------------------------------------------------------------
# System Information
# ------------------------------------------------------------

$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$ComputerSystemProduct = Get-CimInstance -ClassName Win32_ComputerSystemProduct
$BIOS = Get-CimInstance -ClassName Win32_BIOS
$OperatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem

# ------------------------------------------------------------
# Windows MachineGuid
# ------------------------------------------------------------

$MachineGuid = Get-RegistryValue `
    -Path 'HKLM:\SOFTWARE\Microsoft\Cryptography' `
    -Name 'MachineGuid'

# ------------------------------------------------------------
# Windows Product Information
# ------------------------------------------------------------

$WindowsProductName = Get-RegistryValue `
    -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
    -Name 'ProductName'

$WindowsDisplayVersion = Get-RegistryValue `
    -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
    -Name 'DisplayVersion'

$WindowsCurrentBuild = Get-RegistryValue `
    -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
    -Name 'CurrentBuild'

$WindowsUBR = Get-RegistryValue `
    -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
    -Name 'UBR'

$WindowsProductID = Get-RegistryValue `
    -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' `
    -Name 'ProductId'

# ------------------------------------------------------------
# Liongard Agent Registry Locations
# ------------------------------------------------------------

$LiongardRegistryPaths = @(
    'HKLM:\SOFTWARE\Liongard\LiongardAgent',
    'HKLM:\SOFTWARE\WOW6432Node\Liongard\LiongardAgent'
)

$LiongardMachineID = $null
$LiongardRegistryPath = $null
$LiongardRegistryProperties = $null

foreach ($Path in $LiongardRegistryPaths) {

    if (Test-Path $Path) {

        $Properties = Get-ItemProperty -Path $Path

        if (-not $LiongardMachineID) {
            $LiongardMachineID = $Properties.MachineID
        }

        if (-not $LiongardRegistryPath) {
            $LiongardRegistryPath = $Path
        }

        if (-not $LiongardRegistryProperties) {
            $LiongardRegistryProperties = $Properties
        }
    }
}

# ------------------------------------------------------------
# Liongard Agent Service
# ------------------------------------------------------------

$LiongardServices = Get-CimInstance Win32_Service |
    Where-Object {
        $_.Name -match 'Liongard|Roar' -or
        $_.DisplayName -match 'Liongard|Roar'
    }

# ------------------------------------------------------------
# Liongard Installed Software
# ------------------------------------------------------------

$LiongardSoftware = @()

$UninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

foreach ($Path in $UninstallPaths) {

    $LiongardSoftware += Get-ItemProperty $Path |
        Where-Object {
            $_.DisplayName -match 'Liongard'
        } |
        Select-Object DisplayName,
                      DisplayVersion,
                      Publisher,
                      InstallDate,
                      InstallLocation,
                      UninstallString
}

# ------------------------------------------------------------
# Liongard Agent Files
# ------------------------------------------------------------

$LiongardPossiblePaths = @(
    'C:\Program Files\Liongard',
    'C:\Program Files (x86)\Liongard',
    'C:\ProgramData\Liongard',
    'C:\Liongard'
)

$LiongardFiles = foreach ($Path in $LiongardPossiblePaths) {

    if (Test-Path $Path) {

        Get-ChildItem `
            -Path $Path `
            -Recurse `
            -File `
            -ErrorAction SilentlyContinue |
            Select-Object FullName,
                          Length,
                          LastWriteTime
    }
}

# ------------------------------------------------------------
# Network Information
# ------------------------------------------------------------

$NetworkAdapters = Get-CimInstance Win32_NetworkAdapterConfiguration |
    Where-Object {
        $_.IPEnabled -eq $true
    } |
    Select-Object Description,
                  MACAddress,
                  IPAddress,
                  IPSubnet,
                  DefaultIPGateway,
                  DNSServerSearchOrder

# ------------------------------------------------------------
# Build Output Object
# ------------------------------------------------------------

$Output = [PSCustomObject]@{

    'Computer Name'              = $env:COMPUTERNAME

    'FQDN'                       = try {
        [System.Net.Dns]::GetHostEntry($env:COMPUTERNAME).HostName
    }
    catch {
        $null
    }

    'Domain / Workgroup'         = $ComputerSystem.Domain

    'Domain Role'                = $ComputerSystem.DomainRole

    'Manufacturer'               = $ComputerSystem.Manufacturer

    'Model'                      = $ComputerSystem.Model

    'BIOS Serial Number'         = $BIOS.SerialNumber

    'BIOS Version'               = ($BIOS.SMBIOSBIOSVersion -join ', ')

    'System UUID'                = $ComputerSystemProduct.UUID

    'Windows MachineGuid'        = $MachineGuid

    'Windows Product ID'         = $WindowsProductID

    'Windows Product Name'       = $WindowsProductName

    'Windows Display Version'    = $WindowsDisplayVersion

    'Windows Build'              = "$WindowsCurrentBuild.$WindowsUBR"

    'Windows Version'            = $OperatingSystem.Version

    'OS Architecture'            = $OperatingSystem.OSArchitecture

    'OS Install Date'            = $OperatingSystem.InstallDate

    'Last Boot Time'             = $OperatingSystem.LastBootUpTime

    'Liongard MachineID'         = $LiongardMachineID

    'Liongard Registry Path'     = $LiongardRegistryPath

    'Liongard Agent Service'     = $LiongardServices

    'Liongard Installed Software'= $LiongardSoftware

    'Network Adapters'            = $NetworkAdapters
}

# ------------------------------------------------------------
# Display Main Information
# ------------------------------------------------------------

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " WINDOWS / LIONGARD IDENTITY INFORMATION" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

$Output | Format-List

# ------------------------------------------------------------
# Display Liongard Registry Details
# ------------------------------------------------------------

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " LIONGARD REGISTRY DETAILS" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

if ($LiongardRegistryProperties) {
    $LiongardRegistryProperties |
        Format-List
}
else {
    Write-Host "Liongard Agent registry information was not found." `
        -ForegroundColor Yellow
}

# ------------------------------------------------------------
# Display Liongard Software
# ------------------------------------------------------------

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " LIONGARD INSTALLED SOFTWARE" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

if ($LiongardSoftware) {
    $LiongardSoftware | Format-List
}
else {
    Write-Host "Liongard Agent installation was not detected." `
        -ForegroundColor Yellow
}

# ------------------------------------------------------------
# Display Network Information
# ------------------------------------------------------------

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " NETWORK INFORMATION" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

$NetworkAdapters | Format-List

# ------------------------------------------------------------
# Identity Conflict Check
# ------------------------------------------------------------

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " IDENTITY CHECK" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

if ([string]::IsNullOrWhiteSpace($MachineGuid)) {

    Write-Host "Windows MachineGuid : NOT FOUND" -ForegroundColor Yellow
}
else {

    Write-Host "Windows MachineGuid : $MachineGuid"
}

if ([string]::IsNullOrWhiteSpace($LiongardMachineID)) {

    Write-Host "Liongard MachineID  : NOT FOUND" -ForegroundColor Yellow
}
else {

    Write-Host "Liongard MachineID  : $LiongardMachineID"
}

if ($MachineGuid -and $LiongardMachineID) {

    Write-Host ""
    Write-Host "Both Windows MachineGuid and Liongard MachineID were detected." `
        -ForegroundColor Green
}

Write-Host ""
Write-Host "Collection complete." -ForegroundColor Cyan

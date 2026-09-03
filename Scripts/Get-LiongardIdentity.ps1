<#
.SYNOPSIS
    Collects Windows and Liongard Agent identity information.

.DESCRIPTION
    Read-only diagnostic script for investigating Liongard Agent
    identity conflicts caused by cloned Windows machines.

    Collects:
    - Computer name
    - FQDN
    - Domain / Workgroup
    - Manufacturer
    - Model
    - BIOS serial number
    - BIOS version
    - System UUID
    - Windows MachineGuid
    - Windows Product ID
    - Windows version/build
    - OS installation date
    - Last boot time
    - Liongard MachineID
    - Liongard Agent service
    - Liongard Agent version
    - Liongard registry information
    - Network information

.NOTES
    READ-ONLY.
    This script does not modify Windows or Liongard.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'SilentlyContinue'

# ============================================================
# SYSTEM INFORMATION
# ============================================================

$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
$ComputerSystemProduct = Get-CimInstance -ClassName Win32_ComputerSystemProduct
$BIOS = Get-CimInstance -ClassName Win32_BIOS
$OperatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem

# ============================================================
# FQDN
# ============================================================

$FQDN = $null

try {
    $FQDN = [System.Net.Dns]::GetHostEntry(
        $env:COMPUTERNAME
    ).HostName
}
catch {
    $FQDN = $env:COMPUTERNAME
}

# ============================================================
# WINDOWS MACHINE GUID
# ============================================================

$MachineGuid = $null

try {
    $MachineGuid = (
        Get-ItemProperty `
            -Path 'HKLM:\SOFTWARE\Microsoft\Cryptography' `
            -Name 'MachineGuid' `
            -ErrorAction Stop
    ).MachineGuid
}
catch {
    $MachineGuid = 'NOT FOUND'
}

# ============================================================
# WINDOWS INFORMATION
# ============================================================

$WindowsRegistryPath = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'

$WindowsProductName = $null
$WindowsDisplayVersion = $null
$WindowsBuild = $null
$WindowsUBR = $null
$WindowsProductID = $null

try {
    $WindowsInfo = Get-ItemProperty -Path $WindowsRegistryPath

    $WindowsProductName = $WindowsInfo.ProductName
    $WindowsDisplayVersion = $WindowsInfo.DisplayVersion
    $WindowsBuild = $WindowsInfo.CurrentBuild
    $WindowsUBR = $WindowsInfo.UBR
    $WindowsProductID = $WindowsInfo.ProductId
}
catch {
}

$FullWindowsBuild = $WindowsBuild

if ($WindowsBuild -and ($null -ne $WindowsUBR)) {
    $FullWindowsBuild = "$WindowsBuild.$WindowsUBR"
}

# ============================================================
# LIONGARD REGISTRY
# ============================================================

$LiongardRegistryPaths = @(
    'HKLM:\SOFTWARE\Liongard\LiongardAgent',
    'HKLM:\SOFTWARE\WOW6432Node\Liongard\LiongardAgent'
)

$LiongardMachineID = $null
$LiongardRegistryPath = $null
$LiongardRegistryProperties = $null

foreach ($Path in $LiongardRegistryPaths) {

    if (Test-Path $Path) {

        try {
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
        catch {
        }
    }
}

if (-not $LiongardMachineID) {
    $LiongardMachineID = 'NOT FOUND'
}

if (-not $LiongardRegistryPath) {
    $LiongardRegistryPath = 'NOT FOUND'
}

# ============================================================
# LIONGARD SERVICES
# ============================================================

$LiongardServices = @(
    Get-CimInstance -ClassName Win32_Service |
        Where-Object {
            $_.Name -match 'Liongard|Roar' -or
            $_.DisplayName -match 'Liongard|Roar'
        } |
        Select-Object Name,
                      DisplayName,
                      State,
                      Status,
                      StartMode,
                      StartName,
                      PathName
)

# ============================================================
# LIONGARD INSTALLED SOFTWARE
# ============================================================

$LiongardSoftware = @()

$UninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)

foreach ($Path in $UninstallPaths) {

    $LiongardSoftware += @(
        Get-ItemProperty -Path $Path |
            Where-Object {
                $_.DisplayName -match 'Liongard'
            } |
            Select-Object DisplayName,
                          DisplayVersion,
                          Publisher,
                          InstallDate,
                          InstallLocation,
                          UninstallString
    )
}

# ============================================================
# NETWORK INFORMATION
# ============================================================

$NetworkAdapters = @(
    Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration |
        Where-Object {
            $_.IPEnabled -eq $true
        } |
        Select-Object Description,
                      MACAddress,
                      IPAddress,
                      IPSubnet,
                      DefaultIPGateway,
                      DNSServerSearchOrder
)

# ============================================================
# MAIN INFORMATION OBJECT
# ============================================================

$Output = [PSCustomObject]@{

    'Computer Name'               = $env:COMPUTERNAME

    'FQDN'                        = $FQDN

    'Domain / Workgroup'          = $ComputerSystem.Domain

    'Domain Role'                 = $ComputerSystem.DomainRole

    'Manufacturer'                = $ComputerSystem.Manufacturer

    'Model'                       = $ComputerSystem.Model

    'BIOS Serial Number'           = $BIOS.SerialNumber

    'BIOS Version'                 = ($BIOS.SMBIOSBIOSVersion -join ', ')

    'System UUID'                  = $ComputerSystemProduct.UUID

    'Windows MachineGuid'          = $MachineGuid

    'Windows Product ID'           = $WindowsProductID

    'Windows Product Name'         = $WindowsProductName

    'Windows Display Version'      = $WindowsDisplayVersion

    'Windows Build'                = $FullWindowsBuild

    'Windows Version'              = $OperatingSystem.Version

    'OS Architecture'              = $OperatingSystem.OSArchitecture

    'OS Install Date'              = $OperatingSystem.InstallDate

    'Last Boot Time'                = $OperatingSystem.LastBootUpTime

    'Liongard MachineID'            = $LiongardMachineID

    'Liongard Registry Path'        = $LiongardRegistryPath

    'Liongard Agent Service'        = $LiongardServices

    'Liongard Installed Software'   = $LiongardSoftware

    'Network Adapters'               = $NetworkAdapters
}

# ============================================================
# DISPLAY MAIN INFORMATION
# ============================================================

Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' WINDOWS / LIONGARD DEVICE IDENTITY INFORMATION' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

$Output | Format-List

# ============================================================
# LIONGARD REGISTRY DETAILS
# ============================================================

Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' LIONGARD REGISTRY DETAILS' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

if ($LiongardRegistryProperties) {

    $LiongardRegistryProperties |
        Format-List

}
else {

    Write-Host 'Liongard Agent registry information was not found.' `
        -ForegroundColor Yellow
}

# ============================================================
# LIONGARD SOFTWARE
# ============================================================

Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' LIONGARD INSTALLED SOFTWARE' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

if ($LiongardSoftware.Count -gt 0) {

    $LiongardSoftware |
        Format-List

}
else {

    Write-Host 'Liongard Agent installation was not detected.' `
        -ForegroundColor Yellow
}

# ============================================================
# NETWORK INFORMATION
# ============================================================

Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' NETWORK INFORMATION' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

if ($NetworkAdapters.Count -gt 0) {

    $NetworkAdapters |
        Format-List

}
else {

    Write-Host 'No active network adapters found.' `
        -ForegroundColor Yellow
}

# ============================================================
# IDENTITY SUMMARY
# ============================================================

Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' IDENTITY SUMMARY' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

Write-Host "Computer Name       : $env:COMPUTERNAME"
Write-Host "BIOS Serial Number  : $($BIOS.SerialNumber)"
Write-Host "System UUID         : $($ComputerSystemProduct.UUID)"
Write-Host "Windows MachineGuid : $MachineGuid"
Write-Host "Liongard MachineID  : $LiongardMachineID"
Write-Host ''

# ============================================================
# BASIC DUPLICATE INDICATOR
# ============================================================

if (
    $MachineGuid -and
    $LiongardMachineID -and
    $MachineGuid -ne 'NOT FOUND' -and
    $LiongardMachineID -ne 'NOT FOUND'
) {

    Write-Host 'Identity information successfully collected.' `
        -ForegroundColor Green
}
else {

    Write-Host 'One or more identity values could not be detected.' `
        -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'Collection complete.' -ForegroundColor Cyan
Write-Host ''

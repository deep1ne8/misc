#Requires -Version 5.1
#Requires -RunAsAdministrator

[CmdletBinding()]
param (
    [Parameter(Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$DownloadPath = "$env:TEMP\Autodesk",
    
    [Parameter(Position = 1)]
    [ValidateNotNullOrEmpty()]
    [string]$SourceUrl = "https://up1.autodesk.com/2024/RVT/63F8F057-85FB-337D-8493-CD003BAEAC52/Revit_2024_2_0.exe",
    
    [Parameter(Position = 2)]
    [switch]$Silent,
    
    [Parameter(Position = 3)]
    [switch]$NoRestart
)

function Write-LogMessage {
    param (
        [string]$Message,
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMessage = "[$timestamp] [$Level] $Message"
    
    switch ($Level) {
        'Info'    { Write-Host $formattedMessage -ForegroundColor Cyan }
        'Warning' { Write-Host $formattedMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $formattedMessage -ForegroundColor Red }
    }
}

function Get-InstalledRevitVersions {
    $revitKeyPath = "HKLM:\SOFTWARE\Autodesk\Revit\InstalledVersions"
    if (Test-Path $revitKeyPath) {
        $installedVersions = Get-ChildItem -Path $revitKeyPath | ForEach-Object { $_.PSChildName }
        if ($installedVersions) {
            Write-LogMessage "Installed Revit Versions: $($installedVersions -join ', ')"
        } else {
            Write-LogMessage "No previous versions of Revit found."
        }
    } else {
        Write-LogMessage "Revit is not installed."
    }
}

function Test-Requirements {
    Write-LogMessage "Checking system requirements..."
    
    $drive = Split-Path -Qualifier $DownloadPath
    $freeSpace = (Get-WmiObject -Query "SELECT FreeSpace FROM Win32_LogicalDisk WHERE DeviceID='$drive'").FreeSpace / 1GB
    if ($freeSpace -lt 30) {
        throw "Insufficient disk space: $([math]::Round($freeSpace, 2))GB available, minimum 30GB required."
    }
    
    Write-LogMessage "System requirements check passed."
}

function Initialize-DownloadLocation {
    if (-not (Test-Path -Path $DownloadPath)) {
        New-Item -Path $DownloadPath -ItemType Directory -Force | Out-Null
        Write-LogMessage "Created download directory: $DownloadPath"
    }
    return Join-Path -Path $DownloadPath -ChildPath (Split-Path -Leaf $SourceUrl)
}

function Get-InstallerFile {
    param (
        [string]$Url,
        [string]$Destination
    )
    Write-LogMessage "Downloading installer from $Url..."
    Invoke-WebRequest -Uri $Url -OutFile $Destination
    Write-LogMessage "Download completed: $Destination"
}

function Install-Revit {
    param (
        [string]$InstallerPath
    )
    Write-LogMessage "Starting Revit installation..."
    $args = "/q"
    Start-Process -FilePath $InstallerPath -ArgumentList $args -Wait -PassThru
    Write-LogMessage "Installation completed."
}

# Main execution
try {
    Write-LogMessage "Starting Revit installation process..."
    Get-InstalledRevitVersions
    Test-Requirements
    $installerPath = Initialize-DownloadLocation
    if (-not (Test-Path -Path $installerPath)) {
        Get-InstallerFile -Url $SourceUrl -Destination $installerPath
    }
    Install-Revit -InstallerPath $installerPath
    Write-LogMessage "Revit installation process completed."
} catch {
    Write-LogMessage "Error: $_" -Level Error
}

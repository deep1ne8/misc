<#
.SYNOPSIS
    Installs Windows Package Manager (winget) via the latest App Installer package.
.DESCRIPTION
    - Verifies if winget is present.
    - If missing, queries GitHub for the latest release of Microsoft.DesktopAppInstaller (.msixbundle).
    - Downloads and installs the bundle via Add-AppxPackage.
    - Requires Administrator privileges and that sideloading of signed packages is allowed.
.NOTES
    Tested on Windows 10/11 with PowerShell 5.1+
#>

# Ensure we're running elevated
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Write-Error "This script must be run as Administrator."
    exit 1
}

# Check for existing winget
try {
    $null = Get-Command winget -ErrorAction Stop
    $version = winget --version 2>$null
    Write-Output "✔ winget is already installed ($version)."
    exit 0
}
catch {
    Write-Output "ℹ winget not detected. Proceeding with installation..."
}

# GitHub API endpoint for latest winget-cli release
$apiUrl = 'https://api.github.com/repos/microsoft/winget-cli/releases/latest'

try {
    # Retrieve release metadata with proper error handling
    $release = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing -ErrorAction Stop
}
catch {
    Write-Error "Failed to query GitHub API: $($_.Exception.Message)"
    exit 1
}

# Find the MSIX bundle asset
$bundle = $release.assets |
    Where-Object { $_.name -match 'Microsoft\.DesktopAppInstaller_.*\.msixbundle$' } |
    Sort-Object -Property name -Descending |
    Select-Object -First 1

if (-not $bundle) {
    Write-Error "Could not locate the App Installer bundle in the latest release."
    exit 1
}

$downloadUrl = $bundle.browser_download_url
$outFile = Join-Path $env:TEMP $bundle.name

Write-Output "➡ Downloading `"$($bundle.name)`" from GitHub..."
try {
    Invoke-WebRequest -Uri $downloadUrl -OutFile $outFile -UseBasicParsing -ErrorAction Stop
}
catch {
    Write-Error "Download failed: $($_.Exception.Message)"
    exit 1
}

Write-Output "➡ Installing App Installer package..."
try {
    Add-AppxPackage -Path $outFile -ErrorAction Stop
}
catch {
    Write-Error "Installation failed: $($_.Exception.Message)"
    exit 1
}

# Verify installation
try {
    $null = Get-Command winget -ErrorAction Stop
    $version = winget --version 2>$null
    Write-Output "✔ winget successfully installed. Version: $version"
}
catch {
    Write-Error "Installation seemed to complete, but winget is still not available."
    exit 1
}

# Clean up
Remove-Item $outFile -ErrorAction SilentlyContinue
Write-Output "🧹 Temporary files cleaned up."

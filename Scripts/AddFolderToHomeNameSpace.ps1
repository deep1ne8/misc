# Add-FolderToExplorerHome.ps1
# This script adds a custom folder directly under the Home section in File Explorer

# Run with administrative privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "This script requires administrative privileges. Please run as Administrator."
    exit
}

# Configuration - Edit these values
$folderPath = "C:\Users\jjfnk\JJ - RegeneRx Biopharmaceuticals, Inc" # Full path to the folder you want to add
$displayName = "JJ - RegeneRx Biopharmaceuticals, Inc"      # The name to display in Explorer

# Get input from user if default values are unchanged
if ($null -eq $folderPath) {
    $folderPath = Read-Host "Enter the full path to the folder you want to add to Home"
}

if ($null -eq $displayName) {
    $displayName = Read-Host "Enter the display name for this folder in Explorer"
}

# Ensure the folder exists
if (-not (Test-Path -Path $folderPath)) {
    Write-Error "The specified folder does not exist: $folderPath"
    return
}

# Generate a unique GUID for the namespace
$clsidGuid = [guid]::NewGuid().ToString("B")
Write-Host "Generated GUID: $clsidGuid" -ForegroundColor Cyan

# Create registry keys for Explorer namespace
try {
    # Create CLSID entry
    $clsidPath = "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$clsidGuid"
    New-Item -Path $clsidPath -Force | Out-Null
    Set-ItemProperty -Path $clsidPath -Name "(Default)" -Value $displayName
    
    # Create DefaultIcon entry
    $iconPath = Join-Path $clsidPath "DefaultIcon"
    New-Item -Path $iconPath -Force | Out-Null
    Set-ItemProperty -Path $iconPath -Name "(Default)" -Value "shell32.dll,4" # Use folder icon
    
    # Create InProcServer32 entry
    $serverPath = Join-Path $clsidPath "InProcServer32"
    New-Item -Path $serverPath -Force | Out-Null
    Set-ItemProperty -Path $serverPath -Name "(Default)" -Value "%SystemRoot%\system32\shell32.dll"
    
    # Create Instance entry
    $instancePath = Join-Path $clsidPath "Instance"
    New-Item -Path $instancePath -Force | Out-Null
    Set-ItemProperty -Path $instancePath -Name "CLSID" -Value "{0E5AAE11-A475-4c5b-AB00-C66DE400274E}"
    
    # Create InitPropertyBag
    $bagPath = Join-Path $instancePath "InitPropertyBag"
    New-Item -Path $bagPath -Force | Out-Null
    Set-ItemProperty -Path $bagPath -Name "Attributes" -Value 0x11 -Type DWord
    Set-ItemProperty -Path $bagPath -Name "TargetFolderPath" -Value $folderPath
    
    # Create Shell entry to hide from navigation pane (if desired)
    $shellPath = Join-Path $clsidPath "ShellFolder"
    New-Item -Path $shellPath -Force | Out-Null
    Set-ItemProperty -Path $shellPath -Name "Attributes" -Value 0xF080004D -Type DWord
    Set-ItemProperty -Path $shellPath -Name "FolderValueFlags" -Value 0x28 -Type DWord

    # Add to HomeFolder (the actual Home section in File Explorer)
    $homeFolderPath = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HomeFolder\NameSpace"
    if (-not (Test-Path $homeFolderPath)) {
        New-Item -Path $homeFolderPath -Force | Out-Null
    }
    
    $namespaceEntryPath = Join-Path $homeFolderPath $clsidGuid
    New-Item -Path $namespaceEntryPath -Force | Out-Null
    
    # Also add to NameSpace for legacy Explorer views
    $namespacePath = "Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace"
    $namespaceEntryPath = Join-Path $namespacePath $clsidGuid
    New-Item -Path $namespaceEntryPath -Force | Out-Null
    
    Write-Host "Successfully added '$displayName' to Explorer Home" -ForegroundColor Green
    Write-Host "Please restart Explorer or sign out and back in to see the changes" -ForegroundColor Yellow
    
    # Offer to restart Explorer
    $restartExplorer = Read-Host "Would you like to restart Explorer now? (Y/N)"
    if ($restartExplorer -eq 'Y') {
        Write-Host "Restarting Explorer..." -ForegroundColor Cyan
        Get-Process explorer | Stop-Process
        Start-Sleep -Seconds 2
        Start-Process explorer
    }
    
} catch {
    Write-Error "Failed to modify registry: $_"
}

# Provide removal instructions
Write-Host "`nTo remove this folder from Explorer Home in the future, run the following commands:" -ForegroundColor Cyan
Write-Host "Remove-Item -Path 'Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$clsidGuid' -Recurse -Force" -ForegroundColor Yellow
Write-Host "Remove-Item -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\HomeFolder\NameSpace\$clsidGuid' -Force" -ForegroundColor Yellow
Write-Host "Remove-Item -Path 'Registry::HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\$clsidGuid' -Force" -ForegroundColor Yellow
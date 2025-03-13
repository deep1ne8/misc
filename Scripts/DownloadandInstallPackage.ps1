# Function to install a remote package
function Install-RemotePackage {
    param (
        [string]$Url,
        [string]$DestinationPath = "C:\Temp\$(Split-Path -Leaf $Url)",
        [string]$LogPath,
        [switch]$CleanupAfterInstall,
        [switch]$Force
    )
    
    Write-Host "Starting installation for: $Url"
    
    if (-Not (Test-Path "C:\Temp")) {
        New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
    }
    
    if (Test-Path $DestinationPath -and -Not $Force) {
        Write-Host "File already exists at $DestinationPath. Use -Force to overwrite." -ForegroundColor Yellow
    } else {
        Write-Host "Downloading file to: $DestinationPath"
        Invoke-WebRequest -Uri $Url -OutFile $DestinationPath
    }
    
    if ($Url -match "\.msi$") {
        $Arguments = "/i `"$DestinationPath`" /qn"
        if ($LogPath) { $Arguments += " /log `"$LogPath`"" }
        Write-Host "Executing MSI installation with arguments: $Arguments"
        Start-Process -FilePath "msiexec.exe" -ArgumentList $Arguments -Wait -NoNewWindow
    } elseif ($Url -match "\.exe$") {
        Write-Host "Executing EXE installation: $DestinationPath"
        Start-Process -FilePath $DestinationPath -ArgumentList "/S" -Wait -NoNewWindow
    } else {
        Write-Host "Unsupported file type for installation." -ForegroundColor Red
        return
    }
    
    Write-Host "Installation completed for: $Url" -ForegroundColor Green
    
    if ($CleanupAfterInstall) {
        Write-Host "Cleaning up installer: $DestinationPath"
        Remove-Item -Path $DestinationPath -Force
    }
}

# Basic silent installation of an EXE
Install-RemotePackage -Url "https://download.example.com/application.exe"

# Silent MSI installation with custom path and cleanup
Install-RemotePackage -Url "https://download.example.com/package.msi" `
                     -DestinationPath "D:\Software\package.msi" `
                     -LogPath "D:\Logs\install.log" `
                     -CleanupAfterInstall

# Force reinstallation with verbose output
Install-RemotePackage -Url "https://download.example.com/setup.exe" `
                     -Force

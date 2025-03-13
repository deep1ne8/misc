# Function to install a remote package
function Install-RemotePackage {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Url = (Read-Host "Enter the URL of the remote package to install"),
        
        [Parameter(Mandatory=$false)]
        [string]$DestinationPath = "C:\Temp\$(Split-Path -Leaf $Url)",
        
        [Parameter(Mandatory=$false)]
        [string]$LogPath,
        
        [Parameter(Mandatory=$false)]
        [switch]$CleanupAfterInstall,
        
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    
    if ($LogPath) {
        $LogPath = Join-Path -Path $LogPath -ChildPath "$(Split-Path -Leaf $Url).log"
    }
    
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

Install-RemotePackage

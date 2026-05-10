[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; $tempDir = Join-Path $env:TEMP "WingetInstall"; 
New-Item -Path $tempDir -ItemType Directory -Force | Out-Null; 
Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.3" -OutFile "$tempDir\xaml.nupkg"; 
Rename-Item -Path "$tempDir\xaml.nupkg" -NewName "xaml.zip"; 
Expand-Archive -Path "$tempDir\xaml.zip" -DestinationPath "$tempDir\xaml" -Force; 
$latest = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest"; 
$msixBundle = ($latest.assets | Where-Object { $_.name -match '.msixbundle' } | Select-Object -First 1).browser_download_url; $license = ($latest.assets | Where-Object { $_.name -match '.xml' } | Select-Object -First 1).browser_download_url; 
Invoke-WebRequest -Uri $msixBundle -OutFile "$tempDir\winget.msixbundle"; 
Invoke-WebRequest -Uri $license -OutFile "$tempDir\license.xml"; 
Add-AppxPackage -Path "$tempDir\xaml\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx"; 
Add-AppxProvisionedPackage -Online -PackagePath "$tempDir\winget.msixbundle" -LicensePath "$tempDir\license.xml" -ErrorAction SilentlyContinue; Remove-Item -Path $tempDir -Recurse -Force; 
Write-Host "Winget installed. Restart your terminal or system if needed."

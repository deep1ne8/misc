# Interactive PowerShell Script to Add a Trusted Network Location for Office Apps

Add-Type -AssemblyName Microsoft.VisualBasic

# Prompt for UNC path
$trustedPath = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter the UNC path to your trusted macro location (e.g. \\fileserver\macros):",
    "Trusted Location Path",
    "\\YourServer\TrustedMacros"
)

if (-not $trustedPath -or $trustedPath -eq "") {
    Write-Host "No path provided. Exiting..." -ForegroundColor Red
    exit 1
}

# Prompt for description
$description = [Microsoft.VisualBasic.Interaction]::InputBox(
    "Enter a description for this trusted location:",
    "Trusted Location Description",
    "Trusted Macro Network Share"
)

if (-not $description -or $description -eq "") {
    $description = "Trusted Macro Network Share"
}

# Office versions to apply
$officeVersions = @("16.0") # Office 2016/2019/2021/365

# Office apps to apply
$officeApps = @("Word", "Excel", "PowerPoint", "Access")

foreach ($version in $officeVersions) {
    foreach ($app in $officeApps) {
        $baseKey = "HKCU:\Software\Microsoft\Office\$version\$app\Security\Trusted Locations"

        try {
            # Allow network locations
            New-Item -Path $baseKey -Force | Out-Null
            Set-ItemProperty -Path $baseKey -Name "AllowNetworkLocations" -Value 1

            # Create unique trusted location key (e.g., Location1, Location2...)
            $i = 1
            do {
                $locationKeyPath = "$baseKey\Location$i"
                $exists = Test-Path $locationKeyPath
                $i++
            } while ($exists)

            # Add trusted location
            New-Item -Path $locationKeyPath -Force | Out-Null
            Set-ItemProperty -Path $locationKeyPath -Name "Path" -Value "$trustedPath\"
            Set-ItemProperty -Path $locationKeyPath -Name "AllowSubFolders" -Value 1
            Set-ItemProperty -Path $locationKeyPath -Name "Description" -Value $description
            Set-ItemProperty -Path $locationKeyPath -Name "Date" -Value (Get-Date).ToString()

            Write-Host "$($app) $($version): ✅ Trusted location added -> $trustedPath" -ForegroundColor Green
        } catch {
            Write-Warning "$($app) $($version): ❌ Failed to add trusted location. $_"
        }
    }
}

Write-Host "`nAll done. Please reopen your Office apps to apply the changes." -ForegroundColor Cyan

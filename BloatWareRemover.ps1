function Uninstall-DellBloatware {
    $appNames = @(
        "Dell SupportAssist",
        "Dell SupportAssist OS Recovery Plugin for Dell Update",
        "Dell SupportAssist Remediation",
        "Dell Optimizer",
        "Dell Display Manager",
        "Dell Peripheral Manager",
        "Dell Pair",
        "Dell Core Services",
        "Dell Trusted Device"
    )

    Write-Host "Uninstalling Dell bloatware..." -ForegroundColor Yellow
    foreach ($appName in $appNames) {
        $app = Get-Package -Name $appName -ErrorAction SilentlyContinue
        if ($app) {
            Write-Host "Uninstalling $appName..." -ForegroundColor Cyan
            $app | Uninstall-Package -Confirm:$false
        } else {
            Write-Host "$appName not found." -ForegroundColor Magenta
        }
    }
}


function Remove-OfficeLanguages {
# Define ODT folder and setup path
$odtFolder = "C:\ODT"
$setupPath = "$odtFolder\setup.exe"
$xmlPath = "$odtFolder\RemoveLanguages.xml"
$downloadUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/ODTTool/setup.exe"

Write-Host "Starting Office Language Remover..." -ForegroundColor Yellow
Start-Sleep -Seconds 2
Write-Host "`n"
# Create ODT directory if it doesn't exist
if (!(Test-Path $odtFolder)) {
    Write-Host "Creating ODT directory at $odtFolder..." -ForegroundColor Yellow
    New-Item -Path $odtFolder -ItemType Directory -Force | Out-Null
}

# Download setup.exe if not found
if (!(Test-Path $setupPath)) {
    Write-Host "Downloading Office Deployment Tool (ODT)..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $setupPath
} else {
    Write-Host "setup.exe already exists in C:\ODT. Skipping download." -ForegroundColor Green
}

# Verify download
if (!(Test-Path $setupPath)) {
    Write-Host "❌ Download failed. Check the URL or try again." -ForegroundColor Red
    return
}

Write-Host "✅ ODT setup.exe is ready at $setupPath" -ForegroundColor Green

# Try to retrieve installed Office languages
$officeLanguages = @()
$registryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Scenario"
)

foreach ($regPath in $registryPaths) {
    if (Test-Path $regPath) {
        $regData = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if ($regData.InstallationLanguages) {
            $officeLanguages += $regData.InstallationLanguages -split ";"
        } elseif ($regData.InstallationLanguage) {
            $officeLanguages += $regData.InstallationLanguage
        }
    }
}

# Remove duplicates and keep only unique values
$installedLanguages = $officeLanguages | Select-Object -Unique

if ($installedLanguages.Count -eq 0) {
    Write-Host "No additional Office languages found. No action needed." -ForegroundColor Cyan
    exit
}

# Exclude en-us and keep only the unwanted languages
$unwantedLanguages = $installedLanguages | Where-Object { $_ -ne "en-us" }

if ($unwantedLanguages.Count -eq 0) {
    Write-Host "Only en-us is installed. No need to remove languages." -ForegroundColor Cyan
    exit
}

# Create XML Configuration
$xmlContent = @"
<Configuration>
    <Remove>
        <Product ID="O365ProPlusRetail">
            $(foreach ($lang in $unwantedLanguages) { "<Language ID=`"$lang`" />" } -join "`n")
        </Product>
    </Remove>
</Configuration>
"@

# Save the XML File
$xmlContent | Set-Content -Path $xmlPath -Encoding UTF8

Write-Host "Generated RemoveLanguages.xml with the following languages:" -ForegroundColor Yellow
$unwantedLanguages | ForEach-Object { Write-Host " - $_" -ForegroundColor Magenta }

# Run Office Deployment Tool to remove the languages
Write-Host "`nStarting Office Deployment Tool to remove unwanted languages..." -ForegroundColor Green
Start-Process -FilePath $setupPath -ArgumentList "/configure $xmlPath" -NoNewWindow -Wait
Write-Host "Office language removal process completed." -ForegroundColor Green
}


# Run the function
Uninstall-DellBloatware
Write-Host "`n"
Start-Sleep -Seconds 3
Remove-OfficeLanguages
return
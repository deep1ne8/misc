<#
Dell Bloatware and Office Language Remover

Author: EDaniels
Date: 12/2024

Description:
This script removes Dell Bloatware and unwanted Office languages.

Usage:
Run the script and follow the prompts to uninstall Dell bloatware and remove unwanted Office languages.

Requirements:
- PowerShell 5.1 or later


#>

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
        if ($null -eq $app) {
            Write-Host "${appName} not installed." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Stopping uninstallation process and moving on to office language remover..." -ForegroundColor Magenta
            Remove-OfficeLanguages
            return
        } elseif ($app) {
            Write-Host "Uninstalling $appName..." -ForegroundColor Cyan
            $app | Uninstall-Package
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
$ListInstalledLanguages = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources" -Name "UIFallbackLanguages").UIFallbackLanguages
$ODTlog = Join-Path -Path $odtFolder -ChildPath "ODTlog" -ErrorAction SilentlyContinue

if (-not (Test-Path $ODTlog -PathType Container)) {
    Write-Host "Verbose: ODT log folder $ODTlog not found. Creating log folder." -foregroundColor Red
    New-Item -Path $ODTlog -ItemType Directory
}

if ($null -eq $ListInstalledLanguages) {
    Write-Host "`n"
    Write-Host "VERBOSE: No installed languages found." -foregroundColor Yellow
    Write-Host "VERBOSE: Skipping Office Language Remover..." -ForegroundColor Red
    Write-Host "Verbose: Script completed..." -ForegroundColor Green
    return
}

Write-Host "Verbose: Starting Office Language Remover..." -ForegroundColor Yellow
Start-Sleep -Seconds 2
Write-Host "`n"
# Create ODT directory if it doesn't exist
if (!(Test-Path $odtFolder -PathType Container) -or !(Test-Path $ODTlog -PathType Container)) {
    Write-Host "Verbose: Creating ODT directory at $odtFolder and log folder $ODTlog..." -ForegroundColor Yellow
    try {
        New-Item -Path $odtFolder -ItemType Directory -Force -ErrorAction SilentlyContinue
        New-Item -Path $ODTlog -ItemType Directory -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Host "Verbose: Failed to create ODT directory at $odtFolder or log folder $ODTlog. Error: $_" -foregroundColor Red
        return
    }
}

# Download setup.exe if not found
if (!(Test-Path $setupPath)) {
    Write-Host "Verbose: Downloading Office Deployment Tool (ODT)..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $setupPath
} else {
    Write-Host "Verbose: setup.exe already exists in C:\ODT. Skipping download." -ForegroundColor Green
}

# Verify download
if (!(Test-Path $setupPath)) {
    Write-Host "❌ Download failed. Check the URL or try again." -ForegroundColor Red
    return
}

Write-Host "✅ ODT setup.exe is ready at $setupPath" -ForegroundColor Green

<#
# Define the languages you want to uninstall
$unwantedLanguages = @(
    "af-za", "sq-al", "ar-sa", "hy-am", "bn-in", "bg-bg", "ca-es", "hr-hr", "cs-cz", 
    "da-dk", "nl-nl", "et-ee", "fi-fi", "fr-fr", "de-de", "el-gr", "gu-in", 
    "he-il", "hi-in", "hu-hu", "is-is", "id-id", "it-it", "ja-jp", "kn-in", "ko-kr", 
    "la-lt", "lv-lv", "lt-lt", "mk-mk", "ml-in", "mr-in", "nb-no", "pl-pl", "pt-br", 
    "pt-pt", "pa-in", "ro-ro", "ru-ru", "sr-latn-rs", "sk-sk", "sl-si", "es-es", "sv-se", 
    "ta-in", "tr-tr", "uk-ua", "vi-vn", "cy-gb", "xh-za", "zu-za"
)
#>

# Split the string by semicolons and exclude x-none and en-us
$languageArray = ($ListInstalledLanguages -split ";") | Where-Object { $_ -ne "x-none" -and $_ -ne "en-us" }

# Format the output
Write-Host "UIFallbackLanguages contains the following languages (excluding x-none):"
$unwantedLanguages = $languageArray | ForEach-Object { 
    Write-Host "- $UnwantedLanguages" -ForegroundColor Cyan
}

# Generate XML content to remove unwanted languages
$xmlContent = @"
<Configuration>
    <Remove>
        <Product ID="LanguagePack">
            $(foreach ($lang in $unwantedLanguages) { "<Language ID=`"$lang`" />" } -join "`n")
        </Product>
    </Remove>
    <Display Level="None" AcceptEULA="TRUE"/>
    <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
    <Logging Level="Standard" Path="$ODTlog" />
</Configuration>
"@

# Save the XML content to file
$xmlContent | Set-Content -Path $xmlPath -Encoding UTF8

Write-Host "`n"
Write-Host "Verbose: Generated RemoveLanguages.xml with the following languages:" -ForegroundColor Yellow

# List the unwanted languages
Write-Host "Verbose: Unwanted languages to be removed:" -ForegroundColor Yellow
Write-Host "`n"

$unwantedLanguages

Write-Host "`n"
Start-Sleep -Seconds 2
Write-Host "Are you sure you want to remove these languages? (Y/N) " -ForegroundColor Yellow -NoNewline
$confirmation = Read-Host
if ($confirmation -ne "Y" -and $confirmation -ne "y") {
    Write-Host "Operation canceled." -ForegroundColor Red
    return
}else {
# Run Office Deployment Tool to remove the languages
Write-Host "`nStarting Office Deployment Tool to remove unwanted languages..." -ForegroundColor Green
Start-Process -FilePath $setupPath -ArgumentList "/configure $xmlPath" -NoNewWindow -Wait
Write-Host "`n"
Get-Output -Path "$ODTlog\*.log" -Tail 100
Write-Host "`n"
Write-Host "Verbose: Office language removal process completed." -ForegroundColor Green

Write-Host "`n"
Write-Host "Please restart your computer for the changes to take effect." -ForegroundColor Yellow
    }
}


# Run the function
$manufacturer = (Get-WmiObject -Class Win32_ComputerSystem).Manufacturer
if ($manufacturer -like "*Dell*") {
    Write-Host "Verbose: This device is a $manufacturer. Proceeding with removal of Dell Bloatware."
    Uninstall-DellBloatware
    Write-Host "`n"
    Start-Sleep -Seconds 3
    Remove-OfficeLanguages
    return
}else {
    Write-Host "Verbose: This device is not a Dell. Skipping removal of Dell bloatware. Starting Office Language Remover..." -ForegroundColor Yellow
    Write-Host "`n"
    Start-Sleep -Seconds 3
    Remove-OfficeLanguages
    return
}

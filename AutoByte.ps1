Write-Host "`n=== Configuring System for Secure TLS 1.2 Connections ===`n" -ForegroundColor Cyan

# --- Permanent TLS 1.2 Fix ---
# Apply at runtime:
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Write-Host "✔ TLS 1.2 enforced for this session." -ForegroundColor Green

# Apply permanent fix in registry (for future .NET and PowerShell sessions):
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\.NETFramework\v4.0.30319",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\.NETFramework\v4.0.30319"
)
foreach ($path in $regPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
    New-ItemProperty -Path $path -Name "SchUseStrongCrypto" -Value 1 -PropertyType DWord -Force | Out-Null
}
Write-Host "✔ Registry updated to permanently enable strong crypto/TLS 1.2." -ForegroundColor Green


$GitHubScripts = @(

    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/TeamsModernMigration.ps1"; Description = "Teams Classic to Modern Migration Tool" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/OfficeWindowsCompat.ps1"; Description = "Office and Windows Compatibility Check" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/Get-PrinterSupplies.ps1"; Description = "Get Printer Supplies" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Scripts/WindowsOnlineRepair.ps1"; Description = "WindowsOnlineRepair" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/GetWindowsEvents.ps1"; Description = "Get Windows Events" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DellCommandUpdate.ps1"; Description = "Dell Command Update" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1"; Description = "ResetandInstallWindowsUpdate" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/WindowsSystemRepair.ps1"; Description = "WindowsSystemRepair" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"; Description = "ResetandClearWindowsSearchdb" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Scripts/M365_Office32bitRepair.ps1"; Description = "M365_DeployOffice32bit" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1"; Description = "CheckDriveSpace" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetSpeedTest.ps1"; Description = "InternetSpeedTest" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InternetLatencyTest.ps1"; Description = "InternetLatencyTest" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Scripts/M365_Office64bitRepair.ps1"; Description = "M365_DeployOffice64bit" }
)

function InitiateScriptFromUrl {
    param (
        [string]$Url
    )
    try {
        $scriptContent = Invoke-WebRequest -Uri $Url -UseBasicParsing | Select-Object -ExpandProperty Content
        Invoke-Expression $scriptContent
    }
    catch {
        Write-Host "Failed to download or run script from URL: $Url" -ForegroundColor Red
    }
}

function Show-AutoByteMenu {
Clear-Host
Clear-Host
$separator = "================================================================"

Write-Host "`n"
Write-Host $separator -ForegroundColor Cyan
Write-Host "                        AutoByte Menu                           " -BackgroundColor Blue -ForegroundColor White
Write-Host $separator -ForegroundColor Cyan
Write-Host "`n"
Write-Host "   ______           __           ____             __" -ForegroundColor Cyan
Write-Host "  /\  _  \         /\ \__       /\  _ \          /\ \__" -ForegroundColor White
Write-Host "  \ \ \L\ \  __  __\ \  _\   __ \ \ \_\ \      __\ \  _\   __" -ForegroundColor Red
Write-Host "   \ \  __ \/\ \/\ \\ \ \/  / __ \ \  _ < /\ \/\ \\ \ \/  / __ \" -ForegroundColor Red
Write-Host "    \ \ \/\ \ \ \_\ \\ \ \_/\ \L\ \ \ \L\ \ \ \_\ \\ \ \_/\  __/" -ForegroundColor White
Write-Host "     \ \_\ \_\ \____/ \ \__\ \____/\ \____/\/ ____ \\ \__\ \____\" -ForegroundColor Cyan
Write-Host "      \/_/\/_/\/___/   \/__/\/___/  \/___/   /___/> \\/__/\/____/" -ForegroundColor Cyan
Write-Host "                                               /\___/" -ForegroundColor White
Write-Host "                                               \/__/" -ForegroundColor Red
Write-Host "`n"
Write-Host "Choose a script to run from the following options:" -BackgroundColor Blue -ForegroundColor White
Write-Host $separator -ForegroundColor Cyan
Write-Host ""


    $index = 1
    foreach ($script in $GitHubScripts) {
        Write-Host "$index. $($script.Description)" -ForegroundColor Green
        $index++
    }
    Write-Host "14. Exit" -ForegroundColor Red
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan

    $choice = Read-Host "Enter your choice (1-14)"  
    switch ($choice) {
        "1" { InitiateScriptFromUrl $GitHubScripts[0].ScriptUrl; Show-ReturnMenu }
        "2" { InitiateScriptFromUrl $GitHubScripts[1].ScriptUrl; Show-ReturnMenu }
        "3" { InitiateScriptFromUrl $GitHubScripts[2].ScriptUrl; Show-ReturnMenu }
        "4" { InitiateScriptFromUrl $GitHubScripts[3].ScriptUrl; Show-ReturnMenu }
        "5" { InitiateScriptFromUrl $GitHubScripts[4].ScriptUrl; Show-ReturnMenu }
        "6" { InitiateScriptFromUrl $GitHubScripts[5].ScriptUrl; Show-ReturnMenu }
        "7" { InitiateScriptFromUrl $GitHubScripts[6].ScriptUrl; Show-ReturnMenu }
        "8" { InitiateScriptFromUrl $GitHubScripts[7].ScriptUrl; Show-ReturnMenu }
        "9" { InitiateScriptFromUrl $GitHubScripts[8].ScriptUrl; Show-ReturnMenu }
        "10" { InitiateScriptFromUrl $GitHubScripts[9].ScriptUrl; Show-ReturnMenu }
        "11" { InitiateScriptFromUrl $GitHubScripts[10].ScriptUrl; Show-ReturnMenu }
        "12" { InitiateScriptFromUrl $GitHubScripts[11].ScriptUrl; Show-ReturnMenu }
        "13" { InitiateScriptFromUrl $GitHubScripts[12].ScriptUrl; Show-ReturnMenu }
        "14" {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
            return
        }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-AutoByteMenu
        }
    }
}

function Show-ReturnMenu {
    $returnChoice = Read-Host "`nReturn to the menu or exit? (Enter [1] for 'Yes' or [2] for 'exit')"
    switch ($returnChoice.ToLower()) {
        "1" { Show-AutoByteMenu }
        "2" {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
            return
        }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-ReturnMenu
        }
    }
}


Show-AutoByteMenu

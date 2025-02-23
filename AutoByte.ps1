$GitHubScripts = @(
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DiskCleaner.ps1"; Description = "DiskCleaner" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/EnableFilesOnDemand.ps1"; Description = "EnableFilesOnDemand" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DownloadFile.ps1"; Description = "DownloadFile" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/NetworkScan.ps1"; Description = "NetworkScan" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/BloatWareRemover.ps1"; Description = "DellBloatWareRemover" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallWindowsUpdate.ps1"; Description = "ResetandInstallWindowsUpdate" },
    #@{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/DeployPrinter.ps1"; Description = "DeployPrinter" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/ResetandClearWindowsSearchDB.ps1"; Description = "ResetandClearWindowsSearchdb" },
    #@{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckIfOneDriveSyncFolder.ps1"; Description = "CheckIfOneDriveSyncFolder" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckDriveSpace.ps1"; Description = "CheckDriveSpace" },
    #@{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/CheckIfUserIsAdmin.ps1"; Description = "CheckIfUserIsAdmin" },
    #@{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/Get-CurrentLoggedInUser.ps1"; Description = "Get-CurrentLoggedInUser" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/main/Scripts/InstallTeams.ps1"; Description = "InstallTeams" }
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
    Write-Host "10. Exit" -ForegroundColor Red
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan

    $choice = Read-Host "Enter your choice (1-10)"  
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
        "10" {
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

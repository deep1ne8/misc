$GitHubScripts = @(
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/DiskCleaner.ps1"; Description = "DiskCleaner" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/EnableFilesOnDemand.ps1"; Description = "EnableFilesOnDemand" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/CheckIfSystemNeedsToRestart.ps1"; Description = "CheckIfSystemNeedsToRestart" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/NetworkScan.ps1"; Description = "NetworkScan" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/RemoveWindowsUpdatePolicy.ps1"; Description = "RemoveWindowsUpdatePolicy" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/UpgradeWindowsToLatestFeature.ps1"; Description = "UpgradeWindowsToLatestFeature" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/DeployPrinter.ps1"; Description = "DeployPrinter" }
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

function Show-ScriptStarZMenu {
    Clear-Host
    Write-Host "================================================================"
    Write-Host "================================================================"
    Write-Host "==========  ____            _       _   ____  _                 " -ForegroundColor Yellow
    Write-Host "========== / ___|  ___ _ __(_)_ __ | |_/ ___|| |_ __ _ _ __     " -ForegroundColor Yellow
    Write-Host "========== \___ \ / __| '__| | '_ \| __\___ \| __/ _` | '__|    " -ForegroundColor Yellow
    Write-Host "==========  ___) | (__| |  | | |_) | |_ ___) | || (_| | |       " -ForegroundColor Yellow
    Write-Host "========== |____/ \___|_|  |_| .__/ \__|____/ \__\__,_|_|       " -ForegroundColor Yellow
    Write-Host "==========                   |_|                                " -ForegroundColor Yellow
    Write-Host "================================================================"
    Write-Host "================================================================"
    Write-Host "========== PowerShell scripts for system maintenance:           " -ForegroundColor Green
    Write-Host "----------------------------------------------------------------"
    $index = 1
    foreach ($script in $GitHubScripts) {
        Write-Host "$index. $($script.Description)" -ForegroundColor Green
        $index++
    }
    Write-Host "8. Exit" -ForegroundColor Red
    Write-Host ""
    Write-Host "==============================" -ForegroundColor Cyan

    $choice = Read-Host "Enter your choice (1-8)"
    switch ($choice) {
        "1" { InitiateScriptFromUrl $GitHubScripts[0].ScriptUrl; Show-ReturnMenu }
        "2" { InitiateScriptFromUrl $GitHubScripts[1].ScriptUrl; Show-ReturnMenu }
        "3" { InitiateScriptFromUrl $GitHubScripts[2].ScriptUrl; Show-ReturnMenu }
        "4" { InitiateScriptFromUrl $GitHubScripts[3].ScriptUrl; Show-ReturnMenu }
        "5" { InitiateScriptFromUrl $GitHubScripts[4].ScriptUrl; Show-ReturnMenu }
        "6" { InitiateScriptFromUrl $GitHubScripts[5].ScriptUrl; Show-ReturnMenu }
        "7" { InitiateScriptFromUrl $GitHubScripts[6].ScriptUrl; Show-ReturnMenu }
        "8" {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
            return
        }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-ScriptStarZMenu
        }
    }
}

function Show-ReturnMenu {
    $returnChoice = Read-Host "`nWould you like to return to the menu or exit? (Enter 'Yes' or 'exit')"
    switch ($returnChoice.ToLower()) {
        "Yes" { Show-ScriptStarZMenu }
        "exit" {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
            return
        }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-ReturnMenu
        }
    }
}

Show-ScriptStarZMenu

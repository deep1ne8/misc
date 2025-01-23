$GitHubScripts = @(
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/DiskCleaner.ps1"; Description = "DiskCleaner" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/RemoveAllTeams.ps1"; Description = "RemoveAllTeams" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/CheckIfSystemNeedsToRestart.ps1"; Description = "CheckIfSystemNeedsToRestart" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/NetworkScan.ps1"; Description = "NetworkScan" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/RemoveWindowsUpdatePolicy.ps1"; Description = "RemoveWindowsUpdatePolicy" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/UpgradeWindowsToLatestFeature.ps1"; Description = "UpgradeWindowsToLatestFeature" },
    @{ ScriptUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/DeployPrinter.ps1"; Description = "DeployPrinter" }
)

function Run-ScriptFromUrl {
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
    cls
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "       ScriptStar             " -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Cyan
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
        "1" { Run-ScriptFromUrl $GitHubScripts[0].ScriptUrl; Show-ReturnMenu }
        "2" { Run-ScriptFromUrl $GitHubScripts[1].ScriptUrl; Show-ReturnMenu }
        "3" { Run-ScriptFromUrl $GitHubScripts[2].ScriptUrl; Show-ReturnMenu }
        "4" { Run-ScriptFromUrl $GitHubScripts[3].ScriptUrl; Show-ReturnMenu }
        "5" { Run-ScriptFromUrl $GitHubScripts[4].ScriptUrl; Show-ReturnMenu }
        "6" { Run-ScriptFromUrl $GitHubScripts[5].ScriptUrl; Show-ReturnMenu }
        "7" { Run-ScriptFromUrl $GitHubScripts[6].ScriptUrl; Show-ReturnMenu }
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

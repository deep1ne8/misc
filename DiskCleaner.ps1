<#
	This PowerShell script automates disk cleanup and system optimization with a menu-driven interface. 
	It offers two modes:

		Dry-Run Mode: Simulates cleanup actions without making changes.
		Normal Mode: Executes the cleanup process.
	
	Features:
		Menu Interface: User-friendly options to select the mode or exit.
		Targeted Cleanup: Removes temporary files, system caches, update files, crash dumps, and old logs.
		Additional Tasks: Clears Recycle Bin, DNS cache, and performs a DISM cleanup.
		Disk Space Reporting: Displays before-and-after disk space and space recovered.
		Error Handling: Provides detailed error messages.

	Author: EDaniels
	Date: 12/2024
#>

function Start-AdvancedSystemCleanup {
    param (
        [int]$DaysOld = 30,
        [int]$LargeFileSizeGB = 1,
        [switch]$DryRun
    )

    Write-Host "Starting Advanced System Cleanup..." -ForegroundColor Green

    # Get initial disk space
    $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $initialFreeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor Cyan

    # Define cleanup locations with descriptions
    $LoggedInUser = quser | ForEach-Object { ($_ -split '\s{2,}')[0].TrimStart('>', ' ') } | Where-Object { $_ -notmatch 'USERNAME' }
    $LocalAppDataTempFolder = "$env:SystemDrive\Users\$LoggedInUser\appdata\local\Temp"
    $LocalSoftwareDistributionFolder = "$env:SystemRoot\SoftwareDistribution"
    $WindowsExplorerCacheFolder = "$env:SystemDrive\Users\$LoggedInUser\appdata\local\Microsoft\Windows\Explorer"
    $LocalAppCrashDumpsFolder = "$env:SystemDrive\Users\$LoggedInUser\appdata\local\CrashDumps"
    $cleanupPaths = @(
        @{ Path = "$env:TEMP"; Description = "Temporary Files" },
        @{ Path = "$env:SystemRoot\Temp"; Description = "Windows Temp" },
        @{ Path = "$LocalAppDataTempFolder"; Description = "Local App Temp" },
        @{ Path = "$LocalSoftwareDistributionFolder"; Description = "Windows Update Cache" },
        @{ Path = "$WindowsExplorerCacheFolder"; Description = "Explorer Cache" },
        @{ Path = "$env:SystemRoot\Logs"; Description = "Windows Logs" },
        @{ Path = "$LocalAppCrashDumpsFolder"; Description = "Crash Dumps" }
    )

    # Track space cleaned
    $totalSpaceCleaned = 0

    foreach ($item in $cleanupPaths) {
        $path = $item.Path
        $description = $item.Description

        if (Test-Path $path) {
            $beforeSize = (Get-ChildItem $path -Recurse -Verbose -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1GB
            Write-Host "`nProcessing $description at $path..." -ForegroundColor Yellow
            
            if ($DryRun) {
                Write-Host "Dry run: Would clean files older than $DaysOld days or larger than $LargeFileSizeGB GB" -ForegroundColor Cyan
            } else {
                try {
                    Get-ChildItem -Path $path -Recurse -File -Verbose -ErrorAction SilentlyContinue |
                        Where-Object { 
                            $_.LastWriteTime -lt (Get-Date).AddDays(-$DaysOld) -or 
                            $_.Length -gt ($LargeFileSizeGB * 1GB)
                        } |
                        Remove-Item -Force -Verbose -ErrorAction SilentlyContinue

                    $afterSize = (Get-ChildItem $path -Recurse -Verbose -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1GB
                    $spaceCleaned = [math]::Round($beforeSize - $afterSize, 2)
                    $totalSpaceCleaned += $spaceCleaned
                    Write-Host "Cleaned $spaceCleaned GB from $description" -ForegroundColor Green
                } catch {
				Write-Host "Error cleaning ${description}: $_" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Path not found: $path" -ForegroundColor DarkYellow
        }
    }

    # Get final disk space
    $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $finalFreeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)

    # Summary
    Write-Host "`nCleanup Summary:" -ForegroundColor Cyan
    Write-Host "Hostname: $(hostname)" -ForegroundColor Yellow
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor White
    Write-Host "Final free space: $finalFreeSpace GB" -ForegroundColor White
    Write-Host "Total space recovered: $([math]::Round($finalFreeSpace - $initialFreeSpace, 2)) GB" -ForegroundColor Green
    Write-Host "Detailed cleanup completed: $([math]::Round($totalSpaceCleaned, 2)) GB" -ForegroundColor Green

    if (-not $DryRun) {
        Write-Host "`nRecommendation: Please restart your computer to complete the cleanup process." -ForegroundColor Yellow
    }
}

# Menu for Dry-Run
function Show-CleanupMenu {
    cls
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "  Advanced System Cleanup Menu" -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "1. Run Cleanup in Dry-Run Mode"
    Write-Host "2. Run Cleanup Normally"
    Write-Host "3. Exit"
    Write-Host ""

    $choice = Read-Host "Enter your choice (1-3)"
    switch ($choice) {
        "1" {
            Write-Host "`nRunning in Dry-Run mode..." -ForegroundColor Green
            Start-AdvancedSystemCleanup -DryRun
        }
        "2" {
            Write-Host "`nRunning cleanup normally..." -ForegroundColor Green
            Start-AdvancedSystemCleanup
        }
        "3" {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
            return
        }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-CleanupMenu
        }
    }
}

# Show the menu
Show-CleanupMenu

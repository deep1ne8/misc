<#
	This PowerShell script automates disk cleanup and system optimization with a menu-driven interface. 
	It offers three modes:

		Dry-Run Mode: Simulates cleanup actions without making changes.
		Normal Mode: Executes the cleanup process.
  		Large File Scanner: Scans for files larger then 1GB on the C: drive
    		User Profile Scanner
      		Dell bloatware remover
	
	Features:
		Menu Interface: User-friendly options to select the mode or exit.
		Targeted Cleanup: Removes temporary files, system caches, update files, crash dumps, and old logs.
		Additional Tasks: Clears Recycle Bin, DNS cache, and performs a DISM cleanup.
		Disk Space Reporting: Displays before-and-after disk space and space recovered.
		Error Handling: Provides detailed error messages.
  		Large Files Scanner: Scans for files larger then 1GB on the C: drive and outputs to terminal
    		Scans for user profiles older then 90 days, and their disk space consumed
      		Removes Dell Support Assist, Remediation and OS Recovery plugin bloatware

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
    #$LoggedInUser = quser | ForEach-Object { ($_ -split '\s{2,}')[0].TrimStart('>', ' ') } | Where-Object { $_ -notmatch 'USERNAME' }
    try {
    # Try using 'quser'
    $LoggedInUser = quser | ForEach-Object { ($_ -split '\s{2,}')[0].TrimStart('>', ' ') } | Where-Object { $_ -notmatch 'USERNAME' }
    	if (-not $LoggedInUser) {
        	throw "No users found with quser."
    	}
    } catch {
    # If 'quser' fails, fall back to 'Get-WmiObject'
    $LoggedInUser = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
    	if (-not $LoggedInUser) {
        	throw "Failed to retrieve logged-in user information."
    	}
    }

   # Output the result
   	if ($LoggedInUser) {
    Write-Host "Logged-in User(s): $LoggedInUser"
	} else {
    Write-Host "No logged-in users found."
    }
    
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
    Write-Host "Logged-in User(s): $LoggedInUser" -ForegroundColor White
    Write-Host "Hostname: $(hostname)" -ForegroundColor Yellow
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor White
    Write-Host "Final free space: $finalFreeSpace GB" -ForegroundColor White
    Write-Host "Total space recovered: $([math]::Round($finalFreeSpace - $initialFreeSpace, 2)) GB" -ForegroundColor Green
    Write-Host "Detailed cleanup completed: $([math]::Round($totalSpaceCleaned, 2)) GB" -ForegroundColor Green

    if (-not $DryRun) {
        Write-Host "`nRecommendation: Please restart your computer to complete the cleanup process." -ForegroundColor Yellow
    	}
    }


    function LargeFiles {
    param (
        [string]$SourcePath = "C:\",               # Source directory to scan
        [string]$DestinationPath = "D:\Temp",      # Destination directory (doesn't matter, we use /L for list mode)
        [int64]$SizeThresholdBytes = 1073741824    # Size threshold in bytes (1GB)
    )

    try {
        Write-Host "`n"
        Write-Host "Starting scan for files larger than $([math]::Round($SizeThresholdBytes / 1GB, 2)) GB in path: $SourcePath" -ForegroundColor Cyan
        Write-Host "`n"
        
        # Run robocopy with /L to list files larger than the threshold
        $robocopyCommand = "robocopy $SourcePath $DestinationPath /L /MIN:$SizeThresholdBytes"
        $robocopyOutput = & cmd.exe /c $robocopyCommand

        # Parse robocopy output to extract file paths and sizes
        $robocopyOutput | ForEach-Object {
            # Match lines that show "New File" with the file size (size will be the 3rd part)
            if ($_ -match "\s+New\s+File\s+([0-9\.]+)\s+(g|m)\s+(.+)$") {
                $fileSize = [double]$matches[1]
                $fileUnit = $matches[2]
                $filePath = $matches[3].Trim()

                # Convert size to bytes (if in GB, convert to bytes)
                if ($fileUnit -eq "g") {
                    $fileSizeBytes = $fileSize * 1GB
                }
                elseif ($fileUnit -eq "m") {
                    $fileSizeBytes = $fileSize * 1MB
                }

                if ($fileSizeBytes -gt $SizeThresholdBytes) {
                    # Display the full file path
                    $fullPath = Join-Path -Path $SourcePath -ChildPath $filePath
                    Write-Host "Large File Found: $fullPath - Size: $([math]::Round($fileSizeBytes / 1GB, 2)) GB" -ForegroundColor Yellow
                    Write-Host "`n"
                }
            }
        }
        
        Write-Host "Scan completed." -ForegroundColor Green
        Write-Host "`n"
    }
    catch {
        Write-Host "Error occurred: $_" -ForegroundColor Red
        Write-Host "`n"
        Return
    }
}

function ListUserProfiles {
    # Parameters
    $usersPath = "C:\Users"
    $daysOld = 90
    $totalSpace = 0

    # Exclude specific directories
    $excludeProfiles = @("Public", "TEMP", "defaultuser1", "All Users", "default", "Default User", "DefaultAppPool", "HvmService")

    # Progress bar setup
    $userProfiles = Get-ChildItem -Path $usersPath -Directory | Where-Object {
        ($_.Attributes -notmatch 'Hidden|System') -and
        $_.Name -notin $excludeProfiles
    }
    $totalProfiles = $userProfiles.Count
    $currentProfile = 0

    # Create an array to store results
    $results = @()

    # Enumerate user profiles
    foreach ($profile in $userProfiles) {
        $currentProfile++
        Write-Progress -Activity "Scanning Profiles" -Status "$($currentProfile)/$($totalProfiles) profiles scanned" -PercentComplete (($currentProfile / $totalProfiles) * 100)
        
        try {
            $lastWriteTime = (Get-Item $profile.FullName).LastWriteTime
            $profileSize = (Get-ChildItem -Path $profile.FullName -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum

            # Check if profile is older than 90 days
            if ((Get-Date).AddDays(-$daysOld) -gt $lastWriteTime -and $profileSize -gt 0) {
                $profileSizeGB = [math]::Round($profileSize / 1GB, 2)
                $results += [PSCustomObject]@{
                    ProfileName = $profile.Name
                    LastUsed    = $lastWriteTime
                    SpaceUsedGB = $profileSizeGB
                }
                $totalSpace += $profileSize
            } else {
                Write-Host "Skipping $($profile.Name): Recently used or empty" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error processing $($profile.FullName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Display results
    if ($results.Count -gt 0) {
        Write-Host "`nProfiles older than $daysOld days:" -ForegroundColor Cyan
        $results | Format-Table -AutoSize
    } else {
        Write-Host "`nNo profiles older than $daysOld days found." -ForegroundColor Yellow
    }

    # Convert total space to GB and display
    $totalSpaceGB = [math]::Round($totalSpace / 1GB, 2)
    Write-Host "`nTotal Space Used by Profiles Older Than $daysOld Days: $totalSpaceGB GB" -ForegroundColor Green
}

function CheckAndUninstallDellApps {
    # Check if the computer is a Dell system
    $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer

    if ($manufacturer -like "*Dell*") {
        Write-Host "This is a Dell system. Proceeding with uninstallation of Dell apps..." -ForegroundColor Green

        try {
            # Run the Dell uninstallation script
            Invoke-Expression (Invoke-WebRequest -Uri "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Uninstall-DellApps.ps1" -UseBasicParsing).Content
            Write-Host "Uninstallation script executed successfully." -ForegroundColor Green
        } catch {
            Write-Host "Failed to execute the uninstallation script: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "This is not a Dell system. No action required." -ForegroundColor Yellow
    }
}


# Menu for Dry-Run
function Show-CleanupMenu {
    Clear-Host
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "  Advanced System Cleanup Menu" -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "1. Run Cleanup in Dry-Run Mode"
    Write-Host "2. Run Cleanup Normally"
    Write-Host "3. Run Large File Scanner"
    Write-Host "4. Run Large User Profiles scanner"
    Write-Host "5. Run Dell BloatWare Remover"
    Write-Host "6. Exit"
    Write-Host ""

    $choice = Read-Host "Enter your choice (1-6)"
    switch ($choice) {
        "1" {
            Write-Host "`nRunning in Dry-Run mode..." -ForegroundColor Green
            Start-AdvancedSystemCleanup -DryRun
            Show-ReturnMenu
        }
        "2" {
            Write-Host "`nRunning cleanup normally..." -ForegroundColor Green
            Start-AdvancedSystemCleanup
            Show-ReturnMenu
        }
        "3" {
            Write-Host "`nRunning large file scanner..." -ForegroundColor Yellow
            LargeFiles -SourcePath "C:\" -DestinationPath "D:\Temp"
            Show-ReturnMenu
        }
        "4" {
            Write-Host "`nRunning large user profile scanner..." -ForegroundColor Yellow
            ListUserProfiles
	        Show-ReturnMenu
        }
	    "5" {
 	        Write-Host "`Running Dell bloatware remover..." -ForegroundColor Yellow
      	    CheckAndUninstallDellApps
	        Show-ReturnMenu
	    }
 	    "6" {
 	    Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
	    return
	    }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-ReturnMenu
        }
    }
}

function Show-ReturnMenu {
    $returnChoice = Read-Host "`nWould you like to return to the menu or exit? (Enter 'menu' or 'exit')"
    switch ($returnChoice.ToLower()) {
        "menu" {
            Show-CleanupMenu
        }
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

Show-CleanupMenu

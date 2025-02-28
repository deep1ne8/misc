function Start-AdvancedSystemCleanup {
    param (
        [switch]$DryRun
    )

    $DaysOld = 30
    $LargeFileSizeGB = 1
    Write-Host "Starting Advanced System Cleanup..." -ForegroundColor Cyan

    # Get initial disk space
    $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    if (!$drive) {
        Write-Host "Failed to retrieve drive information." -ForegroundColor Red
        return
    }
    $initialFreeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor Green

    # Define cleanup locations with descriptions
    $cleanupPaths = @(
        @{ Path = "$env:TEMP"; Description = "Temporary Files" },
        @{ Path = "$env:SystemRoot\Temp"; Description = "Windows Temp" },
        @{ Path = "$env:SystemRoot\Logs"; Description = "Windows Logs" },
        @{ Path = "$env:SystemRoot\SoftwareDistribution"; Description = "Windows Update Cache" },
        @{ Path = "$env:SystemRoot\CrashDumps"; Description = "Crash Dumps" }
    )

    # Track space cleaned
    $totalSpaceCleaned = 0

    foreach ($item in $cleanupPaths) {
        $path = $item.Path
        $description = $item.Description

        if (Test-Path $path) {
            $beforeSize = (Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1GB
            Write-Host "Processing $description at $path..." -ForegroundColor Yellow

            if ($DryRun) {
                Write-Host "Dry run: Would clean files older than $DaysOld days or larger than $LargeFileSizeGB GB" -ForegroundColor Yellow
            } else {
                try {
                    Get-ChildItem -Path $path -Recurse -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$DaysOld) -or $_.Length -gt ($LargeFileSizeGB * 1GB) } |
                        Remove-Item -Force -ErrorAction SilentlyContinue

                    $afterSize = (Get-ChildItem $path -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1GB
                    $spaceCleaned = [math]::Round($beforeSize - $afterSize, 2)
                    $totalSpaceCleaned += $spaceCleaned
                    Write-Host "Cleaned $spaceCleaned GB from $description" -ForegroundColor Green
                } catch {
                    Write-Host "Error cleaning ${description}: $_" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "Path not found: $path" -ForegroundColor Magenta
        }
    }

    # Get final disk space
    $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $finalFreeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)

    # Summary
    Write-Host "Cleanup Summary:" -ForegroundColor Cyan
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor Green
    Write-Host "Final free space: $finalFreeSpace GB" -ForegroundColor Green
    Write-Host "Total space recovered: $([math]::Round($finalFreeSpace - $initialFreeSpace, 2)) GB" -ForegroundColor Green
    Write-Host "Detailed cleanup completed: $([math]::Round($totalSpaceCleaned, 2)) GB" -ForegroundColor Green

    if (-not $DryRun) {
        Write-Host "Recommendation: Please restart your computer to complete the cleanup process." -ForegroundColor Yellow
    }
}

function Find-LargeFiles {
    param (
        [string]$SourcePath = "C:\",
        [string]$DestinationPath = "D:\Temp",
        [double]$SizeThresholdGB = 1
    )

    $SizeThresholdBytes = $SizeThresholdGB * 1GB

    try {
        Write-Host "Starting scan for files larger than $SizeThresholdGB GB in path: $SourcePath" -ForegroundColor Cyan
        
        # Run robocopy with /L to list files larger than the threshold
        $robocopyCommand = "robocopy $SourcePath $DestinationPath /L /MIN:$SizeThresholdBytes"
        $robocopyOutput = & cmd.exe /c $robocopyCommand

        # Create an array to store results for table output
        $largeFiles = @()

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
                    $largeFiles += [PSCustomObject]@{
                        Path = $fullPath
                        SizeGB = [math]::Round($fileSizeBytes / 1GB, 2)
                    }
                }
            }
        }
        
        # Display results as a table
        if ($largeFiles.Count -gt 0) {
            $largeFiles | Sort-Object -Property SizeGB -Descending | Format-Table -AutoSize
        } else {
            Write-Host "No files larger than $SizeThresholdGB GB found." -ForegroundColor Green
        }
        
        Write-Host "Scan completed." -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error occurred: $_" -ForegroundColor Red
        Return
    }
}

function Get-OldUserProfiles {
    param (
        [int]$DaysOld = 90,
        [string]$UsersPath = "C:\Users"
    )
    
    # Exclude specific directories
    $excludeProfiles = @("Public", "TEMP", "defaultuser1", "All Users", "default", "Default User", "DefaultAppPool", "HvmService")

    # Progress bar setup
    $userProfiles = Get-ChildItem -Path $UsersPath -Directory | Where-Object {
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

            # Check if profile is older than specified days
            if ((Get-Date).AddDays(-$DaysOld) -gt $lastWriteTime -and $profileSize -gt 0) {
                $profileSizeGB = [math]::Round($profileSize / 1GB, 2)
                $daysSinceLastUse = [math]::Round(((Get-Date) - $lastWriteTime).TotalDays, 0)
                
                $results += [PSCustomObject]@{
                    ProfileName = $profile.Name
                    LastUsed    = $lastWriteTime
                    DaysInactive = $daysSinceLastUse
                    SpaceUsedGB = $profileSizeGB
                }
            } else {
                Write-Host "Skipping $($profile.Name): Recently used or empty" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error processing $($profile.FullName): $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Display results
    if ($results.Count -gt 0) {
        Write-Host "Profiles older than $DaysOld days:" -ForegroundColor Cyan
        $results | Sort-Object -Property SpaceUsedGB -Descending | Format-Table -AutoSize
        
        # Calculate and display total space used by old profiles
        $totalSpaceUsed = ($results | Measure-Object -Property SpaceUsedGB -Sum).Sum
        Write-Host "Total space used by old profiles: $([math]::Round($totalSpaceUsed, 2)) GB" -ForegroundColor Green
    } else {
        Write-Host "No profiles older than $DaysOld days found." -ForegroundColor Green
    }
}

function Uninstall-DellApps {
    # Check if the computer is a Dell system
    $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer

    if ($manufacturer -like "*Dell*") {
        Write-Host "This is a Dell system. Proceeding with uninstallation of Dell apps..." -ForegroundColor Cyan

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

function Show-CleanupMenu {
    # Measure script execution time
    $functionTimer = [System.Diagnostics.Stopwatch]::StartNew()
    
    Clear-Host
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "  Advanced System Cleanup Menu" -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "1. Run Advanced System Cleanup" -ForegroundColor Green
    Write-Host "2. Run Cleanup Dry Run" -ForegroundColor Green
    Write-Host "3. List User Profiles" -ForegroundColor Green
    Write-Host "4. List Large Files" -ForegroundColor Green
    Write-Host "5. Clean Dell Bloatware" -ForegroundColor Green
    Write-Host "6. Exit" -ForegroundColor Green
    Write-Host "==============================" -ForegroundColor Cyan

    try {
        [int]$choice = Read-Host "Enter your choice (1-6)"
        if ($choice -eq 0) {
            throw "Choice cannot be empty."
        }

        # Reset timer when function is actually called
        $functionTimer.Restart()
        
        switch ($choice) {
            1 { Start-AdvancedSystemCleanup }
            2 { Start-AdvancedSystemCleanup -DryRun }
            3 { Get-OldUserProfiles }
            4 { Find-LargeFiles }
            5 { Uninstall-DellApps }
            6 {
                Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
                return
            }
            default {
                Write-Host "Invalid choice, please try again." -ForegroundColor Red
                Show-CleanupMenu
                return
            }
        }
        
        # Stop the timer and display execution time
        $functionTimer.Stop()
        $executionTime = $functionTimer.Elapsed
        Write-Host "`nExecution Time: $($executionTime.Hours.ToString('00')):$($executionTime.Minutes.ToString('00')):$($executionTime.Seconds.ToString('00')).$($executionTime.Milliseconds.ToString('000'))" -ForegroundColor Cyan
        
        Show-ReturnMenu
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Show-CleanupMenu
    }
}

function Show-ReturnMenu {
    try {
        [int]$returnChoice = Read-Host "`nReturn to the menu or exit? (Enter [1] for 'Yes' or [2] for 'exit')"
        if ($returnChoice -eq 0) {
            throw "Choice cannot be empty."
        }

        switch ($returnChoice) {
            1 { Show-CleanupMenu }
            2 {
                Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
                return
            }
            default {
                Write-Host "Invalid choice, please try again." -ForegroundColor Red
                Show-ReturnMenu
            }
        }
    } catch {
        Write-Host "An error occurred: $_" -ForegroundColor Red
        Show-ReturnMenu
    }
}

# Display the cleanup menu
Show-CleanupMenu
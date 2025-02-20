function Start-AdvancedSystemCleanup {
    $DaysOld = 30
    $LargeFileSizeGB = 1
    Write-Host "Starting Advanced System Cleanup..." -ForegroundColor Cyan

    # Get initial disk space
    $drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
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

function LargeFiles {

    $SourcePath = "C:\"
    $DestinationPath = "D:\Temp"
    $SizeThresholdBytes = 1073741824

    try {
        Write-Host "Starting scan for files larger than $([math]::Round($SizeThresholdBytes / 1GB, 2)) GB in path: $SourcePath" -ForegroundColor Cyan
        
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
                }
            }
        }
        
        Write-Host "Scan completed." -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error occurred: $_" -ForegroundColor Red
        Return
    }
}

function ListUserProfiles {

    $DaysOld = 90
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
            if ((Get-Date).AddDays(-$DaysOld) -gt $lastWriteTime -and $profileSize -gt 0) {
                $profileSizeGB = [math]::Round($profileSize / 1GB, 2)
                $results += [PSCustomObject]@{
                    ProfileName = $profile.Name
                    LastUsed    = $lastWriteTime
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
        $results | Format-Table -AutoSize
    } else {
        Write-Host "No profiles older than $DaysOld days found." -ForegroundColor Green
    }
}

function CheckAndUninstallDellApps {
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
Clear-Host
function Show-CleanupMenu {
    Clear-Host
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "  Advanced System Cleanup Menu" -ForegroundColor Yellow
    Write-Host "==============================" -ForegroundColor Cyan
    Write-Host "1. Run Advanced System Cleanup" -ForegroundColor Green
    Write-Host "2. Run Cleanup Dry Run" -ForegroundColor Green
    Write-Host "2. Run Cleanup Normally" -ForegroundColor Green
    Write-Host "2. List User Profiles" -ForegroundColor Green
    Write-Host "3. List Large Files" -ForegroundColor Green
    Write-Host "4. Exit" -ForegroundColor Green
    Write-Host "==============================" -ForegroundColor Cyan

    $choice = Read-Host "Enter your choice (1-6)"

    switch ($choice) {
        "1" { Start-AdvancedSystemCleanup }
        "2" { Run-Cleanup -DryRun }
        "3" { Run-Cleanup }
        "4" { ListUserProfiles }
        "5" { LargeFiles }
        "6" {
            Write-Host "Exiting. Goodbye!" -ForegroundColor Yellow
            return
        }
        default {
            Write-Host "Invalid choice, please try again." -ForegroundColor Red
            Show-CleanupMenu
        }
    }
}

function Show-ReturnMenu {
    $returnChoice = Read-Host "`nReturn to the menu or exit? (Enter [1] for 'Yes' or [2] for 'exit')"
    switch ($returnChoice.ToLower()) {
        "1" { Show-CleanupMenu }
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
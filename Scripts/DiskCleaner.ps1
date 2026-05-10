#Requires -Version 5.1
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Advanced System Cleanup Script with enhanced performance and error handling
.DESCRIPTION
    Optimized PowerShell script for comprehensive system cleanup with improved efficiency,
    better error handling, and enhanced user experience.
.NOTES
    Version: 2.0
    Requires: PowerShell 5.1+, Administrator privileges
#>

[CmdletBinding()]
param(
    [ValidateSet('Interactive', 'Silent', 'DryRun')]
    [string]$Mode = 'Interactive',
    [int]$DaysOld = 30,
    [double]$LargeFileSizeGB = 1.0,
    [string]$LogPath = "$env:TEMP\SystemCleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
)

# Initialize logging
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $logEntry -Force
    
    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARN"  { Write-Host $Message -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message -ForegroundColor White }
    }
}

function Get-DiskSpace {
    param([string]$Drive = "C:")
    try {
        $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$Drive'" -ErrorAction Stop
        return [PSCustomObject]@{
            TotalGB = [math]::Round($disk.Size / 1GB, 2)
            FreeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
            UsedGB = [math]::Round(($disk.Size - $disk.FreeSpace) / 1GB, 2)
            FreePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
        }
    }
    catch {
        Write-Log "Failed to retrieve disk information for $Drive : $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Remove-ItemSafely {
    param(
        [string]$Path,
        [switch]$Recurse,
        [switch]$Force = $true
    )
    
    try {
        if (-not (Test-Path $Path)) { return $false }
        
        # Handle long paths
        if ($Path.Length -gt 260) {
            $longPath = "\\?\$($Path -replace '^\\\\', '\\')"
            if (Test-Path $longPath) {
                $item = Get-Item $longPath -Force -ErrorAction Stop
                if ($item.PSIsContainer) {
                    [System.IO.Directory]::Delete($longPath, $true)
                } else {
                    [System.IO.File]::Delete($longPath)
                }
                return $true
            }
        }
        
        # Standard removal
        $params = @{
            Path = $Path
            Force = $Force
            ErrorAction = 'Stop'
        }
        if ($Recurse) { $params.Recurse = $true }
        
        Remove-Item @params
        return $true
    }
    catch {
        Write-Log "Failed to remove $Path : $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Get-FolderSize {
    param([string]$Path)
    
    if (-not (Test-Path $Path)) { return 0 }
    
    try {
        $size = Get-ChildItem -Path $Path -Recurse -File -Force -ErrorAction SilentlyContinue |
                Measure-Object -Property Length -Sum |
                Select-Object -ExpandProperty Sum
        return [math]::Max($size, 0)
    }
    catch {
        Write-Log "Failed to calculate size for $Path : $($_.Exception.Message)" "WARN"
        return 0
    }
}

function Start-AdvancedSystemCleanup {
    param([switch]$DryRun)
    
    Write-Log "Starting Advanced System Cleanup (DryRun: $($DryRun.IsPresent))" "INFO"
    
    # Get initial disk space
    $initialDisk = Get-DiskSpace
    if (-not $initialDisk) { return }
    
    Write-Log "Initial free space: $($initialDisk.FreeGB) GB ($($initialDisk.FreePercent)%)" "INFO"

    # Define cleanup locations with validation
    $cleanupPaths = @(
        @{ Path = $env:TEMP; Description = "User Temp Files"; SkipRunning = $false },
        @{ Path = "$env:SystemRoot\Temp"; Description = "System Temp Files"; SkipRunning = $false },
        @{ Path = "$env:SystemRoot\Logs"; Description = "Windows Logs"; SkipRunning = $false },
        @{ Path = "$env:SystemRoot\SoftwareDistribution\Download"; Description = "Windows Update Cache"; SkipRunning = $true },
        @{ Path = "$env:SystemRoot\CrashDumps"; Description = "Crash Dumps"; SkipRunning = $false },
        @{ Path = "$env:LOCALAPPDATA\Microsoft\Windows\WebCache"; Description = "Web Cache"; SkipRunning = $true },
        @{ Path = "$env:LOCALAPPDATA\Temp"; Description = "Local App Data Temp"; SkipRunning = $false }
    ) | Where-Object { Test-Path $_.Path }

    $totalSpaceCleaned = 0
    $cleanupResults = @()

    foreach ($item in $cleanupPaths) {
        Write-Progress -Activity "System Cleanup" -Status "Processing $($item.Description)" -PercentComplete ((($cleanupPaths.IndexOf($item) + 1) / $cleanupPaths.Count) * 100)
        
        $beforeSize = Get-FolderSize -Path $item.Path
        Write-Log "Processing $($item.Description) at $($item.Path) (Size: $([math]::Round($beforeSize / 1GB, 2)) GB)" "INFO"

        if ($DryRun) {
            $filesToClean = Get-ChildItem -Path $item.Path -Recurse -File -Force -ErrorAction SilentlyContinue |
                           Where-Object { 
                               $_.LastWriteTime -lt (Get-Date).AddDays(-$DaysOld) -or 
                               $_.Length -gt ($LargeFileSizeGB * 1GB) 
                           }
            
            $dryRunSize = ($filesToClean | Measure-Object -Property Length -Sum).Sum
            Write-Log "DryRun: Would clean $([math]::Round($dryRunSize / 1GB, 2)) GB from $($item.Description)" "INFO"
            continue
        }

        # Stop related services if needed
        if ($item.SkipRunning) {
            $services = @("wuauserv", "bits", "cryptsvc")
            $stoppedServices = @()
            
            foreach ($service in $services) {
                try {
                    $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
                    if ($svc -and $svc.Status -eq 'Running') {
                        Stop-Service -Name $service -Force -NoWait -ErrorAction Stop
                        $stoppedServices += $service
                        Write-Log "Stopped service: $service" "INFO"
                    }
                }
                catch {
                    Write-Log "Failed to stop service $service : $($_.Exception.Message)" "WARN"
                }
            }
            Start-Sleep -Seconds 2
        }

        try {
            # Enhanced file filtering with better performance
            $filesToRemove = Get-ChildItem -Path $item.Path -Recurse -File -Force -ErrorAction SilentlyContinue |
                            Where-Object { 
                                ($_.LastWriteTime -lt (Get-Date).AddDays(-$DaysOld)) -or 
                                ($_.Length -gt ($LargeFileSizeGB * 1GB)) 
                            }

            $removedCount = 0
            $removedSize = 0

            foreach ($file in $filesToRemove) {
                if (Remove-ItemSafely -Path $file.FullName) {
                    $removedCount++
                    $removedSize += $file.Length
                }
            }

            # Restart stopped services
            foreach ($service in $stoppedServices) {
                try {
                    Start-Service -Name $service -ErrorAction Stop
                    Write-Log "Restarted service: $service" "INFO"
                }
                catch {
                    Write-Log "Failed to restart service $service : $($_.Exception.Message)" "WARN"
                }
            }

            $afterSize = Get-FolderSize -Path $item.Path
            $spaceCleaned = [math]::Round(($beforeSize - $afterSize) / 1GB, 2)
            $totalSpaceCleaned += $spaceCleaned

            $cleanupResults += [PSCustomObject]@{
                Location = $item.Description
                FilesRemoved = $removedCount
                SpaceCleanedGB = $spaceCleaned
                Status = "Success"
            }

            Write-Log "Cleaned $spaceCleaned GB ($removedCount files) from $($item.Description)" "SUCCESS"
        }
        catch {
            $cleanupResults += [PSCustomObject]@{
                Location = $item.Description
                FilesRemoved = 0
                SpaceCleanedGB = 0
                Status = "Failed: $($_.Exception.Message)"
            }
            Write-Log "Error cleaning $($item.Description): $($_.Exception.Message)" "ERROR"
        }
    }

    Write-Progress -Activity "System Cleanup" -Completed

    # Final disk space check
    $finalDisk = Get-DiskSpace
    
    # Display comprehensive summary
    Write-Log "`n========== CLEANUP SUMMARY ==========" "INFO"
    Write-Log "Initial free space: $($initialDisk.FreeGB) GB ($($initialDisk.FreePercent)%)" "INFO"
    Write-Log "Final free space: $($finalDisk.FreeGB) GB ($($finalDisk.FreePercent)%)" "INFO"
    Write-Log "Space recovered: $([math]::Round($finalDisk.FreeGB - $initialDisk.FreeGB, 2)) GB" "SUCCESS"
    
    if ($cleanupResults.Count -gt 0) {
        Write-Log "`nDetailed Results:" "INFO"
        $cleanupResults | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Log $_ "INFO" }
    }

    if (-not $DryRun) {
        Write-Log "`nRecommendation: Restart your computer to complete the cleanup process." "WARN"
    }
}

function Find-LargeFiles {
    param(
        [string]$SourcePath = "C:\",
        [double]$SizeThresholdGB = 1.0,
        [int]$TopN = 50
    )

    Write-Log "Scanning for files larger than $SizeThresholdGB GB in: $SourcePath" "INFO"
    $sizeThresholdBytes = $SizeThresholdGB * 1GB
    $largeFiles = @()

    try {
        # Use more efficient approach with ForEach-Object for memory management
        Get-ChildItem -Path $SourcePath -Recurse -File -Force -ErrorAction SilentlyContinue |
        Where-Object { $_.Length -gt $sizeThresholdBytes } |
        Sort-Object -Property Length -Descending |
        Select-Object -First $TopN |
        ForEach-Object {
            $largeFiles += [PSCustomObject]@{
                Path = $_.FullName
                SizeGB = [math]::Round($_.Length / 1GB, 2)
                LastModified = $_.LastWriteTime
                Extension = $_.Extension
            }
        }

        if ($largeFiles.Count -gt 0) {
            Write-Log "Found $($largeFiles.Count) large files:" "SUCCESS"
            $largeFiles | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Log $_ "INFO" }
            
            $totalSize = ($largeFiles | Measure-Object -Property SizeGB -Sum).Sum
            Write-Log "Total size of large files: $([math]::Round($totalSize, 2)) GB" "INFO"
        } else {
            Write-Log "No files larger than $SizeThresholdGB GB found." "INFO"
        }
    }
    catch {
        Write-Log "Error during file scan: $($_.Exception.Message)" "ERROR"
    }
}

function Get-OldUserProfiles {
    param(
        [int]$DaysOld = 90,
        [string]$UsersPath = "C:\Users"
    )
    
    $excludeProfiles = @("Public", "TEMP", "defaultuser1", "All Users", "default", "Default User", "DefaultAppPool", "HvmService")
    
    Write-Log "Scanning for user profiles older than $DaysOld days..." "INFO"
    
    try {
        $userProfiles = Get-ChildItem -Path $UsersPath -Directory -Force -ErrorAction Stop |
                       Where-Object { 
                           $_.Name -notin $excludeProfiles -and
                           ($_.Attributes -band [System.IO.FileAttributes]::Hidden) -eq 0
                       }

        $results = @()
        $currentProfile = 0
        $totalProfiles = $userProfiles.Count

        foreach ($profile in $userProfiles) {
            $currentProfile++
            Write-Progress -Activity "Scanning User Profiles" -Status "Processing $($profile.Name)" -PercentComplete (($currentProfile / $totalProfiles) * 100)
            
            try {
                $lastWriteTime = $profile.LastWriteTime
                $daysSinceLastUse = [math]::Round(((Get-Date) - $lastWriteTime).TotalDays, 0)
                
                if ($daysSinceLastUse -gt $DaysOld) {
                    $profileSize = Get-FolderSize -Path $profile.FullName
                    
                    if ($profileSize -gt 0) {
                        $results += [PSCustomObject]@{
                            ProfileName = $profile.Name
                            LastUsed = $lastWriteTime
                            DaysInactive = $daysSinceLastUse
                            SpaceUsedGB = [math]::Round($profileSize / 1GB, 2)
                            FullPath = $profile.FullName
                        }
                    }
                }
            }
            catch {
                Write-Log "Error processing profile $($profile.Name): $($_.Exception.Message)" "WARN"
            }
        }

        Write-Progress -Activity "Scanning User Profiles" -Completed

        if ($results.Count -gt 0) {
            Write-Log "Found $($results.Count) old user profiles:" "SUCCESS"
            $results | Sort-Object -Property SpaceUsedGB -Descending | 
                      Format-Table -AutoSize | Out-String | 
                      ForEach-Object { Write-Log $_ "INFO" }
            
            $totalSpaceUsed = ($results | Measure-Object -Property SpaceUsedGB -Sum).Sum
            Write-Log "Total space used by old profiles: $([math]::Round($totalSpaceUsed, 2)) GB" "SUCCESS"
        } else {
            Write-Log "No user profiles older than $DaysOld days found." "INFO"
        }
        
        return $results
    }
    catch {
        Write-Log "Error scanning user profiles: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Remove-BloatwareApps {
    param([string]$Manufacturer)
    
    Write-Log "Checking for bloatware removal for manufacturer: $Manufacturer" "INFO"
    
    # Common bloatware patterns
    $bloatwarePatterns = @(
        "*McAfee*", "*Norton*", "*WildTangent*", "*Games*", "*Trial*",
        "*Candy*", "*Farm*", "*Bubble*", "*Hidden*", "*March*",
        "*Soda*", "*King*", "*Facebook*", "*Netflix*", "*Spotify*"
    )
    
    # Manufacturer-specific patterns
    $manufacturerPatterns = switch ($Manufacturer.ToLower()) {
        "dell" { @("*Dell*", "*Alienware*", "*SupportAssist*") }
        "hp" { @("*HP*", "*Hewlett*", "*Smart*") }
        "lenovo" { @("*Lenovo*", "*ThinkPad*", "*Vantage*") }
        "asus" { @("*ASUS*", "*ROG*", "*Armoury*") }
        default { @() }
    }
    
    $allPatterns = $bloatwarePatterns + $manufacturerPatterns
    
    try {
        $installedApps = Get-AppxPackage -AllUsers -ErrorAction Stop
        $appsToRemove = @()
        
        foreach ($pattern in $allPatterns) {
            $matchingApps = $installedApps | Where-Object { $_.Name -like $pattern }
            $appsToRemove += $matchingApps
        }
        
        $appsToRemove = $appsToRemove | Sort-Object -Property Name -Unique
        
        if ($appsToRemove.Count -gt 0) {
            Write-Log "Found $($appsToRemove.Count) bloatware apps to remove:" "INFO"
            
            foreach ($app in $appsToRemove) {
                try {
                    Write-Log "Removing: $($app.Name)" "INFO"
                    Remove-AppxPackage -Package $app.PackageFullName -AllUsers -ErrorAction Stop
                    Write-Log "Successfully removed: $($app.Name)" "SUCCESS"
                }
                catch {
                    Write-Log "Failed to remove $($app.Name): $($_.Exception.Message)" "WARN"
                }
            }
        } else {
            Write-Log "No bloatware apps found for removal." "INFO"
        }
    }
    catch {
        Write-Log "Error during bloatware removal: $($_.Exception.Message)" "ERROR"
    }
}

function Show-CleanupMenu {
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    
    do {
        Clear-Host
        $diskInfo = Get-DiskSpace
        
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "    ADVANCED SYSTEM CLEANUP TOOL v2.0     " -ForegroundColor Yellow
        Write-Host "==========================================" -ForegroundColor Cyan
        Write-Host "Current Disk Space: $($diskInfo.FreeGB)GB free / $($diskInfo.TotalGB)GB total ($($diskInfo.FreePercent)%)" -ForegroundColor Green
        Write-Host ""
        Write-Host "1. Run Complete System Cleanup" -ForegroundColor Green
        Write-Host "2. Run Cleanup Preview (Dry Run)" -ForegroundColor Green
        Write-Host "3. Find Large Files (1GB+)" -ForegroundColor Green
        Write-Host "4. Scan Old User Profiles" -ForegroundColor Green
        Write-Host "5. Remove Bloatware Apps" -ForegroundColor Green
        Write-Host "6. View Cleanup Log" -ForegroundColor Green
        Write-Host "7. Exit" -ForegroundColor Green
        Write-Host "==========================================" -ForegroundColor Cyan

        try {
            [int]$choice = Read-Host "Enter your choice (1-7)"
            $stopwatch.Restart()
            
            switch ($choice) {
                1 { 
                    Write-Log "User initiated complete system cleanup" "INFO"
                    Start-AdvancedSystemCleanup 
                }
                2 { 
                    Write-Log "User initiated cleanup dry run" "INFO"
                    Start-AdvancedSystemCleanup -DryRun 
                }
                3 { 
                    $size = Read-Host "Enter minimum file size in GB (default: 1.0)"
                    $sizeGB = if ($size -match '^\d+\.?\d*$') { [double]$size } else { 1.0 }
                    Find-LargeFiles -SizeThresholdGB $sizeGB 
                }
                4 { 
                    $days = Read-Host "Enter days of inactivity (default: 90)"
                    $daysOld = if ($days -match '^\d+$') { [int]$days } else { 90 }
                    Get-OldUserProfiles -DaysOld $daysOld 
                }
                5 { 
                    $manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
                    Remove-BloatwareApps -Manufacturer $manufacturer 
                }
                6 { 
                    if (Test-Path $LogPath) {
                        Get-Content $LogPath | Select-Object -Last 50 | ForEach-Object { Write-Host $_ }
                    } else {
                        Write-Log "No log file found at $LogPath" "WARN"
                    }
                }
                7 { 
                    Write-Log "User exited application" "INFO"
                    Write-Host "Cleanup completed. Log saved to: $LogPath" -ForegroundColor Yellow
                    return 
                }
                default { 
                    Write-Host "Invalid choice. Please select 1-7." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }
            }
            
            $stopwatch.Stop()
            $executionTime = $stopwatch.Elapsed
            Write-Host "`nExecution Time: $($executionTime.Minutes):$($executionTime.Seconds.ToString('00')).$($executionTime.Milliseconds.ToString('000'))" -ForegroundColor Cyan
            
            if ($choice -ne 7) {
                Write-Host "`nPress any key to return to menu..." -ForegroundColor Yellow
                $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
            }
        }
        catch {
            Write-Log "Menu error: $($_.Exception.Message)" "ERROR"
            Write-Host "An error occurred. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    } while ($true)
}

# Script execution
if ($Mode -eq 'Interactive') {
    Show-CleanupMenu
} elseif ($Mode -eq 'DryRun') {
    Start-AdvancedSystemCleanup -DryRun
} else {
    Start-AdvancedSystemCleanup
}

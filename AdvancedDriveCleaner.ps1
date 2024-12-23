function Start-AdvancedSystemCleanup {
    param (
        [int]$DaysOld = 30,
        [int]$LargeFileSizeGB = 1
    )

    Write-Host "Starting Advanced System Cleanup..." -ForegroundColor Green
    
    # Get initial disk space
    $drive = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
    $initialFreeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor Cyan

    # Define cleanup locations with descriptions
    $cleanupPaths = @{
        "$env:TEMP" = "Temporary Files"
        "$env:SystemRoot\Temp" = "Windows Temp"
        "$env:LOCALAPPDATA\Temp" = "Local App Temp"
        #"$env:USERPROFILE\Downloads" = "Downloads"
        #"$env:LOCALAPPDATA\Microsoft\Windows\INetCache" = "Internet Cache"
        "$env:SystemRoot\SoftwareDistribution\Download" = "Windows Update Cache"
        #"$env:LOCALAPPDATA\Microsoft\Teams\Cache" = "Teams Cache"
        "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" = "Explorer Cache"
        "$env:SystemRoot\Logs" = "Windows Logs"
        "$env:LOCALAPPDATA\CrashDumps" = "Crash Dumps"
    }

    # Track space cleaned
    $totalSpaceCleaned = 0

    foreach ($path in $cleanupPaths.GetEnumerator()) {
        if (Test-Path $path.Key) {
            $beforeSize = (Get-ChildItem $path.Key -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1GB
            Write-Host "`nCleaning $($path.Value) at $($path.Key)..." -ForegroundColor Yellow
            
            try {
                # Remove files older than specified days
                Get-ChildItem -Path $path.Key -Recurse -File -ErrorAction SilentlyContinue |
                    Where-Object { 
                        $_.LastWriteTime -lt (Get-Date).AddDays(-$DaysOld) -or 
                        $_.Length -gt ($LargeFileSizeGB * 1GB)
                    } |
                    Remove-Item -Force -ErrorAction SilentlyContinue

                $afterSize = (Get-ChildItem $path.Key -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum / 1GB
                $spaceCleaned = [math]::Round($beforeSize - $afterSize, 2)
                $totalSpaceCleaned += $spaceCleaned

                Write-Host "Cleaned $spaceCleaned GB from $($path.Value)" -ForegroundColor Green
            }
            catch {
                Write-Host "Error cleaning $($path.Value): $_" -ForegroundColor Red
            }
        }
    }

    # Additional Cleanup Tasks
    Write-Host "`nPerforming additional cleanup tasks..." -ForegroundColor Yellow

    # Clear DNS Cache
    ipconfig /flushdns | Out-Null
    Write-Host "DNS Cache cleared" -ForegroundColor Green

    # Clear System Restore Points except the most recent
    Write-Host "Cleaning old System Restore Points..." -ForegroundColor Yellow
    vssadmin delete shadows /for=C: /Quiet /all | Out-Null

    # Clean Package Cache
    Write-Host "Cleaning Package Cache..." -ForegroundColor Yellow
    Get-AppxPackage -AllUsers | Where-Object {-not ($_.IsFramework -or $_.IsBundle)} |
        ForEach-Object {
            Remove-AppxPackage -Package $_.PackageFullName -ErrorAction SilentlyContinue
        }

    # Run Windows Disk Cleanup silently
    Write-Host "Running Windows Disk Cleanup..." -ForegroundColor Yellow
    Start-Process cleanmgr -ArgumentList "/sagerun:1" -Wait -NoNewWindow

    # Get final disk space
    $drive = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
    $finalFreeSpace = [math]::Round($drive.FreeSpace / 1GB, 2)
    
    # Summary
    Write-Host "`nCleanup Summary:" -ForegroundColor Cyan
    Write-Host "Initial free space: $initialFreeSpace GB" -ForegroundColor White
    Write-Host "Final free space: $finalFreeSpace GB" -ForegroundColor White
    Write-Host "Total space recovered: $([math]::Round($finalFreeSpace - $initialFreeSpace, 2)) GB" -ForegroundColor Green
    Write-Host "Detailed cleanup completed: $([math]::Round($totalSpaceCleaned, 2)) GB" -ForegroundColor Green

    # Recommend restart
    Write-Host "`nRecommendation: Please restart your computer to complete the cleanup process." -ForegroundColor Yellow
}

# Run the cleanup with default parameters
Start-AdvancedSystemCleanup

# Prompt for restart
#$restart = Read-Host "`nWould you like to restart your computer now? (Y/N)"
#if ($restart -eq 'Y') {
#    #Restart-Computer -Force
#}
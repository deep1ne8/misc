# Function to display a browsable directory listing
function Get-BrowsableDirectory {
    param (
        [string]$currentPath = [Environment]::GetFolderPath("Desktop")
    )

    while ($true) {
        Clear-Host
        Write-Host "Current Directory: $currentPath" -ForegroundColor Cyan
        Write-Host "----------------------------------------" -ForegroundColor Cyan

        # Get the list of directories in the current path
        $directories = Get-ChildItem -Path $currentPath -Directory

        # Display the directories with numbers for selection
        for ($i = 0; $i -lt $directories.Count; $i++) {
            Write-Host "$($i + 1). $($directories[$i].Name)"
        }

        # Add options for going back or selecting the current directory
        Write-Host "----------------------------------------" -ForegroundColor Cyan
        Write-Host "0. Select this directory"
        if ($currentPath -ne (Split-Path -Path $currentPath -Parent)) {
            Write-Host "B. Go back to the parent directory"
        }
        Write-Host "Q. Quit"

        # Prompt the user for their choice
        $choice = Read-Host "Enter your choice (number, B, or Q)"

        if ($choice -eq "0") {
            # User selected the current directory
            return $currentPath
        } elseif ($choice -eq "B" -and $currentPath -ne (Split-Path -Path $currentPath -Parent)) {
            # Go back to the parent directory
            $currentPath = Split-Path -Path $currentPath -Parent
        } elseif ($choice -eq "Q") {
            # Quit the script
            Write-Host "Exiting script." -ForegroundColor Yellow
            exit
        } elseif ($choice -match "^\d+$" -and [int]$choice -le $directories.Count -and [int]$choice -gt 0) {
            # Navigate into the selected subdirectory
            $currentPath = $directories[[int]$choice - 1].FullName
        } else {
            Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
}

# Get the directory path from the user
$dirToCheck = Get-BrowsableDirectory

# Get the OneDrive root path
$onedrivePath = (Get-ItemProperty -Path "HKCU:\Environment" -Name "OneDrive").OneDrive

# Check if the directory is within the OneDrive path
if ($dirToCheck -like "$onedrivePath*") {
    Write-Host "The directory '$dirToCheck' is part of OneDrive sync." -ForegroundColor Green
} else {
    Write-Host "The directory '$dirToCheck' is NOT part of OneDrive sync." -ForegroundColor Red
}
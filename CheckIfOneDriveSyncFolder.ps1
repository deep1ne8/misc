# Function to prompt user for directory selection
function Get-DirectoryPath {
    Write-Host "Select an option to provide the directory path:" -ForegroundColor Cyan
    Write-Host "1. Enter the directory path manually"
    Write-Host "2. Browse and select a directory"
    $choice = Read-Host "Enter your choice (1 or 2)"

    if ($choice -eq 1) {
        # Prompt user to manually enter the directory path
        $dirToCheck = Read-Host "Enter the full path of the directory"
    } elseif ($choice -eq 2) {
        # Use a folder browser dialog to select the directory
        Add-Type -AssemblyName System.Windows.Forms
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Select a directory to check for OneDrive sync"
        $folderBrowser.RootFolder = [System.Environment+SpecialFolder]::Desktop

        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $dirToCheck = $folderBrowser.SelectedPath
        } else {
            Write-Host "No directory selected. Exiting script." -ForegroundColor Yellow
            exit
        }
    } else {
        Write-Host "Invalid choice. Exiting script." -ForegroundColor Red
        exit
    }

    return $dirToCheck
}

# Get the directory path from the user
$dirToCheck = Get-DirectoryPath

# Get the OneDrive root path
$onedrivePath = (Get-ItemProperty -Path "HKCU:\Environment" -Name "OneDrive").OneDrive

# Check if the directory is within the OneDrive path
if ($dirToCheck -like "$onedrivePath*") {
    Write-Host "The directory '$dirToCheck' is part of OneDrive sync." -ForegroundColor Green
} else {
    Write-Host "The directory '$dirToCheck' is NOT part of OneDrive sync." -ForegroundColor Red
}
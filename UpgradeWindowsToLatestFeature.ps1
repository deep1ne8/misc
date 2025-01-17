# Function to log verbose messages
function Log-Verbose {
    param (
        [string]$Message
    )
    Write-Verbose -Message "$(Get-Date -Format u): $Message"
}

# Prompt the user for manual or automatic mode
$Choice = Read-Host "Enter 'auto' to run automatically or 'manual' to choose OS version for upgrade"

if ($Choice -eq "manual") {
    $UpgradeChoice = Read-Host "Enter '10' for Windows 10 latest feature or '11' for Windows 11 latest feature"

    if ($UpgradeChoice -eq "10") {
        $Version = "10"
        $Release = "22H2"
    } elseif ($UpgradeChoice -eq "11") {
        $Version = "11"
        $Release = "24H2"
    } else {
        Write-Host "Invalid choice. Exiting." -ForegroundColor Red
        exit 1
    }
} elseif ($Choice -eq "auto") {
    # Determine the OS version
    Log-Verbose "Determining OS version..."
    $OSVersion = (Get-CimInstance Win32_OperatingSystem).Version
    if ($OSVersion -like "10.*") {
        $Version = "10"
        $Release = "22H2"
        Log-Verbose "Detected Windows 10. Preparing for $Release."
    } elseif ($OSVersion -like "11.*") {
        $Version = "11"
        $Release = "24H2"
        Log-Verbose "Detected Windows 11. Preparing for $Release."
    } else {
        Write-Host "Unsupported OS version: $OSVersion" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "Invalid input. Exiting." -ForegroundColor Red
    exit 1
}

# Ensure the setup directory exists
Set-ExecutionPolicy UnRestricted LocalMachine -Force -Confirm:$false

if (!(Test-Path "C:\WindowsSetup")) {
    Log-Verbose "Creating setup directory at C:\WindowsSetup."
    New-Item -ItemType Directory -Path "C:\WindowsSetup" | Out-Null
} else {
    Log-Verbose "Setup directory already exists at C:\WindowsSetup."
}

# Define file paths and URI
$DownloadPath = "C:\WindowsSetup\Win_${Version}_${Release}_English_x64v1.iso"
$FidoPath = "C:\WindowsSetup\Fido.ps1"

if (!(Test-Path $FidoPath)) {
    Write-Host "Fido.ps1 is not found in the current directory." -ForegroundColor Red
    Write-Host "Downloading Fido..."
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Fido.ps1" -OutFile "C:\WindowsSetup\Fido.ps1" -Verbose
}

# Check if the ISO file already exists
if (-not(Test-Path $DownloadPath)) {
    Log-Verbose "ISO not found at $DownloadPath. Starting download..."
    $URI = & C:\WindowsSetup\Fido.ps1 -Win $Version -Rel $Release -Arch x64 -Ed Pro -Lang English -GetUrl -Headers @{
    'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)'}
    }

# Validate the URL
if (-not $URI -or $URI -notmatch "^https?://") {
    Write-Host "Failed to retrieve a valid download URL. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host "Download URL successfully retrieved: $URI" -ForegroundColor Green

# Define the download path
$DownloadPath = "C:\WindowsSetup\Win_10_22H2_English_x64.iso"

# Start the BITS transfer to download the file
# Define the URL of the file to download and the local path to save it
$remoteUrl = $URI
$localPath = $DownloadPath
$jobName = "WindowsISODownload"

# Create a new BITS job
bitsadmin /create /download $jobName

# Add the file to the BITS job
bitsadmin /addfile $jobName $remoteUrl $localPath

# Set the priority of the BITS job to FOREGROUND for fastest download
bitsadmin /setpriority $jobName FOREGROUND

# Resume the BITS job to start downloading
bitsadmin /resume $jobName

# Monitor the download progress
do {
    # Get the current state of the BITS job
    $state = bitsadmin /getstate $jobName
    
    # Get the total bytes and bytes transferred
    $totalBytes = bitsadmin /getbytestotal $jobName
    $transferredBytes = bitsadmin /getbytestransferred $jobName
    
    # Calculate the progress percentage
    if ($totalBytes -gt 0) {
        $progress = [math]::Round(($transferredBytes / $totalBytes) * 100, 2)
    } else {
        $progress = 0
    }
    
    # Display the progress
    Write-Progress -Activity "Downloading $DownloadPath" -Status "$progress% Complete" -PercentComplete $progress
    
    # Wait for a short interval before checking the status again
    Start-Sleep -Seconds 5
} while ($state -eq "BG_JOB_STATE_TRANSFERRING")

# Complete the BITS job if the download is finished
if ($state -eq "BG_JOB_STATE_TRANSFERRED") {
    bitsadmin /complete $jobName
    Write-Output "Download completed successfully: $localPath"
} else {
    Write-Output "Download failed or was interrupted."
}

# Verify the file exists after download
if (!(Test-Path $DownloadPath)) {
    Write-Host "Download failed. ISO file not found at $DownloadPath." -ForegroundColor Red
    exit 1
}

Write-Host "File downloaded and verified: $DownloadPath" -ForegroundColor Green

# Mount the ISO and extract setup files
try {
    Log-Verbose "Mounting ISO from $DownloadPath..."
    $mountResult = Mount-DiskImage -ImagePath $DownloadPath
    $driveLetter = ($mountResult | Get-Volume).DriveLetter

    if (-not $driveLetter) {
        Log-Verbose "Failed to determine drive letter after mounting ISO. Exiting."
        Write-Host "Failed to determine drive letter. Exiting." -ForegroundColor Red
        exit 1
    }

    $ExtractPath = "${driveLetter}:\*"
    Log-Verbose "Copying setup files from mounted ISO..."
    Copy-Item -Path $ExtractPath -Destination "C:\WindowsSetup\" -Recurse -Force -Verbose

    Log-Verbose "Dismounting ISO..."
    Dismount-DiskImage -ImagePath $DownloadPath
} catch {
    Log-Verbose "Error occurred during mounting or file extraction: $_"
    Write-Host "An error occurred: $_" -ForegroundColor Red
    exit 1
}

# Cleanup ISO file
try {
    Log-Verbose "Removing ISO file from $DownloadPath..."
    #Remove-Item $DownloadPath -Force -Verbose
} catch {
    Log-Verbose "Failed to remove ISO file: $_"
    Write-Host "Failed to remove ISO file: $_" -ForegroundColor Yellow
}

# Verify setup file exists before starting
$SetupFile = "C:\WindowsSetup\setup.exe"
if (!(Test-Path $SetupFile)) {
    Log-Verbose "Setup file not found at $SetupFile. Exiting."
    Write-Host "Setup file not found. Exiting." -ForegroundColor Red
    exit 1
}

# Start the setup process
$ArgumentList = "/auto upgrade /eula accept /quiet"
Log-Verbose "Starting Windows setup with arguments: $ArgumentList"
Start-Process -NoNewWindow -Wait -FilePath $SetupFile -ArgumentList $ArgumentList -Verbose
Log-Verbose "Windows setup process initiated."
Write-Host "Windows setup process completed."

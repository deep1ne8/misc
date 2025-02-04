<#
.SYNOPSIS
    A PowerShell script that will deploy MSI and EXE apps silently, by downloading the app from URL using BITS with a nice progress bar - with continue feature.
.DESCRIPTION
    This script will download the application from the provided URL and install it silently.
    If the download is interrupted, the script will resume from the last downloaded byte.
    The script will also display a nice progress bar while downloading the application.
.PARAMETER Url
    The URL of the application to download and install.
.PARAMETER Path
    The path where the application will be saved.
.PARAMETER Continue
    If set to $true, the script will resume the download from the last downloaded byte.
.EXAMPLE
    .\Deploy-App.ps1 -Url "https://example.com/app.msi" -Path "C:\Temp\app.msi" -Continue $true
#>

param (
    [Parameter(Mandatory=$true)]
    [string]
    $Url,

    [Parameter(Mandatory=$true)]
    [string]
    $Path,

    [Parameter(Mandatory=$false)]
    [bool]
    $Continue = $false
)

$jobName = "Deploy-App"

# Create a new BITS job
$job = Start-BitsTransfer -Source $Url -Destination $Path -DisplayName $jobName

# If continue is set to $true, resume the download from the last downloaded byte
if ($Continue) {
    $job | Resume-BitsTransfer
}

# Wait for the download to complete
$job | Wait-BitsTransfer

# Display the progress bar
$progress = $job | Get-BitsTransfer
Write-Progress -Activity "Downloading $Url" -Status "$($progress.BytesTransferred / 1MB) MB of $($progress.BytesTotal / 1MB) MB" -PercentComplete ($progress.BytesTransferred / $progress.BytesTotal) * 100

# Install the application silently
if ($Path -like "*.msi") {
    Start-Process -FilePath "msiexec.exe" -ArgumentList "/i `"$Path`" /quiet /norestart" -Wait -NoNewWindow
} elseif ($Path -like "*.exe") {
    Start-Process -FilePath $Path -ArgumentList "/quiet /norestart" -Wait -NoNewWindow
} else {
    Write-Error "Unsupported file type. Only .msi and .exe files are supported."
    return
}

# Clean up the job
Remove-BitsTransfer -Job $jobName -Confirm:$false
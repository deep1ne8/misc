param (
    [Parameter(Mandatory=$true)]
    [string]$Url,
    [Parameter(Mandatory=$true)]
    [string]$DestinationPath,
    [Parameter(Mandatory=$false)]
    [bool]$Continue = $True
)

# Check if the destination path exists
if (!(Test-Path $DestinationPath)) {
    Write-Host "Destination path does not exist. Creating it."
    New-Item -ItemType Directory -Path $DestinationPath | Out-Null
}

# Check if the script is running as administrator
$currentPrincipal = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
if (!$currentPrincipal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as an administrator."
    exit
}

# Install BITS if it is not installed
if (-not (Get-WindowsFeature -Name BITS -ErrorAction SilentlyContinue).Installed) {
    Write-Host "BITS is not installed. Installing BITS."
    Install-WindowsFeature -Name BITS
}

# Install WGET if it is not installed
if (-not (Get-Command wget -ErrorAction SilentlyContinue)) {
    Write-Host "WGET is not installed. Installing WGET."
    # Install WGET using Chocolatey
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Host "Chocolatey is not installed. Installing Chocolatey."
        Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
    }
    Import-Module $env:ChocolateyInstall\helpers\chocolateyProfile.psm1
    refreshenv
    Start-Sleep -Seconds 3
    choco install wget -y -force
}

$jobName = "InteractiveDownload"

# Choose the method of download
$downloadMethod = Read-Host "Choose download method: BITS or WGET"

if ($downloadMethod -eq "BITS") {
    # Create a new BITS job
    try {
        $job = Start-BitsTransfer -Source $Url -Destination $DestinationPath -DisplayName $jobName -Asynchronous
    } catch {
        Write-Host "Error starting BITS transfer: $_"
        if ($job.JobState -eq "Suspended") {
            $job | Resume-BitsTransfer
        }
    }

    # If continue is set to $true, resume the download from the last downloaded byte
    if ($Continue) {
        $job | Resume-BitsTransfer
    }

    # Monitor the download progress
    do {
        $progress = $job | Get-BitsTransfer
        if ($progress.BytesTotal -ne 0) {
            Write-Progress -Activity "Downloading $Url" -Status "$($progress.BytesTransferred / 1MB) MB of $($progress.BytesTotal / 1MB) MB" -PercentComplete (($progress.BytesTransferred / $progress.BytesTotal) * 100)
        } else {
            Start-Sleep -Seconds 10
        }
        Start-Sleep -Seconds 2
    } while ($job.JobState -ne "Transferred" -and $job.JobState -ne "Suspended")

    # Complete the BITS job
    Complete-BitsTransfer -BitsJob $job

} elseif ($downloadMethod -eq "WGET") {
    $wgetCommand = "wget `"$Url`" -O `"$DestinationPath`""
    if ($Url -notmatch '^(https?|ftp)://[^\s/$.?#].[^\s]*$') {
        Write-Host "Invalid URL format."
        return
    }
    if ($null -eq $DestinationPath -or $DestinationPath -notmatch '^.+$') {
        Write-Host "Invalid characters in destination path."
        return
    }
    try {
        Invoke-Expression $wgetCommand
    } catch {
        Write-Host "Error during WGET download: $_"
    }

} else {
    Write-Host "Invalid download method selected. Please choose either BITS or WGET."
}


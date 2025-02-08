param (
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$Url,
    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string]$DestinationPath,
    [Parameter(Mandatory=$false)]
    [bool]$Continue = $false
)

# Check if the destination path exists
if (!(Test-Path $DestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $DestinationPath | Out-Null
    } catch {
        Write-Host "Failed to create destination path: $_" -ForegroundColor Red
        return
    }
}

# Install BITS if it is not installed
if (-not (Where.exe BITS)) {
    try {
        dism /online /Get-Features | findstr /i "BITS"
        dism /online /Enable-Feature /FeatureName:"BITS"
    } catch {
        Write-Host "Failed to install BITS: $_" -ForegroundColor Red
        return
    }
}

# Install WGET if it is not installed
if (-not (Where.exe wget)) {
    try {
        # Install WGET using Chocolatey
        if (-not (Where.exe choco -ErrorAction SilentlyContinue)) {
            Write-Host "Chocolatey is not installed. Installing Chocolatey." -ForegroundColor Black -CommandColor Green
            Set-ExecutionPolicy Bypass -Scope Process -Force; Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        }
        Import-Module $env:ChocolateyInstall\helpers\chocolateyProfile.psm1
        refreshenv
        Start-Sleep -Seconds 3
        choco install wget -y -force
    } catch {
        Write-Host "Failed to install WGET: $_" -ForegroundColor Red
        return
    }
}

$jobName = "InteractiveDownload"

# Choose the method of download
$downloadMethod = Read-Host "Choose download method: BITS or WGET"

if ($downloadMethod -eq "BITS") {
    # Create a new BITS job
    try {
        $job = Start-BitsTransfer -Source $Url -Destination $DestinationPath -DisplayName $jobName -Asynchronous
    } catch {
        Write-Host "Error starting BITS transfer: $_" -ForegroundColor Red
        if ($job.JobState -eq "Suspended") {
            try {
                $job | Resume-BitsTransfer
            } catch {
                Write-Host "Failed to resume BITS transfer: $_" -ForegroundColor Red
                return
            }
        }
    }

    # If continue is set to $true, resume the download from the last downloaded byte
    if ($Continue) {
        try {
            $job | Resume-BitsTransfer
        } catch {
            Write-Host "Failed to resume BITS transfer: $_" -ForegroundColor Red
            return
        }
    }

    # Monitor the download progress
    do {
        try {
            $progress = $job | Get-BitsTransfer
        } catch {
            Write-Host "Failed to get BITS job progress: $_" -ForegroundColor Red
            return
        }
        if ($progress.BytesTotal -ne 0) {
            Write-Progress -Activity "Downloading $Url" -Status "$($progress.BytesTransferred / 1MB) MB of $($progress.BytesTotal / 1MB) MB" -PercentComplete (($progress.BytesTransferred / $progress.BytesTotal) * 100)
        } else {
            Start-Sleep -Seconds 10
        }
        Start-Sleep -Seconds 2
    } while ($job.JobState -ne "Transferred" -and $job.JobState -ne "Suspended")

    # Complete the BITS job
    try {
        Complete-BitsTransfer -BitsJob $job
    } catch {
        Write-Host "Failed to complete BITS job: $_" -ForegroundColor Red
        return
    }

} elseif ($downloadMethod -eq "WGET") {
    $wgetCommand = "wget `"$Url`" -O `"$DestinationPath`""
    if ($Url -notmatch '^(https?|ftp)://[^\s/$.?#].[^\s]*$') {
        Write-Host "Invalid URL format." -ForegroundColor Red
        return
    }
    if ($null -eq $DestinationPath -or $DestinationPath -notmatch '^.+$') {
        Write-Host "Invalid characters in destination path." -ForegroundColor Red
        return
    }
    try {
        Invoke-Expression $wgetCommand
    } catch {
        Write-Host "Error during WGET download: $_" -ForegroundColor Red
        return
    }

} else {
    Write-Host "Invalid download method selected. Please choose either BITS or WGET." -ForegroundColor Red
    return
}


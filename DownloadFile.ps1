param (
    [Parameter(Mandatory=$true)]
    [string]$Url,

    [Parameter(Mandatory=$true)]
    [string]$DestinationPath,

    [Parameter(Mandatory=$false)]
    [bool]$Continue = $false
)

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
        exit
    }
    if ($null -eq $DestinationPath -or $DestinationPath -notmatch '^\w+://[^\s/$.?#].[^\s]*$') {
        Write-Host "Invalid characters in destination path."
        exit
    }
    $wgetCommand = "wget $Url -O $DestinationPath"
    if ($Continue) {
    Write-Host "Invalid download method selected. Please choose either 'BITS' or 'WGET'."
    }
    try {
        Invoke-Expression $wgetCommand
    } catch {
        Write-Host "Error during WGET download: $_"
    }

} else {
    Write-Host "Invalid download method selected. Please choose either BITS or WGET."
}


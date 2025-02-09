<#
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
#>

Write-Host "`n"
# Check if the destination path exists
Write-Host "Enter destination path: ==>  " -ForegroundColor Blue -NoNewline
$DestinationPath = Read-Host
Start-Sleep -Seconds 2
Write-Host "`n"
Write-Host "Enter URL: ==>  " -ForegroundColor Blue -NoNewline
$Url = Read-Host
Start-Sleep -Seconds 2
Write-Host "`n"
Write-Host "Starting BITS download..." -ForegroundColor White -BackgroundColor Green
Write-Host "Please wait..." -ForegroundColor White -BackgroundColor Green
Start-Sleep -Seconds 2
Write-Host "`n"

# Check if the destination path exists and is writable
if (!(Test-Path $DestinationPath)) {
    try {
        New-Item -ItemType Directory -Path $DestinationPath | Out-Null
    } catch {
        Write-Host "Failed to create destination path: $_" -ForegroundColor Red
        return
    }
}

# Downloading with BITS
if (Get-Service -Name "BITS") {
    try {
            # Define the BITS job name
            $jobName = "InteractiveDownload"

            # Create a new BITS job
            try {
                $job = powershell.exe -Command "Start-BitsTransfer -Source $Url -Destination $DestinationPath -DisplayName $jobName -Asynchronous"
            } catch {
                Write-Host "Error starting BITS transfer: $_" -ForegroundColor Red
                if ($job.JobState -eq "Suspended") {
                    try {
                        $job | powershell.exe -Command "Resume-BitsTransfer"
                    } catch {
                        Write-Host "Failed to resume BITS transfer: $_" -ForegroundColor Red
                        return
                    }
                }
            }
            
                # Monitor the download progress
                do {
                    try {
                        $progress = $job | powershell.exe -Command "Get-BitsTransfer"
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
                    powershell.exe -Command "Complete-BitsTransfer -BitsJob $job"
                } catch {
                    Write-Host "Failed to complete BITS job: $_" -ForegroundColor Red
                    return
                }
         Write-Host "Download completed successfully!" -ForegroundColor Green
            return

        } catch {
        Write-Host "Error during BITS download: $_" -ForegroundColor Red
            return
    }
    Start-Sleep -Seconds 1
    return
}



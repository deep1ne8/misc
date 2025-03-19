# PowerShell script to troubleshoot and fix Windows 11 temporary profile issues
$logFilePath = "C:\Logs\Profile_Troubleshoot_Log.txt"

# Create the log file if it doesn't exist
if (-not (Test-Path "C:\Logs")) {
    New-Item -Path "C:\" -Name "Logs" -ItemType Directory
}

# Function to write to both console and log file
function Write-Message {
    param (
        [string]$message,
        [string]$color = "White"
    )
    
    # Write to console
    Write-Host $message -ForegroundColor $color
    
    # Append to log file
    $message | Out-File -Append -FilePath $logFilePath
}

# 1. Check Event Viewer for user profile errors
Write-Message "Checking Event Viewer for profile errors..."
$profileErrors = Get-WinEvent -LogName Application | Where-Object { $_.Id -in (1511, 1515, 1500, 1530, 1533) }
if ($profileErrors) {
    Write-Message "User Profile Service Errors Found!" "Red"
    $profileErrors | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize | Out-File -Append -FilePath $logFilePath
} else {
    Write-Message "No User Profile Errors Found in Event Viewer." "Green"
}

# 2. Check for profile corruption in the Windows Registry
Write-Message "Checking Windows Registry for profile corruption..."
$corruptProfiles = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Where-Object { $_.Name -match '\.bak$' }
if ($corruptProfiles) {
    Write-Message "Corrupt profile detected:" "Red"
    $corruptProfiles.Name | Out-File -Append -FilePath $logFilePath
    Write-Message "Deleting corrupt profile registry entry..." "Red"
    Remove-Item -Path $corruptProfiles.PSPath -Recurse -Force
    Write-Message "Corrupt profile registry entry deleted. Restart the computer." "Green"
} else {
    Write-Message "No corrupt profiles found in registry." "Green"
}

# 3. Check if the user profile exists on disk
Write-Message "Checking if user profile exists on disk..."
$profileList = Get-ChildItem "C:\Users" | Select-Object -ExpandProperty Name
$profileList | ForEach-Object { Write-Message "Existing Profile: $_" }
if ($profileList -contains 'TEMP') {
    Write-Message "Temporary profile detected. Deleting..." "Red"
    Remove-Item -Path "C:\Users\TEMP" -Recurse -Force
    Write-Message "TEMP profile deleted. Restart the computer." "Green"
} else {
    Write-Message "No TEMP profile found." "Green"
}

# 4. Check profile folder permissions
Write-Message "Checking profile folder permissions..."
$profileFolder = "C:\Users\BMason" # Replace with the actual profile name
$acl = Get-Acl $profileFolder
$acl | Format-List | Out-File -Append -FilePath $logFilePath
# Ensure the user has FullControl (F)
icacls $profileFolder /grant BMason:(F) /T
Write-Message "Permissions for $profileFolder updated." "Green"

# 5. Check if Group Policy Profile Path is configured
Write-Message "Checking Group Policy Profile Path (if configured)..."
try {
    $gpoProfilePath = Get-GPResultantSetOfPolicy -ReportType Html -Path "C:\GPOReport.html"
    if ($gpoProfilePath -match "ProfilePath") {
        Write-Message "Group Policy Profile Path is configured." "Red"
    } else {
        Write-Message "No GPO Profile Path configured." "Green"
    }
} catch {
    Write-Message "Unable to retrieve Group Policy result. Skipping..." "Yellow"
}

# 6. Check disk for errors
Write-Message "Checking disk for errors (this may take a moment)..."
$diskCheckResult = chkdsk
if ($diskCheckResult -match "is healthy") {
    Write-Message "Disk check completed with no errors." "Green"
} else {
    Write-Message "Potential disk errors detected. Review logs above." "Red"
}

# Final restart reminder
Write-Message "Please restart the computer to apply changes." "Yellow"

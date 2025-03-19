# PowerShell script to troubleshoot and fix Windows 11 temporary profile issues
$logFilePath = "C:\Logs\Profile_Troubleshoot_Log.txt"

# Create the log directory if it doesn't exist
if (-not (Test-Path "C:\Logs")) {
    New-Item -Path "C:\" -Name "Logs" -ItemType Directory | Out-Null
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
try {
    $profileErrors = Get-WinEvent -LogName Application | Where-Object { $_.Id -in (1511, 1515, 1500, 1530, 1533) }
    if ($profileErrors) {
        Write-Message "User Profile Service Errors Found!" "Red"
        $profileErrors | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize | Out-File -Append -FilePath $logFilePath
    } else {
        Write-Message "No User Profile Errors Found in Event Viewer." "Green"
    }
} catch {
    Write-Message "Error while checking Event Viewer: $_" "Red"
}

# 2. Check for profile corruption in the Windows Registry
Write-Message "Checking Windows Registry for profile corruption..."
try {
    $corruptProfiles = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList" | Where-Object { $_.Name -match '\.bak$' }
    if ($corruptProfiles) {
        Write-Message "Corrupt profile detected:" "Red"
        $corruptProfiles.Name | Out-File -Append -FilePath $logFilePath
        foreach ($profile in $corruptProfiles) {
            Write-Message "Deleting corrupt profile registry entry: $($profile.Name)" "Red"
            Remove-Item -Path $profile.PSPath -Recurse -Force
        }
        Write-Message "Corrupt profile registry entries deleted. Restart the computer." "Green"
    } else {
        Write-Message "No corrupt profiles found in registry." "Green"
    }
} catch {
    Write-Message "Error while checking registry: $_" "Red"
}

# 3. Check if the user profile exists on disk
Write-Message "Checking if user profile exists on disk..."
try {
    $profileList = Get-ChildItem "C:\Users" | Select-Object -ExpandProperty Name
    $profileList | ForEach-Object { Write-Message "Existing Profile: $_" }
    if ($profileList -contains 'TEMP') {
        Write-Message "Temporary profile detected...." "Red"
        try {
            #Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            #Remove-Item -Path "C:\Users\TEMP" -Recurse -Force -ErrorAction SilentlyContinue
            $profileList
            Write-Message "TEMP profile detected.." "Green"
        } catch {
            Write-Message "Please delete TEMP profile." "Red"
        }
    } else {
        Write-Message "No TEMP profile found." "Green"
    }
} catch {
    Write-Message "Error while checking user profiles on disk: $_" "Red"
}

# 4. Check profile folder permissions for all users
Write-Message "Checking profile folder permissions for all user profiles..."
try {
    $profiles = Get-ChildItem "C:\Users" | Where-Object { $_.PSIsContainer }
    foreach ($profile in $profiles) {
        $profileFolder = $profile.FullName
        Write-Message "Checking permissions for $profileFolder..."
        $acl = Get-Acl $profileFolder
        $acl | Format-List | Out-File -Append -FilePath $logFilePath
        Write-Host "$acl | Format-List"
    }
} catch {
    Write-Message "Error while checking profile folder permissions: $_" "Red"
}

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
try {
    $diskCheckResult = Repair-Volume -DriveLetter C -Scan
    if ($diskCheckResult -match "NoErrorsFound") {
        Write-Message "Disk check completed with no errors." "Green"
    } else {
        Write-Message "Potential disk errors detected. Review logs above." "Red"
    }
} catch {
    Write-Message "Error while performing disk check: $_" "Red"
}

# Final restart reminder
Write-Message "Please restart the computer to apply changes." "Yellow"

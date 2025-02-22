# PowerShell Verbose Logging Template

# Define log file path
$LogFile = "C:\Windows\Temp\EnableFilesOnDemand_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Function to Write Log
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "[$Timestamp] [$Level] $Message"
    
    # Write to console
    if ($Level -eq "ERROR") {
        Write-Host $LogEntry -ForegroundColor Red
    } elseif ($Level -eq "WARNING") {
        Write-Host $LogEntry -ForegroundColor Yellow
    } elseif ($Level -eq "VERBOSE") {
        Write-Host $Message -ForegroundColor Green
    } else {
        Write-Host $LogEntry -ForegroundColor Green
    }
    
    # Write to log file
    Add-Content -Path $LogFile -Value $LogEntry
}

# Enable Verbose Logging
$VerbosePreference = "Continue"

# Start Logging
Write-Log "Script execution started." "INFO"

try {
    # Enable Files On Demand
    Write-Log "Getting user profile folder..." "VERBOSE"
    Write-Host ""
    # Get the logged-in username (domain\username format)
    $LoggedInUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName

    if ($LoggedInUser -match '\\') {
        # Extract the username part (remove the domain prefix)
        $UserName = $LoggedInUser -replace '^.*\\', ''
        #$UserName = $LoggedInUser

        # Get the profile folder path
        $UserProfileFolder = Get-ChildItem "$env:SystemDrive\Users" | Where-Object {
            $_.Name -like "$UserName*"
        } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

        Write-Log "Getting user profile folder for: $UserName" "VERBOSE"
        Write-Host ""
        if ($UserProfileFolder) {
            Write-Log "User Profile Folder: $($UserProfileFolder.FullName)" "INFO"
            Write-Host ""
        } else {
            Write-Log "No matching user profile folder found for: $UserName" "WARNING"
        }
    } else {
        Write-Log "No logged-in user found or invalid username format." "ERROR"
    }

    # Resolve OneDrive path dynamically
    Write-Log "`n"
    Write-Log "Getting OneDrive folder path for: $UserName" "VERBOSE"
    Write-Host ""
    try {
        $OneDrivePath = (Get-ChildItem -Path "$($UserProfileFolder.FullName)" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "OneDrive - *" }).FullName
    } catch {
        throw "OneDrive folder path not found for user $LoggedInUser."
    }

    Write-Log "OneDrive Folder Path: $OneDrivePath" "INFO"
    Set-Location -Path $OneDrivePath
    Start-Sleep -Seconds 3
    Write-Log "`n"

    # Get the current state of files in the directory
    $CurrentFileState = Get-ChildItem -Path $PWD -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue | 
        Select-Object FullName, Name, Attributes, Mode, Length, CreationTime

    Write-Log "`nCurrent File State:`n" "INFO"
    $CurrentFileState | Format-Table -AutoSize
    Write-Log "`n"
    Start-Sleep -Seconds 3

    Write-Log "Checking if all files are online-only..." "VERBOSE"
    Write-Host ""

    # Identify files that are **not** online-only
    $NotOnlineOnlyFiles = $CurrentFileState | 
        Where-Object { ([int]$_.Attributes -ne 5248544) }  # Ensure they are NOT cloud-only

    if ($NotOnlineOnlyFiles.Count -eq 0) {
        Write-Log "✅ All files are already set to online-only." "INFO"
        return
    } else {
        Write-Log "❌ Some files are not online-only. Preparing to update..." "WARNING"
        Start-Sleep -Seconds 3
    }

    # File attribute reference guide
    Write-Log "`nFile Attribute Guide:" "INFO"
    Write-Host "==========================================" 
    Write-Host "File State            | Attribute Value" -ForegroundColor White
    Write-Host "------------------------------------------"
    Write-Host "Cloud-Only           | 5248544" -ForegroundColor Green
    Write-Host "Always Available     | 525344" -ForegroundColor Yellow
    Write-Host "Locally Available    | ReparsePoint" -ForegroundColor Red
    Write-Host "==========================================`n"
    Start-Sleep -Seconds 5

    # Enable files on demand (Change state to Cloud-Only)
    Write-Log "`nUpdating file states..." "VERBOSE"
    Start-Sleep -Seconds 3

    try {
        foreach ($File in $NotOnlineOnlyFiles) {
            $FilePath = $File.FullName
            $Attributes = [int]$File.Attributes

            if ($Attributes -eq 5248544) {
                Write-Log "✅ File '$FilePath' is already cloud-only." "INFO"
            } else {
                attrib.exe +U "$FilePath"
                Write-Log "🔄 File '$FilePath' state changed to cloud-only." "INFO"
            }
        }
    } catch {
        Write-Log "❌ Error occurred while enabling files on demand: $_" "ERROR"
        return
    }

    Write-Log "`n✅ Process completed successfully." "INFO"
} catch {
    Write-Log "An error occurred: $_" "ERROR"
} finally {
    Write-Log "Script execution completed." "INFO"
}



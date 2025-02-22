# Enable Files On Demand

Write-Host "Getting user profile folder..." -ForegroundColor Green
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

    Write-Host "Getting user profile folder for: $UserName" -ForegroundColor Yellow
    Write-Host ""
    if ($UserProfileFolder) {
        Write-Host "User Profile Folder: $($UserProfileFolder.FullName)" -ForegroundColor Green
        Write-Host ""
    } else {
        Write-Host "No matching user profile folder found for: $UserName" -ForegroundColor Yellow
    }
} else {
    Write-Host "No logged-in user found or invalid username format." -ForegroundColor Red
}

# Resolve OneDrive path dynamically
Write-Host "`n"
Write-Host "Getting OneDrive folder path for: $UserName" -ForegroundColor Yellow
Write-Host ""
try {
    $OneDrivePath = (Get-ChildItem -Path "$($UserProfileFolder.FullName)" -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "OneDrive - *" }).FullName
} catch {
    throw "OneDrive folder path not found for user $LoggedInUser."
}

Write-Host "OneDrive Folder Path: $OneDrivePath" -ForegroundColor Green
Set-Location -Path $OneDrivePath
Start-Sleep -Seconds 3
Write-Host "`n"

# Get the current state of files in the directory
$CurrentFileState = Get-ChildItem -Path $PWD -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue | 
    Select-Object FullName, Name, Attributes, Mode, Length, CreationTime

Write-Host "`nCurrent File State:`n" -ForegroundColor Cyan
$CurrentFileState | Format-Table -AutoSize
Write-Host "`n"
Start-Sleep -Seconds 3

Write-Host "Checking if all files are online-only..." -ForegroundColor Green
Write-Host ""

# Identify files that are **not** online-only
$NotOnlineOnlyFiles = $CurrentFileState | 
    Where-Object { ([int]$_.Attributes -ne 5248544) }  # Ensure they are NOT cloud-only

if ($NotOnlineOnlyFiles.Count -eq 0) {
    Write-Host "✅ All files are already set to online-only." -ForegroundColor Green
    return
} else {
    Write-Host "❌ Some files are not online-only. Preparing to update..." -ForegroundColor Red
    Start-Sleep -Seconds 3
}

# File attribute reference guide
Write-Host "`nFile Attribute Guide:" -ForegroundColor Cyan
Write-Host "==========================================" 
Write-Host "File State            | Attribute Value" -ForegroundColor White
Write-Host "------------------------------------------"
Write-Host "Cloud-Only           | 5248544" -ForegroundColor Green
Write-Host "Always Available     | 525344" -ForegroundColor Yellow
Write-Host "Locally Available    | ReparsePoint" -ForegroundColor Red
Write-Host "==========================================`n"
Start-Sleep -Seconds 5

# Enable files on demand (Change state to Cloud-Only)
Write-Host "`nUpdating file states..." -ForegroundColor Green
Start-Sleep -Seconds 3

try {
    foreach ($File in $NotOnlineOnlyFiles) {
        $FilePath = $File.FullName
        $Attributes = [int]$File.Attributes

        if ($Attributes -eq 5248544) {
            Write-Host "✅ File '$FilePath' is already cloud-only." -ForegroundColor Green
        } else {
            attrib.exe +U "$FilePath"
            Write-Host "🔄 File '$FilePath' state changed to cloud-only." -ForegroundColor Green
        }
    }
} catch {
    Write-Host "❌ Error occurred while enabling files on demand: $_" -ForegroundColor Red
    return
}

Write-Host "`n✅ Process completed successfully." -ForegroundColor Cyan











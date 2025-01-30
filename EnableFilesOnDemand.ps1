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

Write-Host "Getting OneDrive folder path for: $UserName" -ForegroundColor Yellow
Write-Host ""
try {
    $OneDrivePath = (Get-ChildItem -Path "$($UserProfileFolder.FullName)" -Directory -ErrorAction Stop | Where-Object { $_.Name -like "OneDrive - *" }).FullName
} catch {
    throw "OneDrive folder path not found for user $LoggedInUser."
}

Write-Host "OneDrive Folder Path: $OneDrivePath" -ForegroundColor Green
Set-Location -Path $OneDrivePath
Start-Sleep -Seconds 3
Write-Host "`n"

Write-Host "Getting the current file state" -ForeGroundColor Green
Write-Host "`n"
Start-Sleep -Seconds 3

# Check if files are always available
$CurrentFileState = Get-ChildItem -Path $PWD -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue | Format-Table Attributes, Mode, Name, Length, CreationTime
Write-Host "`n"
Write-Output $CurrentFileState
Write-Host "`n"


Write-Host "Enabling files on demand" -ForegroundColor Green
Write-Host ""
# Check if files are online only
$CheckFilesAttrib = Get-ChildItem -Path $PWD -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue | Where-Object {$_.Attributes -eq '5258544'} | Format-Table Attributes
if ($CheckFilesAttrib -eq "5248544"){
	        Write-Host "All the files attributes are already set to online only" -ForegroundColor Green
 	return
  }else {   Write-Host "All the files attributes are not set to online only" -ForegroundColor Red
  Start-Sleep -Seconds 3

# Enable files on demand by removing ReparsePoint attribute
Write-Host "`n"
Write-Host "Below is the guide for the file state, according to it's attribute" -ForeGroundColor Green
Write-Host "`n==================================================================="
Write-Host "File State 	        Attribute" -ForeGroundColor White
Write-Host "------------------------------"
Write-Host "Cloud-Only 	        5248544" -ForeGroundColor Green
Write-Host "Always available 	525344" -ForeGroundColor Green
Write-Host "Locally Available 	ReparsePoint" -ForeGroundColor Green
Write-Host "`n==================================================================="
Start-Sleep -Seconds 5
Write-Host "`n"
Write-Host "Changing the file state" -ForeGroundColor Green
Write-Host ""
Start-Sleep -Seconds 3
try {
    Write-Output $CurrentFileState | 
    ForEach-Object {
        attrib.exe $_.fullname +U /s
    }
} catch {
    Write-Host "Error occurred while enabling files on demand: $_" -ForeGroundColor Red
    return
}  
try { 
    Write-Host "Verifying the updated file state" -ForeGroundColor Green
    Write-Host ""
    Start-Sleep -Seconds 3
    if ($CurrentFileState -eq "5248544") {
        Write-Host "All the files attributes are set to online only" -ForegroundColor Green
        return
    } else {
        Write-Host "All the files attributes are not set to online only" -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }
} catch {
    Write-Host "Error occurred while verifying the updated file state: $_" -ForeGroundColor Red
    return
        }
}







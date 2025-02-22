# Get the logged-in username
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
    } else {
        Write-Host "No matching user profile folder found for: $UserName" -ForegroundColor Yellow
    }
} else {
    Write-Host "No logged-in user found or invalid username format." -ForegroundColor Red
}
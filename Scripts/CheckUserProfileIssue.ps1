# Check for User Profile Service errors in Event Viewer
Write-Host "Checking Event Viewer for profile errors..." -ForegroundColor Cyan
$profileErrors = Get-WinEvent -LogName Application -ErrorAction SilentlyContinue |
    Where-Object { $_.Id -in @(1508,1511,1515,1516) }

if ($profileErrors) {
    Write-Host "User Profile Service Errors Found!" -ForegroundColor Red
    $profileErrors | Select-Object TimeCreated, Id, Message | Format-List
} else {
    Write-Host "No User Profile Service errors found." -ForegroundColor Green
}

# Check Registry for Corrupt Profile Entries
Write-Host "`nChecking Windows Registry for profile corruption..." -ForegroundColor Cyan
$profileList = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
$profiles = Get-ChildItem -Path $profileList -ErrorAction SilentlyContinue

foreach ($profile in $profiles) {
    if ($profile.PSChildName -match "\.bak$") {
        Write-Host "Corrupt profile detected: $($profile.PSChildName)" -ForegroundColor Yellow
    }
}

# Check User Profile Exists on Disk
Write-Host "`nChecking if user profile exists on disk..." -ForegroundColor Cyan
$usersFolder = "C:\Users"
$allUsers = Get-ChildItem -Path $usersFolder -ErrorAction SilentlyContinue
Write-Host "Existing User Profiles on Disk:" -ForegroundColor Magenta
$allUsers | Select-Object Name

# Check if User Profile Folder is Accessible
$userProfilePath = "$usersFolder\$env:UserName"
if (Test-Path $userProfilePath) {
    Write-Host "User profile folder exists: $userProfilePath" -ForegroundColor Green
} else {
    Write-Host "User profile folder NOT found!" -ForegroundColor Red
}

# Check Profile Permissions
Write-Host "`nChecking profile folder permissions..." -ForegroundColor Cyan
$acl = Get-Acl -Path $userProfilePath -ErrorAction SilentlyContinue
if ($acl) {
    Write-Host "Permissions for ${userProfilePath}:" -ForegroundColor Magenta
    $acl.Access | Format-Table IdentityReference, FileSystemRights, AccessControlType
} else {
    Write-Host "Could not retrieve permissions. Possible access issue." -ForegroundColor Red
}

# Check Group Policy Profile Path
Write-Host "`nChecking Group Policy Profile Path (if configured)..." -ForegroundColor Cyan
$gpoProfilePath = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ProfileList" -ErrorAction SilentlyContinue
if ($gpoProfilePath) {
    Write-Host "GPO Profile Path: $gpoProfilePath" -ForegroundColor Magenta
} else {
    Write-Host "No GPO Profile Path configured." -ForegroundColor Green
}

# Check Group Policy Results (Admin Required)
if ($env:USERNAME -eq "Administrator") {
    Write-Host "`nChecking applied Group Policies..." -ForegroundColor Cyan
    gpresult /h C:\Temp\GPReport.html
    Write-Host "Group Policy report saved to C:\Temp\GPReport.html" -ForegroundColor Magenta
} else {
    Write-Host "Run PowerShell as Administrator to check Group Policies." -ForegroundColor Yellow
}

# Run this script as Administrator
Write-Host "Checking for Windows 11 Temporary Profile Issues..." -ForegroundColor Cyan

# ---------------------- CHECK EVENT VIEWER FOR PROFILE ERRORS ----------------------
Write-Host "`nChecking Event Viewer for profile errors..." -ForegroundColor Yellow
$profileErrors = Get-EventLog -LogName Application -Source "User Profile Service" -Newest 20 | Where-Object { $_.EventID -in (1511, 1515) }

if ($profileErrors) {
    Write-Host "User Profile Service Errors Found!" -ForegroundColor Red
    $profileErrors | Select-Object TimeGenerated, EventID, Message | Format-Table -AutoSize
} else {
    Write-Host "No User Profile Errors Found in Event Viewer." -ForegroundColor Green
}

# ---------------------- CHECK WINDOWS REGISTRY FOR CORRUPT PROFILES ----------------------
Write-Host "`nChecking Windows Registry for profile corruption..." -ForegroundColor Yellow
$profileRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
$corruptProfiles = Get-ChildItem -Path $profileRegPath | Where-Object { $_.Name -match '\.bak$' }

if ($corruptProfiles) {
    Write-Host "Corrupt profile detected:" -ForegroundColor Red
    $corruptProfiles | Select-Object Name
} else {
    Write-Host "No corrupt profiles found in the registry." -ForegroundColor Green
}

# ---------------------- CHECK IF USER PROFILE EXISTS ON DISK ----------------------
Write-Host "`nChecking if user profile exists on disk..." -ForegroundColor Yellow
$profilesOnDisk = Get-ChildItem "C:\Users" | Select-Object Name
Write-Host "Existing User Profiles on Disk:" -ForegroundColor Cyan
$profilesOnDisk | Format-Table -AutoSize

# ---------------------- CHECK PROFILE FOLDER PERMISSIONS ----------------------
$profilePath = "C:\Users"
Write-Host "`nChecking profile folder permissions..." -ForegroundColor Yellow
$profilePermissions = Get-Acl $profilePath | Select-Object -ExpandProperty Access
$profilePermissions | Format-Table IdentityReference, FileSystemRights, AccessControlType -AutoSize

# ---------------------- CHECK GPO PROFILE PATH ----------------------
Write-Host "`nChecking Group Policy Profile Path (if configured)..." -ForegroundColor Yellow
$gpoProfilePath = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Policies\System" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "ProfilePath"
if ($gpoProfilePath) {
    Write-Host "Group Policy is redirecting profiles to: $gpoProfilePath" -ForegroundColor Red
} else {
    Write-Host "No GPO Profile Path configured." -ForegroundColor Green
}

# ---------------------- CHECK DISK FOR ERRORS ----------------------
Write-Host "`nChecking disk for errors (this may take a moment)..." -ForegroundColor Yellow
$diskErrors = Get-WinEvent -LogName System -MaxEvents 50 | Where-Object { $_.Id -in (7, 55, 98, 26226) }
if ($diskErrors) {
    Write-Host "Potential disk errors detected! Review the logs below:" -ForegroundColor Red
    $diskErrors | Select-Object TimeCreated, Id, Message | Format-Table -AutoSize
} else {
    Write-Host "No disk errors detected." -ForegroundColor Green
}

Write-Host "`nTroubleshooting completed. Review the results and proceed with fixes if needed." -ForegroundColor Cyan

# ---------------------- END OF SCRIPT ----------------------
try {
    # Try using 'quser' to get logged-in users
    $LoggedInUser = quser | ForEach-Object { ($_ -split '\s{2,}')[0].TrimStart('>', ' ') } | Where-Object { $_ -notmatch 'USERNAME' }
    if (-not $LoggedInUser) {
        throw "No users found with quser."
    }
} catch {
    try {
        # If 'quser' fails, fall back to 'Get-CimInstance'
        $LoggedInUser = (Get-CimInstance -ClassName Win32_ComputerSystem).UserName
        if (-not $LoggedInUser) {
            throw "No logged-in user found with CIM."
        }
    } catch {
        throw "Failed to retrieve logged-in user information using both methods."
    }
}

# Ensure username is sanitized for use in paths
if ($LoggedInUser -match '\\') {
    # If the username contains a domain (e.g., DOMAIN\User), split and extract the username
    $LoggedInUser = $LoggedInUser -replace '^.*\\', ''
}

# Construct user profile path
$UserProfilePath = Join-Path -Path "$env:SystemDrive\Users" -ChildPath $LoggedInUser

# Output results
if (Test-Path $UserProfilePath) {
    Write-Output "Logged-in User: $LoggedInUser"
    Write-Output "User Profile Path: $UserProfilePath"
} else {
    Write-Output "User profile path not found: $UserProfilePath"
}




   # Output the result
   	if ($LoggedInUser) {
    Write-Host "Logged-in User(s): $LoggedInUser" -ForeGroundColor Green
	} else {
    Write-Host "No logged-in users found." -ForeGroundColor Red
    }

$CheckFilesAttrib = Get-ChildItem -Path '*.*' -Force -ErrorAction SilentlyContinue | Format-Table Attributes
if ($CheckFilesAttrib -eq "5248544"){
	Write-Host "All the files attributes are already set to online only" -ForegroundColor Green
 	exit 1
  }else {

Start-Sleep -Seconds 3
# Resolve OneDrive path dynamically
$OneDrivePath = (Get-ChildItem -Path "$env:SystemDrive\Users\$LoggedInUser" -Directory | Where-Object { $_.Name -like "OneDrive - *" }).FullName
if (-not $OneDrivePath) {
    throw "OneDrive folder not found for user $LoggedInUser."
}
Set-Location -Path $OneDrivePath
Start-Sleep -Seconds 3

Start-Sleep -Seconds 5
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
Write-Host "Verifying the current file state" -ForeGroundColor Green
Write-Host ""

Write-Host "Enabling files on demand"
Get-childitem -Path '*.*' -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue |
Where-Object {$_.Attributes -match 'ReparsePoint' -or $_.Attributes -eq '525344' } |
ForEach-Object {
    attrib.exe $_.fullname +U /s
}

Write-Host "`n"
Write-Host "Verifying the updated file state" -ForeGroundColor Green
Write-Host ""
Get-ChildItem -Path '*.*' -Force -ErrorAction SilentlyContinue | Where-Object {$_.Attributes -eq '5248544' } Format-Table Attributes, Mode, Name, Length, CreationTime

}else {
    Write-Host "OneDrive folder not found for user $LoggedInUser." -ForeGroundColor Red
    exit 1
}




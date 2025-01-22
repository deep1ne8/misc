    try {
    # Try using 'quser'
    $LoggedInUser = quser | ForEach-Object { ($_ -split '\s{2,}')[0].TrimStart('>', ' ') } | Where-Object { $_ -notmatch 'USERNAME' }
    	if (-not $LoggedInUser) {
        	throw "No users found with quser."
    	}
    } catch {
    # If 'quser' fails, fall back to 'Get-WmiObject'
    $LoggedInUser = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
    	if (-not $LoggedInUser) {
        	throw "Failed to retrieve logged-in user information."
    	}
    }

   # Output the result
   	if ($LoggedInUser) {
    Write-Host "Logged-in User(s): $LoggedInUser" -ForeGroundColor Green
	} else {
    Write-Host "No logged-in users found." -ForeGroundColor Red
    }

$CheckFilesAttrib = Get-childitem -Path '*.*' -Force | Format-Table Attributes
if ($CheckFilesAttrib -eq "5248544"){
	Write-Host "All the files attributes are already set to online only"
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
Get-childitem -Path '*.*' -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue |
Where-Object {$_.Attributes -match 'ReparsePoint' -or $_.Attributes -eq '525344' } |
ForEach-Object {
    attrib.exe $_.fullname +U /s
}
Start-Sleep -Seconds 5
Write-Host "`n"
Write-Host "Below is the guide for the file state, according to it's attribute" -ForeGroundColor Green
Write-Host "`n==================================================================="
Write-Host "File State 	        Attribute" -ForeGroundColor White
Write-Host "-----------------------------"
Write-Host "Cloud-Only 	        5248544" -ForeGroundColor Green
Write-Host "Always available 	525344" -ForeGroundColor Green
Write-Host "Locally Available 	ReparsePoint" -ForeGroundColor Green
Write-Host "`n==================================================================="
Start-Sleep -Seconds 5
Write-Jost "`n"
Write-Host "Verifying the current file state" -ForeGroundColor Green
Write-Host ""
Write-Output $CheckFilesAttrib
}

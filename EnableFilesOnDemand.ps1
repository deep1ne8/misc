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
    Write-Host "Logged-in User(s): $LoggedInUser"
	} else {
    Write-Host "No logged-in users found."
    }

$CheckFilesAttrib = Get-childitem -Path "*.*" -Force | Format-Table Name,Attributes
Set-Location -Path "$env:SystemDrive\Users\$LoggedInUser\Onedrive - *\" 
Get-childitem -Path "*.*" -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue |
Where-Object {$_.Attributes -match 'ReparsePoint' -or $_.Attributes -eq '525344' } |
ForEach-Object {
    attrib.exe $_.fullname +U /s
}
Write-Host "Below is the guide for the file state, according to it's attribute" -ForeGroundColor Green
Write-Host "`n==================================================================="
Write-Host "File State 	        Attribute" -ForeGroundColor White
Write-Host "-----------------------------"
Write-Host "Cloud-Only 	        5248544" -ForeGroundColor Green
Write-Host "Always available 	525344" -ForeGroundColor Green
Write-Host "Locally Available 	ReparsePoint" -ForeGroundColor Green
Write-Host "`n==================================================================="

Write-Host "Verifying the current file state"
Write-Host ""
Write-Output $CheckFilesAttrib

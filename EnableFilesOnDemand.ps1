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

Get-childitem -Path "$env:SystemDrive\Users\$LoggedInUser\Onedrive - *\" -Force -File -Recurse -Verbose -ErrorAction SilentlyContinue |
Where-Object {$_.Attributes -match 'ReparsePoint' -or $_.Attributes -eq '525344' } |
ForEach-Object {
    attrib.exe $_.fullname +U -P /s
}

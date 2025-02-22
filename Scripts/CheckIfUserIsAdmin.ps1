# Define the target computer (use $env:COMPUTERNAME for local machine)
$ComputerName = $env:COMPUTERNAME

# Get the currently logged-in user from the remote computer
$User = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $ComputerName).UserName
Write-Host "Logged-in User: $User"

# If no user is logged in, exit
if (-not $User) {
    Write-Host "No user currently logged in."
    exit
}

# Extract only the username (removes domain/machine name)
$UserName = $User -replace '.*\\'

# Check if the logged-in user is in the local administrators group
if (Get-LocalGroupMember -Group "Administrators" -Name $UserName -ErrorAction SilentlyContinue) {
    Write-Host "Is Admin: YES"
} else {
    Write-Host "Is Admin: NO"
}


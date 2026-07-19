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

if (-not $UserName) {
    Write-Host "Is Admin: UNKNOWN (could not resolve username)"
    exit
}

# Check if the logged-in user is in the local Administrators group.
# Use -Member (not -Name, which collides with the group-name parameter).
$isAdmin = $false
try {
    $members = Get-LocalGroupMember -Group "Administrators" -ErrorAction Stop
    $isAdmin = ($members.Name -contains ".\$UserName") -or
               ($members.Name -contains "$env:COMPUTERNAME\$UserName") -or
               ($members.Name -contains $UserName)
} catch {
    Write-Host "Is Admin: UNKNOWN (could not read Administrators group: $($_.Exception.Message))"
    exit
}
Write-Host "Is Admin: $(if ($isAdmin) { 'YES' } else { 'NO' })"



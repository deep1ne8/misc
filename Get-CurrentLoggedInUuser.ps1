# Function to get currently logged-in user
function Get-LoggedOnUser {
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    try {
        # Query Win32_ComputerSystem to get the logged-in user
        $User = Get-CimInstance -ClassName Win32_ComputerSystem  | Select-Object UserName  
        
        if ($User) {
            Write-Host "$ComputerName : Logged in user - $User" -ForegroundColor Green
        } else {
            Write-Host "$ComputerName : No user currently logged in" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "$ComputerName : Error - $_" -ForegroundColor Red
    }
}

# Run for the local machine
Get-LoggedOnUser

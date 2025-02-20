# Function to get currently logged-in user
function Get-LoggedOnUser {
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    try {
        # Query Win32_ComputerSystem to get the logged-in user
        $User = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName -ErrorAction Stop | 
                Select-Object -ExpandProperty UserName
        
        if ($User) {
            Write-Output "$ComputerName : Logged in user - $User"
        } else {
            Write-Output "$ComputerName : No user currently logged in"
        }
    }
    catch {
        Write-Output "$ComputerName : Error - $_"
    }
}

# Run for the local machine
Get-LoggedOnUser

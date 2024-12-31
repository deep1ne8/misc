function Get-ExcelAddins {
    param (
        [string]$UserSID = $LoggedInUserSID # Optional: Specify a user SID; if null, it uses the current user
    )

    # Registry paths for Excel add-ins
    $ExcelAddinsPaths = @(
        "HKCU:\Software\Microsoft\Office\Excel\Addins",
        "HKCU:\Software\Microsoft\Office\16.0\Excel\Addins"
    )

    $loggedUser = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty UserName
    # Remove the domain or computer name prefix
    if ($loggedUser -match '\\') {
    $username = $loggedUser.Split('\')[1]
    } else {
    $username = $loggedUser
    }


    # Attempt to get the logged-in user's SID using CIM, fallback to WMI if it fails
    try {
    # Use CIM to retrieve the SID
    $LoggedInUserSID = (Get-CimInstance -ClassName Win32_UserAccount -Filter "LocalAccount=True" | Where-Object { $_.Name -eq $env:USERNAME }).SID
    if (-not $LoggedInUserSID) {
        throw "CIM returned null SID."
    }
    Write-Output "SID retrieved using CIM: $LoggedInUserSID"
    } catch {
    Write-Output "CIM method failed. Attempting WMI..."
    try {
        # Fallback to WMI
        $LoggedInUserSID = (Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount=True" | Where-Object { $_.Name -eq $env:USERNAME }).SID
        if (-not $LoggedInUserSID) {
            throw "WMI returned null SID."
        }
        Write-Output "SID retrieved using WMI: $LoggedInUserSID"
    } catch {
        # If both methods fail
        Write-Output "Failed to retrieve SID using both CIM and WMI. Error: $_"
        $LoggedInUserSID = $null
        }
    }

    $UserSID = $LoggedInUserSID

    # Log the user SID
    Write-Host "`n---------------------------`n"
    Write-Host "Logged-In User: $username" -ForegroundColor Green
    Write-Host "SID: $LoggedInUserSID" -ForegroundColor Yellow
    Write-Host "`n---------------------------`n"
    Start-Sleep 5

    # Check if the user SID is provided and exists
    if (-not $UserSID) {
        Write-Error "User SID not found. Please specify a user SID or use the current user." -ForegroundColor Red
        return
    }

    Write-Output "Retrieving Excel Add-ins..."

    # Loop through registry paths to find add-ins
    foreach ($Path in $ExcelAddinsPaths) {
        if (Test-Path $Path) {
            $Addins = Get-ChildItem -Path $Path
            foreach ($Addin in $Addins) {
                Write-Output "Add-in Name: $($Addin.PSChildName)"
                $Properties = Get-ItemProperty -Path $Addin.PSPath
                foreach ($Property in $Properties.PSObject.Properties) {
                    Write-Output "  $($Property.Name): $($Property.Value)"
                }
                Write-Output "`n---------------------------`n"
            }
        } else {
            Write-Output "Registry path not found: $Path"
        }
    }

    # Retrieve COM Add-ins via Excel Application
    try {
        $Excel = New-Object -ComObject Excel.Application
        $ComAddins = $Excel.ComAddIns
        Write-Output "COM Add-ins from Excel Application:"
        foreach ($ComAddin in $ComAddins) {
            Write-Output "  Name: $($ComAddin.Description)"
            Write-Output "  ProgID: $($ComAddin.ProgID)"
            Write-Output "  Connect: $($ComAddin.Connect)"
            Write-Output "`n---------------------------`n"
        }
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
    } catch {
        Write-Output "Error accessing Excel COM object: $_"
    }
}

# Execute the function
Get-ExcelAddins

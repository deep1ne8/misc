# Stop any running Teams processes
Get-Process -Name Teams -Verbose -ErrorAction SilentlyContinue | Stop-Process -Force -Verbose

# Detect logged-in users
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
    exit
}

# Define Teams packages to uninstall
$teamsPackages = @(
    "Microsoft Teams",
    "Teams Machine-Wide Installer"
)

foreach ($package in $teamsPackages) {
    Get-WmiObject -Query "SELECT * FROM Win32_Product WHERE Name = '$package'" -ErrorAction SilentlyContinue | ForEach-Object {
        $_.Uninstall()
    }
}

# Remove Teams from the user profile for all detected logged-in users
foreach ($user in $LoggedInUser) {
    $userProfilePath = "C:\Users\$user"
    $teamsPaths = @(
        "$userProfilePath\AppData\Roaming\Microsoft\Teams",
        "$userProfilePath\AppData\Local\Microsoft\Teams",
        "$userProfilePath\AppData\Local\Microsoft\TeamsMeetingAddin",
        "$userProfilePath\AppData\Local\Microsoft\TeamsPresenceAddin",
        "$userProfilePath\AppData\Local\Microsoft\TeamsAddin"
    )

    foreach ($path in $teamsPaths) {
        if (Test-Path -Path $path) {
            Remove-Item -Path $path -Recurse -Verbose -Force -ErrorAction SilentlyContinue
        }
    }

    # Remove Teams from startup for each user
    $teamsStartupPath = "$userProfilePath\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\Teams.lnk"
    if (Test-Path -Path $teamsStartupPath) {
        Remove-Item -Path $teamsStartupPath -Force -Verbose -ErrorAction SilentlyContinue
    }
}

# Remove Teams-related registry entries for the current machine
$teamsRegPaths = @(
    "HKCU:\Software\Microsoft\Office\Teams",
    "HKCU:\Software\Microsoft\Teams",
    "HKLM:\Software\Microsoft\Teams"
)

foreach ($regPath in $teamsRegPaths) {
    if (Test-Path -Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force -Verbose -ErrorAction SilentlyContinue
    }
}

Write-Host "Microsoft Teams and all cached settings have been removed for logged-in users: $LoggedInUser"

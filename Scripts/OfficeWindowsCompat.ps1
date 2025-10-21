# Office and Windows Compatibility Check
# Requires PowerShell 5.1 or higher

Write-Host "`n=== Office & Windows Compatibility Check ===`n"

# Get Windows Version
$os = Get-ComputerInfo | Select-Object WindowsProductName, OsVersion, OsBuildNumber
$winVersion = [version]$os.OsVersion
$build = [int]$os.OsBuildNumber

# Get Office Version
$officeRegPaths = @(
  "HKLM:\SOFTWARE\Microsoft\Office",
  "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office"
)
$officeVersion = $null

foreach ($path in $officeRegPaths) {
    $subKeys = Get-ChildItem -Path $path -ErrorAction SilentlyContinue | Where-Object { $_.Name -match "Office" }
    foreach ($key in $subKeys) {
        $ver = ($key.PSChildName -split "\\")[-1]
        if ($ver -match "^\d{2}\.\d$" -or $ver -match "^\d{2,4}$") {
            $officeVersion = $ver
        }
    }
}

if (-not $officeVersion) {
    Write-Host "❌ No Office installation detected."
    exit
}

# Interpret Office Version
switch -regex ($officeVersion) {
    "16"     { $officeName = "Office 2016 / 2019 / 2021 / Microsoft 365" }
    "15"     { $officeName = "Office 2013" }
    "14"     { $officeName = "Office 2010" }
    default  { $officeName = "Unknown Office build ($officeVersion)" }
}

# Compatibility Matrix (simplified)
# Reference: Microsoft lifecycle & compatibility docs
$compatible = $false
$reason = ""

if ($build -ge 22000) {
    # Windows 11
    if ($officeVersion -eq "16") { $compatible = $true; $reason = "Fully supported on Windows 11." }
    elseif ($officeVersion -eq "15") { $reason = "Office 2013 not officially supported on Windows 11."; }
    else { $reason = "Legacy Office versions may fail activation or updates on Windows 11."; }
}
elseif ($build -ge 19041 -and $build -lt 22000) {
    # Windows 10
    if ($officeVersion -ge 15) { $compatible = $true; $reason = "Supported on Windows 10." }
    else { $reason = "Office 2010 and older are unsupported and insecure."; }
}
else {
    $reason = "Older Windows build. Upgrade to Windows 10/11 for continued Office compatibility."
}

# Output Results
Write-Host "Detected Windows: $($os.WindowsProductName) ($($os.OsVersion), Build $build)"
Write-Host "Detected Office:  $officeName"
Write-Host ""
if ($compatible) {
    Write-Host "✅ COMPATIBLE — $reason"
} else {
    Write-Host "⚠️  NOT RECOMMENDED — $reason"
}

Write-Host "`n--- Best Practice ---"
Write-Host "• Keep Office and Windows both in mainstream support."
Write-Host "• Avoid mixing unsupported Office builds (2010/2013) with Windows 11."
Write-Host "• Use Microsoft 365 or Office LTSC 2021 with Windows 10/11 for full security and feature parity."
Write-Host "• Regularly check lifecycle at https://learn.microsoft.com/lifecycle"


# ----------------------------------------------
# CONFIG: GitHub raw URL to your PrinterCleanup.ps1
# ----------------------------------------------
$githubUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/script/PrinterCleanup.ps1"
$userProfilePath = "C:\Users\TMerchant"
$localScript = Join-Path $userProfilePath "PrinterCleanup.ps1"

Write-Host "=== Downloading PrinterCleanup.ps1 from GitHub... ===" -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $githubUrl -OutFile $localScript -UseBasicParsing -ErrorAction Stop
    Write-Host "Script downloaded successfully to $localScript" -ForegroundColor Green
} catch {
    Write-Host "Failed to download script: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# ----------------------------------------------
# Execute the cleanup script inside the user profile
# ----------------------------------------------
Write-Host "=== Executing PrinterCleanup.ps1 under $userProfilePath ===" -ForegroundColor Cyan

# IMPORTANT: If you need to dynamically pass the new printers, modify PrinterCleanup.ps1 to use these names
$printers = @(
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_Color",
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_B&W"
)

# Create a temp wrapper script to inject new printer names if needed
$wrapperScript = @"
`$printers = @(
    "`\\CVEFS02\S1_KM_C450i_Tammy's_Office_Color",
    "`\\CVEFS02\S1_KM_C450i_Tammy's_Office_B&W"
)
. '$localScript'
"@

$wrapperPath = Join-Path $userProfilePath "PrinterCleanup.ps1"
Set-Content -Path $wrapperPath -Value $wrapperScript -Force

# Execute wrapper script
try {
    powershell.exe -ExecutionPolicy Bypass -File $wrapperPath
} catch {
    Write-Host "Execution failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "=== Completed. Please test printing now. ===" -ForegroundColor Green

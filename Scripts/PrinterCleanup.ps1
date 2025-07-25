# ----------------------------------------------
# CONFIG
# ----------------------------------------------
$githubUrl = "https://raw.githubusercontent.com/deep1ne8/misc/refs/heads/main/Scripts/PrinterCleanup.ps1"
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
# Inject new printer names directly into the downloaded script
# (Assuming the script expects a $printers array at the top)
# ----------------------------------------------
$newPrintersBlock = @"
`$printers = @(
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_Color",
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_B&W"
)
"@

# Read existing script
$original = Get-Content $localScript
# Prepend new printers block
Set-Content -Path $localScript -Value ($newPrintersBlock + "`r`n" + ($original -join "`r`n"))

Write-Host "=== Executing PrinterCleanup.ps1 under $userProfilePath ===" -ForegroundColor Cyan
try {
    powershell.exe -ExecutionPolicy Bypass -File $localScript
} catch {
    Write-Host "Execution failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host "=== Completed. Please test printing now. ===" -ForegroundColor Green

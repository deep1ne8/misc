# ----------------------------------------------
# CONFIG
# ----------------------------------------------
<#
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
#>

# ----------------------------------------------
# Inject new printer names directly into the downloaded script
# (Assuming the script expects a $printers array at the top)
# ----------------------------------------------

$printers = @(
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_Color",
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_B&W"
)

Write-Host "=== Konica Minolta Printer Cleanup & Redeploy Script ===" -ForegroundColor Cyan

# ---------------------------
# REMOVE PRINTER QUEUES
# ---------------------------
Write-Host "`n[1/4] Removing existing Konica Minolta printers..." -ForegroundColor Yellow
Get-Printer | Where-Object {$_.Name -like "*Konica*"} | ForEach-Object {
    try {
        Write-Host "Removing printer: $($_.Name)"
        Remove-Printer -Name $_.Name -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Failed to remove printer: $($_.Name) - $($_.Exception.Message)"
    }
}

# ---------------------------
# REMOVE PRINTER DRIVERS
# ---------------------------
Write-Host "`n[2/4] Removing Konica Minolta drivers..." -ForegroundColor Yellow
Get-PrinterDriver | Where-Object {$_.Name -like "*Konica*"} | ForEach-Object {
    try {
        Write-Host "Removing driver: $($_.Name)"
        Remove-PrinterDriver -Name $_.Name -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Failed to remove driver: $($_.Name) - $($_.Exception.Message)"
    }
}

# ---------------------------
# REMOVE DRIVER PACKAGES (PnP)
# ---------------------------
Write-Host "`n[3/4] Removing driver packages (PnPutil)..." -ForegroundColor Yellow
$pnpDriversRaw = (pnputil /enum-drivers)
foreach ($line in $pnpDriversRaw) {
    if ($line -match "Published Name\s*:\s*(oem\d+\.inf)") {
        $inf = $matches[1]
        # Check if that block contains Konica
        $block = ($pnpDriversRaw -join "`n")
        if ($block -match "$inf" -and $block -match "(?i)Konica") {
            Write-Host "Force removing package: $inf"
            Start-Process -FilePath "pnputil.exe" -ArgumentList "/delete-driver $inf /uninstall /force" -NoNewWindow -Wait
        }
    }
}

# ---------------------------
# REDEPLOY PRINTERS
# ---------------------------
Write-Host "`n[4/4] Redeploying printers from server..." -ForegroundColor Yellow
foreach ($printer in $printers) {
    try {
        Write-Host "Adding printer: $printer"
        Add-Printer -ConnectionName $printer -ErrorAction Stop
    } catch {
        Write-Warning "Failed to add printer: $printer - $($_.Exception.Message)"
    }
}

Write-Host "`n=== Completed. Please test printing now. ===" -ForegroundColor Green
Pause

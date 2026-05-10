# ----------------------------------------------
# CONFIG: Printers to deploy
# ----------------------------------------------


$printers = @(
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_Color",
    "\\CVEFS02\S1_KM_C450i_Tammy's_Office_B&W"
)

Write-Host "=== Konica Minolta Printer Cleanup & Redeploy Script ===" -ForegroundColor Cyan

# Check for existing Konica Minolta printers
$IsPrinterInstalled = Get-Printer | Where-Object { $_.Name -like "*Konica*" }

if ($IsPrinterInstalled) {
    Write-Host "Konica Minolta printers detected. Starting cleanup..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3

    # ---------------------------
    # REMOVE PRINTER QUEUES
    # ---------------------------
    Write-Host "`n[1/4] Removing existing Konica Minolta printers..." -ForegroundColor Yellow
    Get-Printer | Where-Object { $_.Name -like "*Konica*" } | ForEach-Object {
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
    Get-PrinterDriver | Where-Object { $_.Name -like "*Konica*" } | ForEach-Object {
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
    exit 0

} else {
    Write-Host "No existing Konica Minolta printers found. Adding printers now..." -ForegroundColor Yellow
    Start-Sleep -Seconds 3

    # ---------------------------
    # REDEPLOY PRINTERS
    # ---------------------------
    Write-Host "`n[1/1] Redeploying printers from server..." -ForegroundColor Yellow
    foreach ($printer in $printers) {
        try {
            Write-Host "Adding printer: $printer"
            Add-Printer -ConnectionName $printer -ErrorAction Stop
        } catch {
            Write-Warning "Failed to add printer: $printer - $($_.Exception.Message)"
        }
    }

    Write-Host "`n=== Completed. Please test printing now. ===" -ForegroundColor Green
    exit 0
}

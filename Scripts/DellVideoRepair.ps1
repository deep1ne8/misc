#Requires -RunAsAdministrator

# Quick fix for DCU-CLI Error 106
Write-Host "Dell Command Update - Error 106 Fix" -ForegroundColor Cyan

# 1. Test connectivity
Write-Host "`n[1/5] Testing Dell server connectivity..." -ForegroundColor Yellow
$testDell = Test-NetConnection downloads.dell.com -Port 443 -WarningAction SilentlyContinue
if ($testDell.TcpTestSucceeded) {
    Write-Host "✓ Dell servers reachable" -ForegroundColor Green
} else {
    Write-Host "✗ Cannot reach Dell servers - Check firewall/proxy" -ForegroundColor Red
}

# 2. Reset DCU configuration
Write-Host "`n[2/5] Resetting DCU configuration..." -ForegroundColor Yellow
dcu-cli /configure -silent -autoSuspendBitLocker=enable -userConsent=disable
Write-Host "✓ Configuration reset" -ForegroundColor Green

# 3. Clear DCU cache
Write-Host "`n[3/5] Clearing DCU cache..." -ForegroundColor Yellow
$cachePath = "$env:ProgramData\Dell\CommandUpdate\cache"
if (Test-Path $cachePath) {
    Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "✓ Cache cleared" -ForegroundColor Green
} else {
    Write-Host "- No cache to clear" -ForegroundColor Gray
}

# 4. Check and restart DCU service
Write-Host "`n[4/5] Checking DCU service..." -ForegroundColor Yellow
$service = Get-Service -Name "DellClientManagementService" -ErrorAction SilentlyContinue
if ($service) {
    Restart-Service -Name "DellClientManagementService" -Force
    Write-Host "✓ Service restarted" -ForegroundColor Green
} else {
    Write-Host "- Service not found" -ForegroundColor Gray
}

# 5. Retry scan
Write-Host "`n[5/5] Retrying scan..." -ForegroundColor Yellow
dcu-cli /scan -silent

if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 500) {
    Write-Host "`n✓ SUCCESS! Error 106 resolved" -ForegroundColor Green
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "`nUpdates available. Run to install:" -ForegroundColor Cyan
        Write-Host "dcu-cli /applyUpdates -category=video -reboot=disable"
    } else {
        Write-Host "`nNo updates available." -ForegroundColor Green
    }
} else {
    Write-Host "`n✗ Still failing with exit code: $LASTEXITCODE" -ForegroundColor Red
    Write-Host "`nAdditional steps to try:" -ForegroundColor Yellow
    Write-Host "1. If behind proxy: dcu-cli /configure -proxy=your.proxy.com:port"
    Write-Host "2. Temporarily disable antivirus"
    Write-Host "3. Reinstall DCU from: https://www.dell.com/support/dcu"
    Write-Host "4. Use Dell SupportAssist instead"
}

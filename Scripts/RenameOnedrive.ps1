# Rename-OneDriveFolders.ps1
# This script modifies how OneDrive folders appear in File Explorer

# Run with administrative privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as Administrator!"
}

# Gather information
$userInitials = "JJ"
$companyName = Read-Host -Prompt "Enter your company name"
$personalDisplayName = "$userInitials - $companyName"
$businessDisplayName = "$userInitials - $companyName - Business"

Write-Host "This script will rename your OneDrive folders to '$personalDisplayName' and '$businessDisplayName'." -ForegroundColor Yellow
Write-Host "Note: You may need to restart Explorer or your computer for changes to fully take effect." -ForegroundColor Yellow
$confirmation = Read-Host "Do you want to continue? (Y/N)"

if ($confirmation -ne 'Y') {
    Write-Host "Operation cancelled." -ForegroundColor Red
    return
}

# Stop OneDrive process
Write-Host "Stopping OneDrive process..." -ForegroundColor Cyan
Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue

# Registry method for Personal OneDrive
try {
    Write-Host "Attempting to modify Personal OneDrive registry settings..." -ForegroundColor Cyan
    $oneDriveAccounts = Get-ChildItem -Path "HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts" -ErrorAction SilentlyContinue
    
    foreach ($account in $oneDriveAccounts) {
        $accountType = Get-ItemProperty -Path $account.PSPath -Name "AccountType" -ErrorAction SilentlyContinue
        
        if ($accountType -and $accountType.AccountType -eq 0) {
            # Personal account
            Write-Host "Found Personal OneDrive account, setting display name..." -ForegroundColor Green
            Set-ItemProperty -Path $account.PSPath -Name "UserName" -Value $personalDisplayName -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $account.PSPath -Name "DisplayName" -Value $personalDisplayName -ErrorAction SilentlyContinue
        }
        elseif ($accountType -and $accountType.AccountType -eq 1) {
            # Business account
            Write-Host "Found Business OneDrive account, setting display name..." -ForegroundColor Green
            Set-ItemProperty -Path $account.PSPath -Name "UserName" -Value $businessDisplayName -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $account.PSPath -Name "DisplayName" -Value $businessDisplayName -ErrorAction SilentlyContinue
        }
    }
}
catch {
    Write-Host "Error modifying registry settings: $_" -ForegroundColor Red
}

# Alternative method: Directory Junction
try {
    Write-Host "Creating directory junctions as an alternative..." -ForegroundColor Cyan
    
    # Determine OneDrive paths
    $personalOneDrivePath = [Environment]::GetFolderPath("UserProfile") + "\OneDrive"
    $businessOneDrivePath = [Environment]::GetFolderPath("UserProfile") + "\OneDrive - $companyName"
    
    # Check if the default paths exist
    if (Test-Path $personalOneDrivePath) {
        # Create junction for personal OneDrive
        $junctionPath = [Environment]::GetFolderPath("UserProfile") + "\$personalDisplayName"
        
        # Remove existing junction if it exists
        if (Test-Path $junctionPath) {
            # Check if it's a junction
            $folder = Get-Item $junctionPath -Force -ErrorAction SilentlyContinue
            if ($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                cmd /c "rmdir `"$junctionPath`""
            }
            else {
                Write-Host "Path $junctionPath exists and is not a junction. Skipping." -ForegroundColor Yellow
            }
        }
        
        if (-not (Test-Path $junctionPath)) {
            cmd /c "mklink /J `"$junctionPath`" `"$personalOneDrivePath`""
            Write-Host "Created junction for Personal OneDrive: $junctionPath" -ForegroundColor Green
        }
    }
    
    if (Test-Path $businessOneDrivePath) {
        # Create junction for business OneDrive
        $junctionPath = [Environment]::GetFolderPath("UserProfile") + "\$businessDisplayName"
        
        # Remove existing junction if it exists
        if (Test-Path $junctionPath) {
            # Check if it's a junction
            $folder = Get-Item $junctionPath -Force -ErrorAction SilentlyContinue
            if ($folder.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                cmd /c "rmdir `"$junctionPath`""
            }
            else {
                Write-Host "Path $junctionPath exists and is not a junction. Skipping." -ForegroundColor Yellow
            }
        }
        
        if (-not (Test-Path $junctionPath)) {
            cmd /c "mklink /J `"$junctionPath`" `"$businessOneDrivePath`""
            Write-Host "Created junction for Business OneDrive: $junctionPath" -ForegroundColor Green
        }
    }
}
catch {
    Write-Host "Error creating directory junctions: $_" -ForegroundColor Red
}

# Try Explorer shell folder method
try {
    Write-Host "Attempting to modify Explorer shell folders..." -ForegroundColor Cyan
    
    # For OneDrive Personal
    $clsidPersonal = Get-ChildItem -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace" -ErrorAction SilentlyContinue | 
        Where-Object { $_.PSChildName -eq "{018D5C66-4533-4307-9B53-224DE2ED1FE6}" }
    
    if ($clsidPersonal) {
        Set-ItemProperty -Path $clsidPersonal.PSPath -Name "(Default)" -Value $personalDisplayName -ErrorAction SilentlyContinue
        Write-Host "Modified Explorer shell folder for Personal OneDrive" -ForegroundColor Green
    }
}
catch {
    Write-Host "Error modifying Explorer shell folders: $_" -ForegroundColor Red
}

# Restart OneDrive
Write-Host "Starting OneDrive process..." -ForegroundColor Cyan
Start-Process "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDrive.exe"

Write-Host "`nScript complete!" -ForegroundColor Green
Write-Host "You may need to restart File Explorer or reboot your computer for all changes to take effect." -ForegroundColor Yellow
Write-Host "To restart File Explorer, you can run: Stop-Process -Name explorer -Force" -ForegroundColor Yellow

# Offer to restart Explorer
$restartExplorer = Read-Host "Would you like to restart Explorer now? (Y/N)"
if ($restartExplorer -eq 'Y') {
    Stop-Process -Name "explorer" -Force
    Start-Process "explorer"
    Write-Host "Explorer restarted." -ForegroundColor Green
}
return
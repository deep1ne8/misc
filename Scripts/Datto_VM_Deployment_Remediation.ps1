Write-Host "===== DATTO RMM =====" -ForegroundColor Cyan

Get-Service CagService -ErrorAction SilentlyContinue |
    Select-Object Name, Status, StartType

Write-Host "`n===== DATTO REGISTRY =====" -ForegroundColor Cyan

Get-ItemProperty "HKLM:\SOFTWARE\CentraStage" -ErrorAction SilentlyContinue |
    Select-Object DeviceID, AgentFolderStatus

Write-Host "`n===== DATTO FILES =====" -ForegroundColor Cyan

Get-ChildItem "C:\ProgramData\CentraStage\AEMAgent" -Force -ErrorAction SilentlyContinue |
    Select-Object Name, Length, LastWriteTime

Write-Host "`n===== VMWARE INFORMATION =====" -ForegroundColor Cyan

Get-CimInstance Win32_ComputerSystem |
    Select-Object Manufacturer, Model, Name, Domain

Get-CimInstance Win32_BIOS |
    Select-Object SerialNumber, SMBIOSBIOSVersion

Write-Host "`n===== COMPUTER UUID =====" -ForegroundColor Cyan

Get-CimInstance Win32_ComputerSystemProduct |
    Select-Object UUID, Vendor, Name, Version

Write-Host "`n===== WINDOWS INSTALLATION =====" -ForegroundColor Cyan

Get-CimInstance Win32_OperatingSystem |
    Select-Object InstallDate, LastBootUpTime, Caption, Version, BuildNumber

Write-Host "`n===== SYSTEM DRIVE =====" -ForegroundColor Cyan

Get-Volume -DriveLetter C |
    Select-Object DriveLetter, FileSystem, Size, SizeRemaining

Write-Host "`n===== DOMAIN =====" -ForegroundColor Cyan

(Get-CimInstance Win32_ComputerSystem).PartOfDomain

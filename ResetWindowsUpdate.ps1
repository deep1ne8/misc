# Function to stop a service with error handling
function Stop-ServiceWithRetry {
    param(
        [string[]]$serviceNames
    )
    foreach ($serviceName in $serviceNames) {
        try {
            Stop-Service -Name $serviceName -Force -ErrorAction Continue | Start-Sleep 2
            Write-Host "Service $serviceName stopped successfully." -ForegroundColor Green
            Start-Sleep 1
        } catch {
            Write-Host "Error stopping service ${serviceName}: $_" -ForegroundColor Red
            exit 1
        }
    }
}

# Function to start a service with error handling
function Start-ServiceWithRetry {
    param(
        [string[]]$serviceNames
    )
    foreach ($serviceName in $serviceNames) {
        try {
            Start-Service -Name $serviceName -ErrorAction Continue | Start-Sleep 2
            Write-Host "Service $serviceName started successfully." -ForegroundColor Green
            Start-Sleep 1
        } catch {
            Write-Host "Error starting service ${serviceName}: $_" -ForegroundColor Red
            exit 1
        }
    }
}

# Function to take ownership and rename an item with error handling
function Take-OwnershipAndRename {
    param(
        [string]$path
    )
    try {
        $acl = Get-Acl -Path $path
        $acl.SetOwner([System.Security.Principal.NTAccount]"BUILTIN\Administrators")
        Set-Acl -Path $path -AclObject $acl
        Write-Host "Took ownership of $path" -ForegroundColor Green

        $newName = [System.IO.Path]::GetFileNameWithoutExtension($path) + "_bak" + [System.IO.Path]::GetExtension($path)
        Rename-Item -Path $path -NewName $newName -Force
        Write-Host "Renamed $path to $newName" -ForegroundColor Green
    } catch {
        Write-Host "Error taking ownership or renaming ${path}: $_" -ForegroundColor Red
        exit 1
    }
}

# Step 1: Stop services
Write-Host "Step 1: Stopping services..." -ForegroundColor Cyan
$servicesToStop = @("wuauserv", "bits", "appidsvc", "cryptsvc")
foreach ($service in $servicesToStop) {
    Write-Host "Stopping service: $service"
    Stop-ServiceWithRetry -serviceNames $service
}

Start-Sleep 1

Write-Host ""
# Step 2: Check if services were stopped successfully
$servicesStopped = $servicesToStop | ForEach-Object { (Get-Service -Name $_).Status -eq 'Stopped' }
if ($servicesStopped -contains $false) {
    Write-Host "Not all services were stopped. Please check and try again." -ForegroundColor Red
    # exit 1
} else {
    Write-Host "All services stopped successfully. Continuing..." -ForegroundColor Green
}

Start-Sleep 1
Write-Host ""

# Scan the C: drive for errors
Write-Host "Scanning C drive for errors...." -ForegroundColor Cyan
Write-Host ""

# Perform the scan and store the results
$ScanResults = Repair-Volume -DriveLetter C -Scan -Verbose

# Check if there are any errors
if ($ScanResults.Errors.Count -gt 0) {
    Write-Host "Scanning for drive errors..." -ForegroundColor Cyan
    Write-Host "Errors found:" -ForegroundColor Red
    $ScanResults.Errors
} else {
    Write-Host "Scanning for drive errors..." -ForegroundColor Cyan
    Write-Host "No errors found." -ForegroundColor Green
}

Start-Sleep 1

Write-Host ""
# Step 3: Reset Windows Update components
Write-Host "Step 3: Resetting Windows Update components and removing SoftwareDistribution folder..." -ForegroundColor Cyan
Write-Host ""
Stop-ServiceWithRetry -serviceNames "wuauserv", "bits", "appidsvc", "cryptsvc"
Start-Sleep 2

Write-Host ""
# Step 4: Rename the software distribution folders
Write-Host "Step 4: Renaming the software distribution folders..." -ForegroundColor Cyan
$softwareDistributionPaths = @(
    "$env:SystemRoot\winsxs\pending.xml",
    "$env:SystemRoot\SoftwareDistribution",
    "$env:SystemRoot\system32\Catroot2",
    "$env:SystemRoot\WindowsUpdate.log"
)

foreach ($path in $softwareDistributionPaths) {
    if (Test-Path $path) {
        Take-OwnershipAndRename -path $path
    }
}

Start-Sleep 1

Write-Host ""
# Step 5: Reset the BITS service and the Windows Update service to the default security descriptor
Write-Host "Step 5: Resetting the BITS service and the Windows Update service to the default security descriptor..." -ForegroundColor Cyan
Start-Sleep 1

Write-Host ""
$securityDescriptorWuauserv = "D:(A;CI;CCLCSWRPLORC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)S:(AU;FA;CCDCLCSWRPWPDTLOSDRCWDWO;;;WD)"
$securityDescriptorBits = "D:(A;CI;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;SAFA;WDWO;;;BA)"
$securityDescriptorCryptsvc = "D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)(A;;CCLCSWRPWPDTLOCRRC;;;SO)(A;;CCLCSWLORC;;;AC)(A;;CCLCSWLORC;;;S-1-15-3-1024-3203351429-2120443784-2872670797-1918958302-2829055647-4275794519-765664414-2751773334)"
$securityDescriptorTrustedInstaller = "D:(A;CI;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)(A;;CCDCLCSWRPWPDTLOCRRC;;;BA)(A;;CCLCSWLOCRRC;;;IU)(A;;CCLCSWLOCRRC;;;SU)S:(AU;SAFA;WDWO;;;BA)"

Start-Sleep 1
Write-Host ""
sc.exe sdset wuauserv $securityDescriptorWuauserv
sc.exe sdset bits $securityDescriptorBits
sc.exe sdset cryptsvc $securityDescriptorCryptsvc
sc.exe sdset trustedinstaller $securityDescriptorTrustedInstaller

Start-Sleep 1
Write-Host ""
# Step 6: Reregister the BITS files and the Windows Update files
Write-Host "Step 6: Reregistering the BITS files and the Windows Update files..." -ForegroundColor Cyan
Write-Host ""
Start-Sleep 1

$bitsFiles = @("atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll", "browseui.dll", "jscript.dll", "vbscript.dll", "scrrun.dll", "msxml.dll", "msxml3.dll", "msxml6.dll", "actxprxy.dll", "softpub.dll", "wintrust.dll", "dssenh.dll", "rsaenh.dll", "gpkcsp.dll", "sccbase.dll", "slbcsp.dll", "cryptdlg.dll", "oleaut32.dll", "ole32.dll", "shell32.dll", "initpki.dll", "wuapi.dll", "wuaueng.dll", "wuaueng1.dll", "wucltui.dll", "wups.dll", "wups2.dll", "wuweb.dll", "qmgr.dll", "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll")

foreach ($file in $bitsFiles) {
    try {
        regsvr32.exe /s $file
        Write-Host "Registered $file successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error registering ${file}: $_" -ForegroundColor Red
        exit 1
    }
}

Start-Sleep 2

Write-Host ""
Write-Host "Removing SoftwareDistribution folder..." -ForegroundColor Cyan
Remove-Item -Path "$env:SystemRoot\SoftwareDistribution" -Recurse -Force -Verbose
Start-Sleep 2

Write-Host ""
Write-Host "Starting services..." -ForegroundColor Cyan
Start-ServiceWithRetry -serviceNames "wuauserv", "bits", "appidsvc", "cryptsvc"


Write-Host ""
Write-Host "Performing SFC scan..." -ForegroundColor Cyan
SFC /scannow
Start-Sleep 2

Write-Host ""
Write-Host "Performing DISM Restore Health..." -ForegroundColor Cyan
dism /online /cleanup-image /restorehealth
Start-Sleep 2

Write-Host ""
Write-Host "Performing Windows update..." -ForegroundColor Cyan
Write-Host ""
Set-ExecutionPolicy UnRestricted -Scope LocalMachine -Force
Write-Host ""
Start-Sleep 2
Install-Module -Name PSWindowsUpdate -Force -Verbose
Write-Host ""
Start-Sleep 2
Import-Module PSWindowsUpdate -Verbose
Write-Host ""
Start-Sleep 2
Get-WindowsUpdate -Install -AcceptAll -ForceDownload -Verbose
Write-Host ""
Start-Sleep 2
Write-Host "End of Script..." -ForegroundColor Cyan
Start-Sleep 2

exit

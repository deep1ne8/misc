function LanSweeper_GUI {
    param([string]$InstallerPath)

    Write-Host "`nStarting Lansweeper GUI installation..." -ForegroundColor Cyan
    Start-Process -FilePath $InstallerPath -Wait -Verb RunAs
}

function LanSweeper_Silent {
    param(
        [string]$InstallerPath,
        [string]$ConfigPath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        Write-Host "`nSilent-install configuration file was not found:" -ForegroundColor Red
        Write-Host $ConfigPath -ForegroundColor Yellow
        Write-Host "Download and configure lansweepernetworkdiscoveryinstallation.cfg from the Lansweeper portal first." -ForegroundColor Yellow
        return
    }

    Write-Host "`nStarting Lansweeper silent installation..." -ForegroundColor Cyan
    Start-Process -FilePath $InstallerPath `
        -ArgumentList '--optionfile', "`"$ConfigPath`"" `
        -Wait `
        -Verb RunAs
}

$DownloadUrl = 'https://download.lansweeper.com/stable/windows/network-discovery/latest'
$DownloadDir = Join-Path $env:TEMP 'Lansweeper'
$Installer   = Join-Path $DownloadDir 'LansweeperNetworkDiscovery.exe'
$ConfigFile  = 'C:\Install\lansweepernetworkdiscoveryinstallation.cfg'

do {
    Clear-Host
    Write-Host '========================================' -ForegroundColor DarkCyan
    Write-Host '     Lansweeper Network Discovery' -ForegroundColor Cyan
    Write-Host '========================================' -ForegroundColor DarkCyan
    Write-Host ''
    Write-Host '[1] GUI / Interactive install'
    Write-Host '[2] Silent install using CFG file'
    Write-Host '[Q] Quit'
    Write-Host ''

    $Choice = Read-Host 'Select installation type'

    switch ($Choice.ToUpper()) {
        '1' {
            New-Item -ItemType Directory -Path $DownloadDir -Force | Out-Null
            Write-Host "`nDownloading latest Lansweeper Network Discovery installer..." -ForegroundColor Cyan
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $Installer
            LanSweeper_GUI -InstallerPath $Installer
            break
        }

        '2' {
            New-Item -ItemType Directory -Path $DownloadDir -Force | Out-Null
            Write-Host "`nDownloading latest Lansweeper Network Discovery installer..." -ForegroundColor Cyan
            Invoke-WebRequest -Uri $DownloadUrl -OutFile $Installer
            LanSweeper_Silent -InstallerPath $Installer -ConfigPath $ConfigFile
            break
        }

        'Q' {
            Write-Host 'Installation cancelled.' -ForegroundColor Yellow
            break
        }

        default {
            Write-Host "`nInvalid selection. Press Enter and try again." -ForegroundColor Red
            Read-Host
        }
    }
}
while ($Choice.ToUpper() -notin @('1','2','Q'))
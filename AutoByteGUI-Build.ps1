<#
.SYNOPSIS
    Installs Python (if needed) and the GUI's pinned requirements, then starts AutoByteGUI.py.

.DESCRIPTION
    Uses an existing Python 3.11+ when present. Otherwise downloads the official python.org
    installer, verifies its Authenticode signature (Python Software Foundation), and installs
    it silently for all users. Requirements come from the requirements.txt next to this script
    (pinned), not from the moving main branch.

.PARAMETER PythonVersion
    python.org version to install when Python is missing.

.PARAMETER NoLaunch
    Install prerequisites only; do not start the GUI.

.EXAMPLE
    .\AutoByteGUI-Build.ps1

.NOTES
    Version: 2.0.0
    Exit codes: 0 success, 1 failure.
#>
#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidatePattern('^\d+\.\d+\.\d+$')][string]$PythonVersion = '3.12.10',
    [switch]$NoLaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

function Get-PythonExe {
    foreach ($candidate in 'py', 'python') {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if (-not $cmd) { continue }
        $prefix = @()
        if ($candidate -eq 'py') { $prefix = @('-3') }
        $ok = & $cmd.Source @prefix -c 'import sys; print(sys.version_info >= (3, 11))' 2>$null
        if ($LASTEXITCODE -eq 0 -and "$ok".Trim() -eq 'True') {
            return , ([string[]](@($cmd.Source) + $prefix))
        }
    }
    return $null
}

try {
    $python = Get-PythonExe
    if (-not $python) {
        $url = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
        $installer = Join-Path $env:TEMP "python-$PythonVersion-amd64.exe"
        if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Install Python $PythonVersion")) {
            Write-Host "Downloading $url" -ForegroundColor Cyan
            Invoke-WebRequest -Uri $url -OutFile $installer -UseBasicParsing
            $sig = Get-AuthenticodeSignature -FilePath $installer
            if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'Python Software Foundation') {
                throw "Python installer signature check failed ($($sig.Status); $($sig.SignerCertificate.Subject))."
            }
            $p = Start-Process -FilePath $installer -ArgumentList '/quiet', 'InstallAllUsers=1', 'PrependPath=1', 'Include_tcltk=1' -Wait -PassThru -NoNewWindow
            Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
            if ($p.ExitCode -notin 0, 3010) { throw "Python installer exited with $($p.ExitCode)." }
            $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
        }
        $python = Get-PythonExe
        if (-not $python) { throw 'Python 3.11+ still not found after install.' }
    }
    $exe = $python[0]; $pre = @($python | Select-Object -Skip 1)
    Write-Host "Using Python: $(& $exe @pre --version)" -ForegroundColor Green

    $req = Join-Path $PSScriptRoot 'requirements.txt'
    if (-not (Test-Path -LiteralPath $req)) { throw "requirements.txt not found next to this script ($req)." }
    & $exe @pre -m pip install --disable-pip-version-check -r $req
    if ($LASTEXITCODE -ne 0) { throw "pip install failed with exit code $LASTEXITCODE." }

    if (-not $NoLaunch) {
        Write-Host 'Starting AutoByte GUI...' -ForegroundColor Cyan
        & $exe @pre (Join-Path $PSScriptRoot 'AutoByteGUI.py')
    }
    exit 0
}
catch {
    Write-Error $_.Exception.Message -ErrorAction Continue
    exit 1
}

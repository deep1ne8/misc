<#
.SYNOPSIS
    Installs the Liongard agent (LTS MSI) silently.

.DESCRIPTION
    Downloads the Liongard LTS agent MSI over TLS 1.2+, installs it with msiexec,
    and cleans up. No credentials are stored in this script: pass them as parameters
    from your RMM's secure variables, or set the environment variables
    LIONGARD_URL, LIONGARD_ACCESS_KEY_ID, LIONGARD_ACCESS_KEY_SECRET and
    LIONGARD_ENVIRONMENT before running.

.PARAMETER LiongardUrl
    Liongard instance host, e.g. us5.app.liongard.com.

.PARAMETER AccessKeyId
    Liongard access key ID.

.PARAMETER AccessKeySecret
    Liongard access key secret. Prefer passing via LIONGARD_ACCESS_KEY_SECRET.

.PARAMETER Environment
    Liongard environment (client) name the agent registers into.

.PARAMETER InstallerUrl
    MSI download URL. Defaults to the official Liongard LTS agent.

.EXAMPLE
    $env:LIONGARD_ACCESS_KEY_SECRET = '<from RMM secret>'
    .\agent-installer.ps1 -LiongardUrl us5.app.liongard.com -AccessKeyId '<id>' -Environment 'Client Name'

.NOTES
    Version: 2.0.0
    Exit codes: 0 success / already installed, 1 failure, 3010 reboot required.
    Rotate any Liongard key that was previously committed to this repository.
#>
#Requires -Version 5.1
#Requires -RunAsAdministrator
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateNotNullOrEmpty()][string]$LiongardUrl = $env:LIONGARD_URL,
    [ValidateNotNullOrEmpty()][string]$AccessKeyId = $env:LIONGARD_ACCESS_KEY_ID,
    [ValidateNotNullOrEmpty()][string]$AccessKeySecret = $env:LIONGARD_ACCESS_KEY_SECRET,
    [ValidateNotNullOrEmpty()][string]$Environment = $env:LIONGARD_ENVIRONMENT,
    [ValidatePattern('^https://')][string]$InstallerUrl = 'https://agents.static.liongard.com/LiongardAgent-lts.msi',
    [switch]$Force
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

foreach ($name in 'LiongardUrl', 'AccessKeyId', 'AccessKeySecret', 'Environment') {
    if ([string]::IsNullOrWhiteSpace((Get-Variable -Name $name -ValueOnly))) {
        Write-Error "Missing required value '$name'. Pass -$name or set the matching LIONGARD_* environment variable."
        exit 1
    }
}

$installed = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
    Where-Object { $_.PSObject.Properties['DisplayName'] -and $_.DisplayName -like 'Liongard*Agent*' }
if ($installed -and -not $Force) {
    Write-Output "Liongard agent already installed ($(@($installed)[0].DisplayVersion)). Use -Force to reinstall."
    exit 0
}

$installerPath = Join-Path $env:TEMP 'LiongardAgent-lts.msi'
$logPath = Join-Path $env:TEMP 'liongard_install.log'
$exitCode = 1
try {
    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Install Liongard agent into environment '$Environment'")) { exit 0 }

    Write-Output 'Downloading Liongard agent installer...'
    Invoke-WebRequest -Uri $InstallerUrl -OutFile $installerPath -UseBasicParsing
    $sig = Get-AuthenticodeSignature -FilePath $installerPath
    if ($sig.Status -ne 'Valid') { throw "Installer signature is not valid ($($sig.Status))." }

    $arguments = @(
        '/i', "`"$installerPath`"",
        "LIONGARDURL=$LiongardUrl",
        "LIONGARDACCESSKEY=$AccessKeyId",
        "LIONGARDACCESSSECRET=$AccessKeySecret",
        "LIONGARDENVIRONMENT=`"$Environment`"",
        '/qn', '/norestart', '/L*v', "`"$logPath`""
    )
    Write-Output 'Installing Liongard agent...'
    $process = Start-Process -FilePath 'msiexec.exe' -ArgumentList $arguments -Wait -PassThru -NoNewWindow
    $exitCode = $process.ExitCode
    switch ($exitCode) {
        0 { Write-Output 'Installation completed successfully.' }
        3010 { Write-Output 'Installation completed; reboot required.' }
        default { throw "msiexec failed with exit code $exitCode." }
    }
}
catch {
    Write-Error "Liongard install failed: $($_.Exception.Message)" -ErrorAction Continue
    if (Test-Path -LiteralPath $logPath) {
        # The MSI log can echo properties; strip the secret before printing.
        Get-Content -LiteralPath $logPath -Tail 50 | ForEach-Object { $_.Replace($AccessKeySecret, '********') }
    }
    $exitCode = 1
}
finally {
    if (Test-Path -LiteralPath $installerPath) { Remove-Item -LiteralPath $installerPath -Force }
}
exit $exitCode

$guid = '{466038BB-646B-491C-A92E-051511D25C9F}'

# Try clean uninstall first, capture exit code
$p = Start-Process msiexec.exe -ArgumentList "/x $guid /qn /norestart" -Wait -PassThru -NoNewWindow
if ($p.ExitCode -eq 1612) {
    Write-Host "Source missing - forcing manual cleanup"
    Stop-Service -Name "LiongardAgent" -Force -ErrorAction SilentlyContinue
    Get-Process -Name "*Liongard*" -ErrorAction SilentlyContinue | Stop-Process -Force
    Remove-Item "C:\Program Files\Liongard*","C:\ProgramData\LionGard*" -Recurse -Force -ErrorAction SilentlyContinue
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$guid",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$guid"
    )
    $paths | ForEach-Object { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }
}
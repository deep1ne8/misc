@(set '(=)||' <# lean and mean cmd / powershell hybrid #> @'

::# Get 11 on 'unsupported' PC via Windows Update or mounted ISO (no patching needed)
::# Works for feature updates (23H2 to 24H2) as well as initial Windows 11 installation
::# V14: Enhanced reliability for 24H2, additional logging, cleanup improvements, and better compatibility with newer builds

@echo off & title Windows 11 Compatibility Bypass || Enhanced from AveYo's script
if /i "%~f0" neq "%SystemDrive%\Scripts\win11bypass.cmd" goto setup
powershell -win 1 -nop -c ";"
set CLI=%*& set SOURCES=%SystemDrive%\$WINDOWS.~BT\Sources& set MEDIA=.& set MOD=CLI& set PRE=WUA& set /a VER=11

:: Create log directory for troubleshooting
if not exist "%SystemDrive%\Scripts\logs" mkdir "%SystemDrive%\Scripts\logs" >nul 2>nul

:: Log start of execution
echo %date% %time% - Script started with arguments: %* > "%SystemDrive%\Scripts\logs\win11bypass.log"

if not defined CLI (exit /b) else if not exist %SOURCES%\SetupHost.exe (
    echo %date% %time% - ERROR: SetupHost.exe not found in %SOURCES% >> "%SystemDrive%\Scripts\logs\win11bypass.log"
    exit /b
)

:: Create WindowsUpdateBox.exe symlink if needed
if not exist %SOURCES%\WindowsUpdateBox.exe (
    echo %date% %time% - Creating WindowsUpdateBox.exe symlink >> "%SystemDrive%\Scripts\logs\win11bypass.log"
    mklink /h %SOURCES%\WindowsUpdateBox.exe %SOURCES%\SetupHost.exe
)

:: Apply registry bypasses for TPM and CPU compatibility
echo %date% %time% - Setting registry keys to bypass compatibility checks >> "%SystemDrive%\Scripts\logs\win11bypass.log"
reg add HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\TargetReleaseVersion /f /v TargetReleaseVersionInfo /d "24H2" /t reg_sz >nul 2>nul
reg add HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate /f /v DisableWUfBSafeguards /d 1 /t reg_dword
reg add HKLM\SYSTEM\Setup\MoSetup /f /v AllowUpgradesWithUnsupportedTPMorCPU /d 1 /t reg_dword
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassTPMCheck /d 1 /t reg_dword
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassSecureBootCheck /d 1 /t reg_dword
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassRAMCheck /d 1 /t reg_dword
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassStorageCheck /d 1 /t reg_dword
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassCPUCheck /d 1 /t reg_dword

:: Enhanced options for feature update reliability
set OPT=/Compat IgnoreWarning /MigrateDrivers All /Telemetry Disable /ShowOOBE None /Quiet

:: Handle restart application error code and other potential errors
set /a restart_application=0x800705BB & (call set CLI=%%CLI:%1 =%%)
set /a incorrect_parameter=0x80070057 & (set SRV=%CLI:/Product Client =%)
set /a launch_option_error=0xc190010a & (set SRV=%SRV:/Product Server =%)

:: Process CLI options
for %%W in (%CLI%) do if /i %%W == /PreDownload (set MOD=SRV)
for %%W in (%CLI%) do if /i %%W == /InstallFile (set PRE=ISO& set "MEDIA=") else if not defined MEDIA set "MEDIA=%%~dpW"

:: Version detection for Windows 11
if %VER% == 11 for %%W in ("%MEDIA%appraiserres.dll") do if exist %%W if %%~zW == 0 set AlreadyPatched=1 & set /a VER=10
if %VER% == 11 findstr /r "P.r.o.d.u.c.t.V.e.r.s.i.o.n...1.0.\..0.\..2.[2-9]" %SOURCES%\SetupHost.exe >nul 2>nul || set /a VER=10

:: Create EI.cfg if not present to prevent edition change during update
if %VER% == 11 if not exist "%MEDIA%EI.cfg" (
    echo %date% %time% - Creating EI.cfg to prevent edition change >> "%SystemDrive%\Scripts\logs\win11bypass.log"
    echo;[Channel]>%SOURCES%\EI.cfg & echo;_Default>>%SOURCES%\EI.cfg
)

:: ISO-specific handling for Windows 11
if %VER%_%PRE% == 11_ISO (
    echo %date% %time% - ISO mode detected, running pre-download >> "%SystemDrive%\Scripts\logs\win11bypass.log"
    %SOURCES%\WindowsUpdateBox.exe /Product Server /PreDownload /Quiet %OPT%
)

:: Handle appraiserres.dll for ISO mode
if %VER%_%PRE% == 11_ISO (
    echo %date% %time% - Handling appraiserres.dll for ISO mode >> "%SystemDrive%\Scripts\logs\win11bypass.log"
    del /f /q %SOURCES%\appraiserres.dll 2>nul & cd.>%SOURCES%\appraiserres.dll & call :canary
)

:: Set arguments based on mode
if %VER%_%MOD% == 11_SRV (set ARG=%OPT% %SRV% /Product Server)
if %VER%_%MOD% == 11_CLI (set ARG=%OPT% %CLI%)

:: Log final command execution
echo %date% %time% - Executing WindowsUpdateBox.exe with arguments: %ARG% >> "%SystemDrive%\Scripts\logs\win11bypass.log"

:: Execute WindowsUpdateBox with appropriate arguments
%SOURCES%\WindowsUpdateBox.exe %ARG%
set EXIT_CODE=%errorlevel%
echo %date% %time% - WindowsUpdateBox.exe exited with code: %EXIT_CODE% >> "%SystemDrive%\Scripts\logs\win11bypass.log"

:: Handle restart application error
if %EXIT_CODE% == %restart_application% (
    echo %date% %time% - Restart application error detected, applying canary fix and retrying >> "%SystemDrive%\Scripts\logs\win11bypass.log"
    call :canary
    %SOURCES%\WindowsUpdateBox.exe %ARG%
    set EXIT_CODE=%errorlevel%
    echo %date% %time% - Second WindowsUpdateBox.exe execution exited with code: %EXIT_CODE% >> "%SystemDrive%\Scripts\logs\win11bypass.log"
)

exit /b

:canary
:: Skip second TPM check by modifying hwreqchk.dll
echo %date% %time% - Applying canary fix to skip secondary TPM check >> "%SystemDrive%\Scripts\logs\win11bypass.log"
set C=  $X='%SOURCES%\hwreqchk.dll'; $Y='SQ_TpmVersion GTE 1'; $Z='SQ_TpmVersion GTE 0'; if (test-path $X) { 
set C=%C%  try { takeown.exe /f $X /a; icacls.exe $X /grant *S-1-5-32-544:f; attrib -R -S $X; [io.file]::OpenWrite($X).close() }
set C=%C%  catch { return }; $R=[Text.Encoding]::UTF8.GetBytes($Z); $l=$R.Length; $i=2; $w=!1;
set C=%C%  $B=[io.file]::ReadAllBytes($X); $H=[BitConverter]::ToString($B) -replace '-';
set C=%C%  $S=[BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes($Y)) -replace '-';
set C=%C%  do { $i=$H.IndexOf($S, $i + 2); if ($i -gt 0) { $w=!0; for ($k=0; $k -lt $l; $k++) { $B[$k + $i / 2]=$R[$k] } } }
set C=%C%  until ($i -lt 1); if ($w) { [io.file]::WriteAllBytes($X, $B); [GC]::Collect() } }
if %VER%_%PRE% == 11_ISO powershell -nop -c iex($env:C) >nul 2>nul
exit /b

:setup
::# elevate with native shell by AveYo
>nul reg add hkcu\software\classes\.Admin\shell\runas\command /f /ve /d "cmd /x /d /r set \"f0=%%2\"& call \"%%2\" %%3"& set _= %*
>nul fltmc|| if "%f0%" neq "%~f0" (cd.>"%temp%\runas.Admin" & start "%~n0" /high "%temp%\runas.Admin" "%~f0" "%_:"=""%" & exit /b)

::# lean xp+ color macros by AveYo
for /f "delims=:" %%s in ('echo;prompt $h$s$h:^|cmd /d') do set "|=%%s"&set ">>=\..\c nul&set /p s=%%s%%s%%s%%s%%s%%s%%s<nul&popd"
set "<=pushd "%appdata%"&2>nul findstr /c:\ /a" &set ">=%>>%&echo;" &set "|=%|:~0,1%" &set /p s=\<nul>"%appdata%\c"

::# toggle when launched without arguments, else jump to arguments: "install" or "remove"
set CLI=%*& (set IFEO=HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options)
wmic /namespace:"\\root\subscription" path __EventFilter where Name="Skip TPM Check on Dynamic Update" delete >nul 2>nul & rem legacy cleanup
reg delete "%IFEO%\vdsldr.exe" /f 2>nul & rem legacy cleanup

:: Create log directory for troubleshooting
if not exist "%SystemDrive%\Scripts\logs" mkdir "%SystemDrive%\Scripts\logs" >nul 2>nul

if /i "%CLI%"=="" (
    reg query "%IFEO%\SetupHost.exe\0" /v Debugger >nul 2>nul && goto remove || goto install
)
if /i "%~1"=="install" (goto install) else if /i "%~1"=="remove" goto remove

:install
echo %date% %time% - Installing Windows 11 bypass >> "%SystemDrive%\Scripts\logs\win11bypass.log"
mkdir %SystemDrive%\Scripts >nul 2>nul
copy /y "%~f0" "%SystemDrive%\Scripts\win11bypass.cmd" >nul 2>nul

:: Set the IFEO registry keys for Windows 11
echo %date% %time% - Setting IFEO registry keys >> "%SystemDrive%\Scripts\logs\win11bypass.log"
reg add "%IFEO%\SetupHost.exe" /f /v UseFilter /d 1 /t reg_dword >nul
reg add "%IFEO%\SetupHost.exe\0" /f /v FilterFullPath /d "%SystemDrive%\$WINDOWS.~BT\Sources\SetupHost.exe" >nul
reg add "%IFEO%\SetupHost.exe\0" /f /v Debugger /d "%SystemDrive%\Scripts\win11bypass.cmd" >nul

:: Set the latest BypassTPMCheck registry keys for Windows 11 24H2
echo %date% %time% - Setting additional bypass registry keys >> "%SystemDrive%\Scripts\logs\win11bypass.log"
reg add HKLM\SYSTEM\Setup\MoSetup /f /v AllowUpgradesWithUnsupportedTPMorCPU /d 1 /t reg_dword >nul
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassTPMCheck /d 1 /t reg_dword >nul
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassSecureBootCheck /d 1 /t reg_dword >nul
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassRAMCheck /d 1 /t reg_dword >nul
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassStorageCheck /d 1 /t reg_dword >nul
reg add HKLM\SYSTEM\Setup\LabConfig /f /v BypassCPUCheck /d 1 /t reg_dword >nul

echo;
%<%:f0 " Windows 11 TPM/CPU Compatibility Bypass "%>>% & %<%:2f " INSTALLED "%>>% & %<%:f0 " run again to remove "%>%
echo %date% %time% - Bypass installed successfully >> "%SystemDrive%\Scripts\logs\win11bypass.log"
if /i "%CLI%"=="" timeout /t 7
exit /b

:remove
echo %date% %time% - Removing Windows 11 bypass >> "%SystemDrive%\Scripts\logs\win11bypass.log"
del /f /q "%SystemDrive%\Scripts\win11bypass.cmd" >nul 2>nul
reg delete "%IFEO%\SetupHost.exe" /f >nul 2>nul
reg delete "HKLM\SYSTEM\Setup\LabConfig" /f >nul 2>nul
echo;
%<%:f0 " Windows 11 TPM/CPU Compatibility Bypass "%>>% & %<%:df " REMOVED "%>>% & %<%:f0 " run again to install "%>%
echo %date% %time% - Bypass removed successfully >> "%SystemDrive%\Scripts\logs\win11bypass.log"
if /i "%CLI%"=="" timeout /t 7
exit /b

'@); $0 = "$env:SystemDrive\Scripts\win11bypass.cmd"; ${(=)||} -split "\r?\n" | out-file $0 -encoding default -force; & $0
# Press Enter to continue
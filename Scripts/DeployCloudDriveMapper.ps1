@echo off
REM ============================================================================
REM Cloud Drive Mapper Reconfiguration - Batch Wrapper
REM ============================================================================
REM This batch file provides a simple wrapper to execute the PowerShell script
REM Usage: Run-CDM-Reconfiguration.cmd
REM Note: License key will be automatically detected from existing installation
REM       or you can specify it below if needed
REM ============================================================================

setlocal EnableDelayedExpansion

REM Configuration (Optional - leave empty to auto-detect from registry)
set "LICENSE_KEY="
set "SCRIPT_PATH=%~dp0Reconfigure-CloudDriveMapper.ps1"
set "LOG_FILE=%ProgramData%\CDM_BatchWrapper_%date:~-4,4%%date:~-10,2%%date:~-7,2%_%time:~0,2%%time:~3,2%%time:~6,2%.log"

REM Remove spaces from log filename
set "LOG_FILE=%LOG_FILE: =0%"

REM ============================================================================
REM Pre-flight Checks
REM ============================================================================

echo.
echo ============================================================
echo Cloud Drive Mapper Reconfiguration Wrapper
echo ============================================================
echo.

REM Check if running as Administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo [ERROR] This script must be run as Administrator!
    echo Right-click and select "Run as administrator"
    echo.
    pause
    exit /b 1
)

echo [OK] Running with Administrator privileges
echo.

REM Check if PowerShell script exists
if not exist "%SCRIPT_PATH%" (
    echo [ERROR] PowerShell script not found: %SCRIPT_PATH%
    echo.
    pause
    exit /b 1
)

echo [OK] PowerShell script found: %SCRIPT_PATH%
echo.

REM Check PowerShell version
for /f "tokens=*" %%i in ('powershell -NoProfile -Command "$PSVersionTable.PSVersion.Major"') do set PS_VERSION=%%i

if %PS_VERSION% LSS 5 (
    echo [ERROR] PowerShell 5.1 or higher required. Current version: %PS_VERSION%
    echo.
    pause
    exit /b 1
)

echo [OK] PowerShell version: %PS_VERSION%
echo.

REM ============================================================================
REM Execute Reconfiguration
REM ============================================================================

echo Starting Cloud Drive Mapper reconfiguration...
echo.
if defined LICENSE_KEY (
    echo License Key: %LICENSE_KEY%
) else (
    echo License Key: Auto-detect from registry
)
echo Log File: %LOG_FILE%
echo.
echo This may take 2-5 minutes. Please wait...
echo.

REM Execute PowerShell script with or without license key
if defined LICENSE_KEY (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%SCRIPT_PATH%' -LicenseKey '%LICENSE_KEY%'" > "%LOG_FILE%" 2>&1
) else (
    PowerShell.exe -NoProfile -ExecutionPolicy Bypass -Command "& '%SCRIPT_PATH%'" > "%LOG_FILE%" 2>&1
)

set EXITCODE=%ERRORLEVEL%

REM ============================================================================
REM Result Processing
REM ============================================================================

echo.
echo ============================================================

if %EXITCODE% EQU 0 (
    echo [SUCCESS] Cloud Drive Mapper reconfigured successfully!
    echo.
    echo Details:
    
    REM Get new drive letter from log
    for /f "tokens=3" %%a in ('findstr /C:"Drive Letter:" "%LOG_FILE%"') do set DRIVE_LETTER=%%a
    
    if defined DRIVE_LETTER (
        echo   - Drive Letter: %DRIVE_LETTER%:
    )
    
    echo   - Log File: %LOG_FILE%
    echo.
    
) else if %EXITCODE% EQU 3010 (
    echo [SUCCESS] Cloud Drive Mapper reconfigured successfully!
    echo [WARNING] System reboot required to complete installation
    echo.
    echo Log File: %LOG_FILE%
    echo.
    
    choice /C YN /M "Do you want to reboot now"
    if !errorlevel! EQU 1 (
        echo Rebooting in 30 seconds...
        shutdown /r /t 30 /c "Cloud Drive Mapper reconfiguration complete. System reboot required."
    )
    
) else (
    echo [FAILED] Reconfiguration failed with exit code: %EXITCODE%
    echo.
    echo Please check the log file for details:
    echo %LOG_FILE%
    echo.
    echo Common issues:
    echo   - Cloud Drive Mapper not installed on this machine
    echo   - Installer cache corrupted
    echo   - No available drive letters
    echo   - Insufficient permissions
    echo.
)

echo ============================================================
echo.

REM ============================================================================
REM Cleanup and Exit
REM ============================================================================

endlocal

if %EXITCODE% NEQ 0 (
    pause
)

exit /b %EXITCODE%

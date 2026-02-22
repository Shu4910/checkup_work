@echo off
setlocal enableextensions

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_SRC=%~dp0startup_launcher.vbs"
set "VBS_DST=%STARTUP_DIR%\checkup_work.vbs"

echo Registering checkup_work to startup folder...
echo.
echo Source: %VBS_SRC%
echo Dest:   %VBS_DST%
echo.

copy /Y "%VBS_SRC%" "%VBS_DST%"

if %errorlevel% == 0 (
    echo.
    echo [OK] Registration complete.
    echo Application will start automatically on next Windows login.
) else (
    echo.
    echo [ERROR] Registration failed.
    echo Please check that startup_launcher.vbs exists in the same folder.
)

echo.
pause

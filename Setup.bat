@echo off
setlocal enabledelayedexpansion
title ASHRAE 1925-TRP Experiment - Setup
cd /d "%~dp0"

echo ============================================================
echo   ASHRAE 1925-TRP Visual Experiment - One-time Setup
echo ============================================================
echo.
echo  This installs everything the experiment needs (Python
echo  packages). It is completely safe to run more than once.
echo.
echo  Press any key to begin...
pause >nul
echo.

REM ------------------------------------------------------------
REM  1) Find a working Python
REM     Prefer the "py" launcher (installed by python.org), then
REM     fall back to "python". The Microsoft Store stub fails
REM     these checks, so we won't accidentally use it.
REM ------------------------------------------------------------
set "PYCMD="
py -3 --version >nul 2>&1
if !errorlevel! equ 0 (
    set "PYCMD=py -3"
) else (
    python --version >nul 2>&1
    if !errorlevel! equ 0 set "PYCMD=python"
)

REM ------------------------------------------------------------
REM  2) If Python is missing, try to install it automatically
REM     with winget; otherwise open the download page.
REM ------------------------------------------------------------
if not defined PYCMD (
    echo  [!] Python was not found on this computer.
    echo.
    where winget >nul 2>&1
    if !errorlevel! equ 0 (
        echo  Trying to install Python automatically. Please approve
        echo  any pop-up that appears, then wait for it to finish...
        echo.
        winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
        echo.
        echo ============================================================
        echo   Python was installed.
        echo   IMPORTANT: close this window, then run Setup.bat AGAIN.
        echo ============================================================
        echo.
        pause
        exit /b 0
    ) else (
        echo  Automatic install is not available on this computer.
        echo.
        echo  Please install Python by hand:
        echo    1. A download page will open now.
        echo    2. Download the latest "Windows installer (64-bit)".
        echo    3. Run it and CHECK the box "Add python.exe to PATH".
        echo    4. Finish the install, then run Setup.bat again.
        echo.
        start "" "https://www.python.org/downloads/"
        pause
        exit /b 1
    )
)

echo  Found Python:
%PYCMD% --version
echo.

REM ------------------------------------------------------------
REM  3) Upgrade pip, then install the required packages.
REM ------------------------------------------------------------
echo  Updating the package installer (pip)...
%PYCMD% -m pip install --upgrade pip
echo.
echo  Installing required packages: openpyxl, pillow ...
%PYCMD% -m pip install openpyxl pillow
if !errorlevel! neq 0 (
    echo.
    echo  [!] Something went wrong installing the packages.
    echo      Check that you are connected to the internet and try
    echo      running Setup.bat again. If it keeps failing, send this
    echo      whole window to the study coordinator.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Setup complete!  You are ready to go.
echo.
echo   To start the experiment, double-click:
echo       Run Experiment.bat
echo ============================================================
echo.
pause
exit /b 0

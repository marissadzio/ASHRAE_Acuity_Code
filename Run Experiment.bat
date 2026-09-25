@echo off
setlocal enabledelayedexpansion
title ASHRAE 1925-TRP Visual Experiment
REM ============================================================
REM  ASHRAE 1925-TRP Visual Experiment - double-click launcher
REM  Runs the experiment. Choose "Practice (Demo)" on the start
REM  screen to rehearse without touching real data.
REM
REM  First time on a new laptop? Run Setup.bat once before this.
REM ============================================================
cd /d "%~dp0"

REM --- Find a working Python (prefer the "py" launcher) ---
set "PYCMD="
py -3 --version >nul 2>&1
if !errorlevel! equ 0 (
    set "PYCMD=py -3"
) else (
    python --version >nul 2>&1
    if !errorlevel! equ 0 set "PYCMD=python"
)

if not defined PYCMD (
    echo.
    echo  [!] Python was not found.
    echo      Please run Setup.bat first, then try again.
    echo.
    pause
    exit /b 1
)

%PYCMD% "ASHRAE_experiment_run.py"
if !errorlevel! neq 0 (
    echo.
    echo  The experiment exited with an error. Details are above.
    echo  If it mentions "openpyxl" or "PIL", run Setup.bat once,
    echo  then start the experiment again.
    echo.
    pause
)

@echo off
REM ============================================================
REM  ASHRAE 1925-TRP Visual Experiment - double-click launcher
REM  Runs the experiment. Choose "Practice (Demo)" on the start
REM  screen to rehearse without touching real data.
REM ============================================================
cd /d "%~dp0"
python "ASHRAE_experiment_run.py"
if errorlevel 1 (
  echo.
  echo The experiment exited with an error. Details are above.
  pause
)

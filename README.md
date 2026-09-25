# ASHRAE Acuity Code

Visual experiment application for ASHRAE 1925-TRP — a calibrated Tkinter app that runs
acuity, contrast, color-matching, and satisfaction-survey tasks for view-clarity research.

> 👉 **Setting this up on a Windows laptop? Start here:
> [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md)** — a step-by-step guide (download,
> install, run) written for non-technical users. Short version: double-click
> **`Setup.bat`** once, then double-click **`Run Experiment.bat`**.

## Files
- `ASHRAE_experiment_run.py` — runnable experiment application
- `ASHRAE experiment code.ipynb` — notebook version (kept in sync with the `.py`)
- `Setup.bat` — one-time Windows setup (installs Python packages, offers to install Python)
- `Run Experiment.bat` — double-click launcher (runs the `.py`)
- `SETUP_INSTRUCTIONS.md` — step-by-step guide for non-technical users (students)

## Running (Windows — for study operators / students)
See **`SETUP_INSTRUCTIONS.md`** for the full step-by-step guide. Short version:
1. Double-click **`Setup.bat`** once (installs everything needed).
2. Double-click **`Run Experiment.bat`** to start.

On the start screen, choose **Practice (Demo)** to rehearse without touching real data.

## Running (manual / other platforms)
```bash
python "ASHRAE_experiment_run.py"
```
Note: the app uses Windows-only display APIs; the Color stage and calibration are
Windows-specific.

## Notes
- Requires Python 3 with `openpyxl` installed (`pip install openpyxl`).
- The app is calibrated for a specific display (physical PPI and viewing distance);
  verify the calibration on your monitor before collecting real data.
- The per-participant design/sequence spreadsheet and results files are **not** included
  in this repo (they contain participant data).

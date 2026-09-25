# ASHRAE Acuity Code

Visual experiment application for ASHRAE 1925-TRP — a calibrated Tkinter app that runs
acuity, contrast, color-matching, and satisfaction-survey tasks for view-clarity research.

## Files
- `ASHRAE_experiment_run.py` — runnable experiment application
- `ASHRAE experiment code.ipynb` — notebook version (kept in sync with the `.py`)
- `Run Experiment.bat` — double-click launcher (runs the `.py`)

## Running
```bash
python "ASHRAE_experiment_run.py"
```
Or double-click `Run Experiment.bat`. On the start screen, choose **Practice (Demo)**
to rehearse without touching real data.

## Notes
- Requires Python 3 with `openpyxl` installed (`pip install openpyxl`).
- The app is calibrated for a specific display (physical PPI and viewing distance);
  verify the calibration on your monitor before collecting real data.
- The per-participant design/sequence spreadsheet and results files are **not** included
  in this repo (they contain participant data).

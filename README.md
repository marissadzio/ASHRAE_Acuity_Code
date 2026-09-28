# ASHRAE Acuity Code

Visual experiment application for ASHRAE 1925-TRP — a calibrated Tkinter app that runs
pre-session vision screening, then acuity, contrast, color-matching, and satisfaction-survey
tasks for view-clarity research.

> 👉 **Setting this up on a Windows laptop? Start here:
> [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md)** — a step-by-step guide (download,
> install, run) written for non-technical users. Short version: double-click
> **`Setup.bat`** once, add the Ishihara images (see below), then double-click
> **`Run Experiment.bat`**.

## Files
- `v2ASHRAE.py` — runnable experiment application (current version, v2)
- `v2ASHRAE.ipynb` — notebook version (kept in sync with the `.py`)
- `Run Experiment.bat` — double-click launcher (runs `v2ASHRAE.py`)
- `Setup.bat` — one-time Windows setup (installs Python packages, offers to install Python)
- `SETUP_INSTRUCTIONS.md` — step-by-step guide for non-technical users (students)
- `final_material_sequence_AorC.xlsx` — randomization / balanced design sequence (the
  per-participant `Full sequence` the app reads; participants are numeric IDs only)
- `ishihara/` — folder for the Ishihara color-vision plate images (**you must add these**)
- `backup/` — the previous version of the experiment (before v2), kept for reference

## Ishihara images (required)
Before a real session, the app runs a short screening: a 20/20 acuity check and two
Ishihara color-vision plates. The plate images are **not** included in the repo — put
them in the **`ishihara/`** folder, next to `v2ASHRAE.py`:

```
ASHRAE_Acuity_Code/
├── v2ASHRAE.py
├── Run Experiment.bat
└── ishihara/
    ├── Ishihara_23.png   ← correct answer: 42
    └── Ishihara_11.png   ← correct answer: 6
```

- Accepted extensions: `.png`, `.jpg`, `.jpeg`, `.bmp`.
- Keep the file names exactly as shown (`Ishihara_23`, `Ishihara_11`).
- Get the images from the study coordinator.

## Running (Windows — for study operators / students)
See **`SETUP_INSTRUCTIONS.md`** for the full step-by-step guide. Short version:
1. Double-click **`Setup.bat`** once (installs everything needed).
2. Put the two Ishihara images in the `ishihara/` folder.
3. Double-click **`Run Experiment.bat`** to start.

On the start screen, choose **Practice (Demo)** to rehearse without touching real data.

> **Windows in S Mode:** if double-clicking the `.bat` files does nothing, or Windows says it
> "only runs Microsoft-verified apps," the laptop is in S Mode, which blocks scripts. Switch
> out of S Mode (free, via Settings → System → Activation) and try again.

## Running (manual / other platforms)
```bash
python "v2ASHRAE.py"
```
Note: the app uses Windows-only display APIs; the Color stage, display switching, and
calibration are Windows-specific.

## Notes
- Requires Python 3 with `openpyxl` and `pillow` installed (`pip install openpyxl pillow`).
  `Setup.bat` installs both.
- On startup, v2 switches Windows to **Duplicate** display mode (same as Win+P → Duplicate).
- The app is calibrated for a specific display (physical PPI and viewing distance);
  verify the calibration on your monitor before collecting real data. Duplicate mode can
  change the external monitor's resolution, so re-check the calibration target size.
- The design spreadsheet and Ishihara images are found relative to the program's own
  folder, so keep them next to `v2ASHRAE.py` — don't move the `.py` on its own.
- The randomization/design sequence spreadsheet (`final_material_sequence_AorC.xlsx`) is
  included. Per-participant **results** files (`participant_*_results.xlsx`) are **not**
  committed — they hold collected data and are ignored by `.gitignore`.
- To run the old version in `backup/`, copy `final_material_sequence_AorC.xlsx` into
  `backup/` first (it looks for the spreadsheet in its own folder).

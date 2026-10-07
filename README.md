# ASHRAE Acuity Code

Visual experiment application for ASHRAE 1925-TRP — a calibrated Tkinter app that runs
pre-session vision screening, then acuity, contrast, color-matching, and satisfaction-survey
tasks for view-clarity research.

> 👉 **Setting this up on a Windows laptop? Start here:
> [SETUP_INSTRUCTIONS.md](SETUP_INSTRUCTIONS.md)** — a step-by-step guide (download,
> install, connect the TV, run) written for non-technical users. Short version: download the
> ZIP, extract it, double-click **`Setup.bat`** once, then double-click **`Run Experiment.bat`**.

## Files
- `ASHRAE-testUIs_v1-d.py` — the experiment program (current version)
- `ASHRAE-testUIs_v1-d.ipynb` — notebook version (kept identical to the `.py`)
- `Run Experiment.bat` — double-click launcher (runs `ASHRAE-testUIs_v1-d.py`)
- `Setup.bat` — one-time Windows setup (installs Python packages, offers to install Python)
- `SETUP_INSTRUCTIONS.md` — step-by-step guide for non-technical users (students)
- `final_material_sequence_AorC.xlsx` — randomization / balanced design sequence (the
  per-participant `Full sequence` the app reads; participants are numeric IDs only)
- `ishihara/` — the two Ishihara color-vision plate images used in the screening
- `backup/` — earlier versions, kept for reference (see below)

## The setup: laptop + SYLVOX 32" TV
The participant sits **3.048 m (10 ft)** from a **SYLVOX 32" 1080p outdoor TV**, which they
view through the glazing being tested. The researcher runs the program on a Windows laptop
connected to the TV by HDMI.

- **Screening, instructions, illuminance entry and the survey** use **Duplicate** mode (the
  laptop and TV show the same picture). The program switches to Duplicate by itself.
- **Acuity, contrast and color matching** use **Extend** mode, switched automatically when
  acuity starts and back to Duplicate before the survey:
  - **Acuity:** the Landolt C only on the TV; four answer arrows on the laptop.
  - **Contrast:** the target letters only on the TV; the typed letters on the laptop.
  - **Color matching:** the target color only on the TV; the color wheel on the laptop
    (the mouse is kept on the laptop).
- Popups remind the researcher to **hide the laptop screen** before each survey and to
  **put the laptop back on the table** after it.
- With only the laptop connected (e.g. practicing at home), everything runs on one screen.

Stimuli are drawn at their real-world size for this TV: 68.84 px per inch
(1920 × 1080 on a 32" diagonal) at 3.048 m. Before collecting data:
1. Set the TV's picture size to **"Just Scan" / "Screen Fit" / "1:1"** (no overscan).
2. Check with a ruler on the TV: the first acuity C (20/320) should be about **71 mm** across.

## Keyboard
| Key | What it does |
|---|---|
| Arrow keys | Acuity answer (gap direction); survey: move the selection |
| A–Z, Backspace, Enter | Contrast: type the letters, clear them, or submit fewer than 3 |
| Enter | Confirm / continue |
| **Alt+← / Alt+→** | Previous / next page (works on every screen) |
| F2 | Skip the current item |
| Esc | Skip the whole current task (asks for confirmation) |
| Exit ✕ / window close | End the session (saves first) |

## Running (Windows — for study operators / students)
See **`SETUP_INSTRUCTIONS.md`** for the full step-by-step guide. Short version:
1. Double-click **`Setup.bat`** once (installs everything needed).
2. Connect the TV, then double-click **`Run Experiment.bat`** to start.

On the start screen, choose **Practice (Demo)** to rehearse without touching real data.

> **Windows in S Mode:** if double-clicking the `.bat` files does nothing, or Windows says it
> "only runs Microsoft-verified apps," the laptop is in S Mode, which blocks scripts. Switch
> out of S Mode (free, via Settings → System → Activation) and try again.

## Running (manual / other platforms)
```bash
python "ASHRAE-testUIs_v1-d.py"
```
Note: the app uses Windows-only display APIs; display switching and calibration are
Windows-specific.

## Notes
- Requires Python 3 with `openpyxl` and `pillow` installed (`pip install openpyxl pillow`).
  `Setup.bat` installs both.
- The pre-session acuity screening is at **20/40** (its results go in the
  `screening_visual_acuity_20_20_*` columns).
- The sun position box accepts a number or `NA`.
- `TEST = True` at the top of the program shows a small debug box (test level, etc.) on the
  laptop. Set it to `False` for real sessions.
- The design spreadsheet and Ishihara images are found relative to the program's own
  folder, so keep them next to the `.py` — don't move the `.py` on its own.
- Per-participant **results** files (`participant_<ID>_results.xlsx`) are saved to the
  **Downloads** folder and are **not** committed (`.gitignore` blocks `*.xlsx` there).

## Earlier versions (`backup/`)
- `backup/testUIs_v1-c/` — two-screen contrast and color matching (acuity on one screen).
- `backup/testUIs_v1-b/` — two-screen color matching only.
- `backup/testUIs_v1/` — first TV version (calibration, Alt+arrows, cursor ring), one screen.
- `backup/v2ASHRAE/` — v2 for the 24" Dell monitor (before the TV).
- `backup/` (top level) — the original version before v2.

To run one of these, copy `final_material_sequence_AorC.xlsx` and the `ishihara/` folder
into that version's folder first (each looks for them next to its own `.py`).

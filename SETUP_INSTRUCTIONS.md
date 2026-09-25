# How to Set Up and Run the Experiment (Windows)

This guide is for **Windows laptops**. Follow it top to bottom. You only do
**Part 1 – 3 once** per laptop. After that, you just double-click one file to run
the experiment.

> ⏱️ First-time setup takes about 10–15 minutes, mostly waiting for downloads.
>
> 💡 If anything goes wrong, don't worry — take a photo/screenshot of the error
> and send it to the study coordinator.

---

## What you need

- A **Windows 10 or 11** laptop.
- An **internet connection** (for the one-time setup).
- About **15 minutes**.

---

## Part 1 — Get the experiment files onto your laptop

You have **two ways** to do this. **Option A (download a ZIP) is the easiest** and
is recommended if you're not comfortable with computers. Option B uses Git.

### Option A — Download as a ZIP (easiest, recommended)

1. Open this link in your web browser:
   **https://github.com/marissadzio/ASHRAE_Acuity_Code**
2. Click the green **`< > Code`** button near the top right.
3. In the little menu, click **Download ZIP**.
4. Open your **Downloads** folder and find `ASHRAE_Acuity_Code-main.zip`.
5. **Right-click** it → **Extract All…** → **Extract**.
6. You now have a folder called `ASHRAE_Acuity_Code-main`. Move it somewhere easy
   to find, like your **Desktop**. This folder is where everything lives.

> ✅ Done with Part 1. Skip to **Part 2**.

### Option B — Use Git (only if you were asked to)

1. **Install Git:**
   - Go to **https://git-scm.com/download/win** — the download starts
     automatically.
   - Run the downloaded installer. When it asks questions, just keep clicking
     **Next** to accept the defaults, then **Install**, then **Finish**.
2. **Download (clone) the project:**
   - Click the **Start** menu, type **`cmd`**, and open **Command Prompt**.
   - Copy and paste these two lines, pressing **Enter** after each one:
     ```
     cd %USERPROFILE%\Desktop
     git clone https://github.com/marissadzio/ASHRAE_Acuity_Code.git
     ```
   - This creates a folder called **`ASHRAE_Acuity_Code`** on your Desktop.

---

## Part 2 — Install everything the experiment needs (run once)

1. Open the project folder (from Part 1).
2. Find the file named **`Setup.bat`**.
3. **Double-click `Setup.bat`.**
   - A black window opens. Press a key to begin when it asks.
   - It will check for **Python** and install the required pieces automatically.
   - **If it says Python is missing:**
     - It will either install Python for you (approve any pop-up), **or** open the
       Python download page.
     - If the download page opens: download the **Windows installer (64-bit)**,
       run it, and on the very first screen **check the box that says
       "Add python.exe to PATH"** before clicking Install. This box is important!
     - After Python finishes installing, **double-click `Setup.bat` again.**
4. When you see **"Setup complete!"**, you're done. You can close the window.

> ℹ️ It's safe to run `Setup.bat` more than once. If you're ever unsure whether
> setup worked, just run it again.

---

## Part 3 — Run the experiment

1. In the project folder, find **`Run Experiment.bat`**.
2. **Double-click `Run Experiment.bat`.**
3. The experiment's start screen appears. Choose:
   - **Practice (Demo)** — to rehearse. Nothing is saved as real data.
   - **Start Real Session** — for an actual participant (needs the participant's
     ID and the study's design file; the coordinator sets this up).
4. Follow the on-screen prompts.

### The controls (shown on the start screen too)

| Stage      | What to press                                               |
|------------|------------------------------------------------------------|
| Acuity     | **Arrow keys** ← ↑ → ↓ = the direction of the gap          |
| Contrast   | **Type the 3 letters** (A–Z). **Backspace** clears them.   |
| Color      | **Click / drag** on the color wheel, then **Enter**        |
| Survey     | **← / →** to move the choice, then **Enter** to confirm    |
| Skip one   | **Tab** = skip just the current item                       |
| Skip task  | **Esc** = skip the whole current task (never quits)        |
| **Quit**   | Click the window's **✕ / Exit** button (this ends the run) |

### Where the results go

Real-session results are saved automatically to your **Downloads** folder as
`participant_<ID>_results.xlsx`.

---

## Troubleshooting

- **"Python was not found."** → Run `Setup.bat` and follow its instructions to
  install Python. Remember to check **"Add python.exe to PATH"** during the Python
  install.
- **An error mentions `openpyxl` or `PIL`/`pillow`.** → Run `Setup.bat` once more,
  then start the experiment again.
- **Double-clicking a `.bat` file does nothing / flashes and closes.** →
  Right-click it → **Run as administrator**. If it still closes instantly, right-
  click → **Edit** is *not* what you want; instead take a photo of any message and
  send it to the coordinator.
- **Windows SmartScreen says "Windows protected your PC."** → Click **More info**
  → **Run anyway**. (These are simple, safe scripts from your study.)
- **The on-screen sizes look wrong for a real session.** → The experiment is
  calibrated for a specific monitor and viewing distance. Only collect real data
  on the coordinator-approved setup. Demo mode is fine on any laptop.

---

## Quick reference (after the first-time setup)

Every time you want to run the experiment, you only do **one thing**:

> **Double-click `Run Experiment.bat`.**

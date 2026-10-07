# How to Set Up and Run the Experiment (Windows)

This guide is for **Windows laptops**. Follow it top to bottom. You only do
**Parts 1 – 2 once** per laptop. After that, you just double-click one file to run
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

## Part 3 — Connect the TV (each session)

The participant sits **10 ft (3.048 m)** from the **SYLVOX 32" TV** and looks at it
through the window being tested. The laptop connects to the TV with an **HDMI cable**.

1. Plug the HDMI cable into the laptop and the TV, and set the TV to that HDMI input.
2. On the TV's remote, set the **picture size** to **"Just Scan"**, **"Screen Fit"** or
   **"1:1"** (not "Zoom" or "16:9 wide"). Otherwise the TV cuts off the edges and every
   size is slightly too big.
3. That's all: the program switches the screens by itself (you don't need Win+P).

> 💡 **What you'll see:** the laptop and TV show the **same picture** for the instructions,
> the Ishihara plates and the survey. During the **20/40 eye screening, acuity, contrast and
> color matching** the screens **split**: the test is shown **only on the TV** and the answer screen (arrows, typed
> letters, color wheel) **only on the laptop**. Switching takes a few seconds and the
> screens may flicker; that's normal.
>
> ✅ The Ishihara plate images used in the screening already come with the download
> (in the `ishihara` folder). Don't rename or move them.

---

## Part 4 — Run the experiment

1. Connect the TV first (Part 3).
2. In the project folder, **double-click `Run Experiment.bat`**.
3. The experiment's start screen appears. Choose:
   - **Practice (Demo)** — to rehearse. Nothing is saved as real data.
   - **Start Real Session** — for an actual participant (needs the participant's
     ID; the coordinator gives you this).
4. Follow the on-screen prompts. Popups will remind you to **hide the laptop screen**
   before each survey and to **put the laptop back on the table** after it.

### The controls (shown on the start screen too)

| Stage      | What to press                                                        |
|------------|---------------------------------------------------------------------|
| Acuity     | **Arrow keys** ← ↑ → ↓ = the direction of the gap                   |
| Contrast   | **Type the 3 letters** (A–Z). **Backspace** clears them. **Enter** submits if fewer than 3 are visible. |
| Color      | **Click / drag** on the color wheel, then **Enter**                 |
| Survey     | **← / →** to move the choice, then **Enter** to confirm             |
| Pages      | **Alt+←** = previous page, **Alt+→** = next page (any screen)       |
| Skip one   | **F2** = skip just the current item                                 |
| Skip task  | **Esc** = skip the whole current task (asks first; never quits)     |
| **Quit**   | Click the window's **✕ / Exit** button (saves, then ends the run)   |

The **Sun position** box accepts a number or **NA**.

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
- **Windows says it "only runs Microsoft-verified apps," or `.bat` files do
  absolutely nothing.** → The laptop is in **S Mode**, which blocks these files.
  Switch out of S Mode (it's free): **Settings → System → Activation → Switch out of
  S mode** (opens the Microsoft Store) → **Get**. Then try again.
- **The Ishihara screening shows a blank or missing picture.** → The images were
  moved or renamed. Both files must be inside the `ishihara` folder, named exactly
  `Ishihara_23.png` and `Ishihara_11.png`. Re-download the project if they're missing.
- **The screens don't split during the tests (everything stays on the laptop).** →
  Check the HDMI cable and that the TV is on the right input, then restart the
  experiment. Without a TV connected, the program runs everything on one screen.
- **Windows SmartScreen says "Windows protected your PC."** → Click **More info**
  → **Run anyway**. (These are simple, safe scripts from your study.)
- **The on-screen sizes look wrong for a real session.** → The experiment is
  calibrated for the SYLVOX 32" TV at 10 ft. Check the TV's picture size setting
  (Part 3). With a ruler, the first acuity "C" on the TV should be about **71 mm**
  across. Only collect real data on the coordinator-approved setup.

---

## Quick reference (after the first-time setup)

Every time you want to run the experiment, you only do **one thing**:

> **Double-click `Run Experiment.bat`.**

import tkinter as tk
import tkinter.font as tkfont
from tkinter import messagebox, ttk
import sys
import math
import ctypes
from ctypes import wintypes
from pathlib import Path
from datetime import datetime
import random
import re
import subprocess
import time

from openpyxl import Workbook, load_workbook

# --- Optional (for color wheel + PIL letter render) ---
try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageOps, ImageFilter
    PIL_OK = True
except Exception:
    PIL_OK = False
    Image = None
    ImageTk = None
    ImageDraw = None
    ImageFont = None
    ImageOps = None
    ImageFilter = None


# Folder this program lives in. Data files (design spreadsheet, Ishihara images)
# are looked up relative to it, so the project works on any machine without
# hard-coded C:\Users\... paths. In a notebook __file__ is undefined, so fall
# back to the working directory (the notebook's folder).
try:
    PROGRAM_DIR = Path(__file__).resolve().parent
except NameError:
    PROGRAM_DIR = Path.cwd()


# =========================================================
# ========================= USER SETTINGS =================
# =========================================================

# ✅ Toggle for debugging annotations
TEST = True  # True => show debug overlay (acuity 20/den, contrast hex/logCS, etc.)

# ===== Pre-session screening =====
# Pre-session acuity screening: two Landolt-C presentations at 20/SCREENING_ACUITY_DENOMINATOR.
# (Results still go in the screening_visual_acuity_20_20_* columns: renaming a column would make
# the program rebuild, i.e. empty, an existing participant results file.)
SCREENING_ACUITY_DENOMINATOR = 40
SCREENING_ACUITY_TRIALS = 2
SCREENING_ACUITY_PASS_N = 2
SCREENING_ORIENTATIONS = ["up", "down", "left", "right"]
SCREENING_KEYSYM_TO_ORIENTATION = {"Up": "up", "Down": "down", "Left": "left", "Right": "right"}

# Ishihara plate image folder: the "ishihara" folder next to this program.
ISHIHARA_DIR = PROGRAM_DIR / "ishihara"
# The program accepts these base names with .png, .jpg, .jpeg, or .bmp extensions.
ISHIHARA_PLATES = [
    {"filename": "Ishihara_23", "expected_answer": "42"},
    {"filename": "Ishihara_11", "expected_answer": "6"},
]

# =========================================================
# ======================= HOTKEY MAP ======================
# =========================================================
# Response / control keys used during the test battery:
#   Acuity   : Arrow keys (Up / Down / Left / Right) = gap direction (answer)
#   Contrast : Letters A-Z = type the 3 Sloan letters ; Backspace / Delete = clear
#   Color    : Mouse click / drag = pick color ; Enter = confirm
#   Survey   : Left / Right = move selection ; Enter = confirm answer
#   All      : Esc = skip the whole current task and move on (NEVER quits the run).
#              The run ends ONLY via a window's close (X) button, an on-screen
#              "Exit" button, or Alt+F4 (all call save_and_exit).
#
# SKIP (accessibility, e.g. for participants who cannot see):
#   F2 = skip the CURRENT task and move on to the next task.
#     - Works in every stage: acuity -> contrast -> color -> survey -> next condition.
#     - Records the skipped task as "SKIPPED" (NOT the same as a wrong answer / Esc fail).
#     - Tab is reserved for keyboard focus navigation between selectable controls.
SKIP_KEYSYM = "F2"  # Tab is reserved for moving keyboard focus between selectable controls

VIEWING_DISTANCE_M = 3.048  # 10 ft back (was 1.0). Use 3.0 if you mark the distance in metric.

# ===== Monitor physical-size calibration (SYLVOX 32" outdoor TV, 1920x1080) =====
# The participant display is a 32" 16:9 1080p TV. Its true physical pixel density is
#   sqrt(1920^2 + 1080^2) / 32 = 2202.9 / 32 = 68.84 px per inch,
# so the stimuli (Landolt-C rings, contrast letters) are drawn with that PPI to come out
# at their intended real-world mm size at VIEWING_DISTANCE_M.
# Verify with a ruler on the TV: the 20/320 Landolt-C should be ~71 mm across and the
# whole picture ~708 mm wide x ~398 mm tall. Set the TV's picture size to "Just Scan" /
# "Screen Fit" / "1:1" (no overscan), otherwise every size is a few percent too large.
USE_PHYSICAL_PPI = True
MONITOR_DIAGONAL_IN = 32.0
MONITOR_RESOLUTION_PX = (1920, 1080)
MONITOR_PHYSICAL_PPI = math.hypot(*MONITOR_RESOLUTION_PX) / MONITOR_DIAGONAL_IN  # ~68.84

# Popups, the survey scale and other on-screen layout keep the exact pixel layout they
# were designed and checked with on a 1920x1080 screen, which used 92 px per inch.
# Only the calibrated stimuli above use the TV's physical PPI.
LAYOUT_DPI = 92.0
POLL_MS = 250

# ===== Contrast letter shape controls (PIL render) =====
# NOTE: use 1.0 for normal; >1.0 stretches. (Your snippet had 10/10 which will explode.)
LETTER_WIDTH_SCALE = 1.00
LETTER_HEIGHT_SCALE = 1.00
LETTER_WEIGHT_PX = 0  # 0=off, 1..3 => thicker strokes

# ✅ Acuity levels (Landolt-C)
ACUITY_LEVELS = [320, 160, 80, 40, 20]
LAST_LEVEL = ACUITY_LEVELS[-1]
ACUITY_PRE_STIMULUS_DELAY_MS = 500  # blank screen before each Landolt-C appears

# ✅ Contrast follow-up (Pelli-style triplets)
# Contrast letters: hit a target on-screen LETTER height (ruler-measured).
# The Sloan glyph renders ~0.83 of its box, so size the box from that fill ratio.
# (20/680 at 3.048 m is ~150 mm overall; measured glyph 125 mm in a 150.7 mm box -> 0.829.)
CONTRAST_TARGET_LETTER_MM = 148.0
CONTRAST_LETTER_FILL = 0.829
CONTRAST_BOX_MM = CONTRAST_TARGET_LETTER_MM / CONTRAST_LETTER_FILL
SHOW_CONTRAST_BOX = False
SLOAN_LETTERS = list("CDHKNORSVZ")

# Lowest contrast level to present. Valid triplet indices are 0 through 7.
# After the participant passes triplet_idx = 7, the contrast test ends and
# proceeds to the color-matching stage; triplet_idx >= 8 is never displayed.
MAX_CONTRAST_TRIPLET_IDX = 7

# ✅ Fixed grayscale sequence (triplet 1..N)
CONTRAST_HEX_SEQUENCE = [
    "#000000",  # Black (high contrast)
    "#939393",
    "#BBBBBB",
    "#D2D2D2",
    "#E0E0E0",
    "#EAEAEA",
    "#F1F1F1",
    "#F5F5F5",
    "#F8F8F8",
    "#FAFAFA",
    "#FCFCFC",
    "#FDFDFD",
]

# ✅ log(CS) mapping per triplet step (same length as HEX sequence)
CONTRAST_LOGCS_BY_TRIPLET = [
    0.00,
    0.15,
    0.30,
    0.45,
    0.60,
    0.75,
    0.90,
    1.05,
    1.20,
    1.35,
    1.50,
    1.65,
]

# Triplet rule: if ≥2 correct out of 3, advance to next triplet
MAX_LETTER_INDEX = 48  # 16 triplets * 3 letters = 48

# ✅ Color matching stage settings (chip + wheel sizes are computed at runtime
# from the monitor in start_color_match; COLOR_WHEEL_SIZE_PX is the fallback/seed).
COLOR_WHEEL_SIZE_PX = 320

# Wheel marker
WHEEL_MARKER_RADIUS_PX = 32
WHEEL_MARKER_STROKE_PX = 6

# Color matching on two screens: the target color is shown ONLY on the TV (seen through
# the glazing) and the color wheel ONLY on the laptop. Duplicate mode shows the same
# picture on both, so Windows is switched to Extend for the color-matching task and
# back to Duplicate afterwards. If only one screen is connected (e.g. a practice run on
# the laptop alone), both color windows share that screen as before.
COLOR_MATCH_TWO_SCREENS = True
# Contrast on two screens: the target letters ONLY on the TV, the typed letters ONLY on the
# laptop. Windows stays in Extend from contrast through color matching.
CONTRAST_TWO_SCREENS = True
# Acuity on two screens: the Landolt C ONLY on the TV, four answer arrows on the laptop.
# Windows switches to Extend when acuity starts and stays there through color matching.
ACUITY_TWO_SCREENS = True
DISPLAY_SWITCH_TIMEOUT_MS = 8000   # give up waiting for Windows after this long
DISPLAY_SWITCH_SETTLE_MS = 700     # extra wait after the screens report the new layout

# ✅ Survey stage settings
SURVEY_VALUE_TO_TEXT = {
    1: "Very dissatisfied",
    2: "Dissatisfied",
    3: "Slightly dissatisfied",
    4: "Neutral",
    5: "Slightly satisfied",
    6: "Satisfied",
    7: "Very satisfied",
}

# Survey questions: (label shown to the participant, results column)
SURVEY_QUESTIONS = [
    ("Clarity of view", "survey_clarity_of_view_7pt"),
    ("Visual privacy", "survey_visual_privacy_7pt"),
    ("Reflections / mirror-effect", "survey_reflections_mirror_effect_7pt"),
    ("Visual comfort (glare)", "survey_visual_comfort_glare_7pt"),
]

DOMAIN = "@berkeley.edu"
DOWNLOADS_DIR = Path.home() / "Downloads"
# XLSX_PATH is set per-participant at startup (participant_<id>_results.xlsx).
SHEET = "participants"

# ===== Experiment design file =====
# The design/sequence spreadsheet must live in the SAME folder as this program.
# It ships with the project download, right next to this .py and the .bat files.
DESIGN_XLSX_NAME = "final_material_sequence_AorC.xlsx"
DESIGN_XLSX_PATH = PROGRAM_DIR / DESIGN_XLSX_NAME
DESIGN_SHEET = "Full sequence"
DESIGN_PARTICIPANT_COL = "Participant"
DESIGN_SEQUENCE_COL = "Full sequence"
DONE_GRAY = "#ADADC9"
CURRENT_YELLOW = "#FFF2CC"

# UI colors
BG_SOFT = "#FAFAFA"
ARROW_BG = "white"
ARROW_HILITE = "#ff9aa2"
QUIT_BG = "#d9d9d9"

# Fonts
QUIT_FONT = ("Segoe UI", 18, "bold")
ARROW_FONT = ("Segoe UI", 22, "bold")
ARROW_TITLE_FONT = ("Segoe UI", 14, "bold")

HUD_TITLE_FONT = ("Segoe UI", 12, "bold")
HUD_TEXT_FONT = ("Segoe UI", 11)
HUD_INSTR_FONT = ("Segoe UI", 12)

# Instruction line shown at the top of the main window during each test.
ACUITY_HUD = ("Press the arrow key for the gap direction. (F2 = skip this task; Tab = move focus; "
              "ESC = confirm skip; Alt+← / Alt+→ = previous / next page)")
COLOR_HUD = ("Click/drag on the wheel to pick the closest color. Press Enter when done. "
             "(F2 = skip this task; Tab = move focus; Alt+← / Alt+→ = previous / next page)")
SURVEY_HUD = ("Use ← / → to move selection. Press Enter to confirm. "
              "(F2 = skip the survey; Tab = move focus; Alt+← / Alt+→ = previous / next page)")


# =========================================================
# ========================= DPI / MONITOR =================
# =========================================================
def enable_per_monitor_dpi_awareness():
    if sys.platform != "win32":
        return
    user32 = ctypes.windll.user32
    try:
        DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = ctypes.c_void_p(-4)
        user32.SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)
        return
    except Exception:
        pass
    try:
        shcore = ctypes.windll.shcore
        PROCESS_PER_MONITOR_DPI_AWARE = 2
        shcore.SetProcessDpiAwareness(PROCESS_PER_MONITOR_DPI_AWARE)
        return
    except Exception:
        pass
    try:
        user32.SetProcessDPIAware()
    except Exception:
        pass


def _windows_dpi_for_window(hwnd):
    if sys.platform != "win32":
        return 96.0
    user32 = ctypes.windll.user32
    try:
        user32.GetDpiForWindow.restype = ctypes.c_uint
        dpi = user32.GetDpiForWindow(hwnd)
        return float(dpi) if dpi else 96.0
    except Exception:
        return 96.0


def get_dpi_for_window(hwnd):
    """Pixels per inch for the calibrated STIMULI (Landolt-C, contrast letters).

    The TV's true physical PPI differs from what Windows reports, so override it
    to make real-world sizes correct.
    """
    if USE_PHYSICAL_PPI:
        return float(MONITOR_PHYSICAL_PPI)
    return _windows_dpi_for_window(hwnd)


def get_layout_dpi(hwnd):
    """Pixels per inch used to lay out popups and the survey (not the stimuli)."""
    if USE_PHYSICAL_PPI:
        return float(LAYOUT_DPI)
    return _windows_dpi_for_window(hwnd)


class RECT(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG), ("right", wintypes.LONG), ("bottom", wintypes.LONG)]


class MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", RECT), ("rcWork", RECT), ("dwFlags", wintypes.DWORD)]


def get_monitor_rect_for_window(hwnd):
    user32 = ctypes.windll.user32
    MONITOR_DEFAULTTONEAREST = 2
    hmon = user32.MonitorFromWindow(wintypes.HWND(hwnd), MONITOR_DEFAULTTONEAREST)

    mi = MONITORINFO()
    mi.cbSize = ctypes.sizeof(MONITORINFO)

    user32.GetMonitorInfoW.argtypes = [wintypes.HMONITOR, ctypes.POINTER(MONITORINFO)]
    ok = user32.GetMonitorInfoW(hmon, ctypes.byref(mi))
    if not ok:
        return RECT(0, 0, user32.GetSystemMetrics(0), user32.GetSystemMetrics(1))

    return mi.rcWork


def set_dialog_half_screen(dialog, root_, frac=0.5):
    """Set a popup dialog to `frac` of the current monitor and center it.

    The marker is also used by the monitor hot-plug handler so the same dialog
    stays centered if Windows re-detects/switches the duplicated display.
    """
    # This is a dialog/page, not a deliberately offset companion window.
    # Always recenter it after a monitor change.
    try:
        dialog._keep_centered_on_monitor = True
    except Exception:
        pass

    try:
        mon = get_monitor_rect_for_window(root_.winfo_id())

        mon_w = mon.right - mon.left
        mon_h = mon.bottom - mon.top

        win_w = int(mon_w * frac)
        win_h = int(mon_h * frac)

        x = mon.left + (mon_w - win_w) // 2
        y = mon.top + (mon_h - win_h) // 2

        dialog.geometry(f"{win_w}x{win_h}+{x}+{y}")
        dialog.minsize(win_w, win_h)
        dialog.maxsize(win_w, win_h)

        return win_w, win_h

    except Exception:
        win_w, win_h = 900, 650
        dialog.geometry(f"{win_w}x{win_h}")
        dialog.minsize(win_w, win_h)
        dialog.maxsize(win_w, win_h)
        return win_w, win_h


# =========================================================
# ============= RESPONSIVE TOPLEVEL / MONITOR LAYOUT =====
# =========================================================
def _iter_toplevels(master):
    """Return every existing Tk Toplevel below `master`."""
    found = []

    def walk(widget):
        try:
            children = widget.winfo_children()
        except Exception:
            return
        for child in children:
            try:
                # The color-matching cursor ring is not a page: never move/focus it.
                if (isinstance(child, tk.Toplevel) and child.winfo_exists()
                        and not getattr(child, "_is_cursor_overlay", False)):
                    found.append(child)
            except Exception:
                pass
            walk(child)

    walk(master)
    return found


def _scale_widget_fonts(widget, scale):
    """Scale fonts in an already-created popup after a monitor-resolution change."""
    if scale <= 0:
        return

    try:
        # Avoid repeated scaling if the effective change is negligible.
        if abs(scale - 1.0) > 0.03 and "font" in widget.keys():
            current = tkfont.Font(font=widget.cget("font")).actual()
            old_size = int(current.get("size", 10))
            # Tk sometimes reports negative pixel sizes. Preserve the sign.
            sign = -1 if old_size < 0 else 1
            new_size = sign * max(8, int(round(abs(old_size) * scale)))
            styles = []
            if current.get("weight") == "bold":
                styles.append("bold")
            if current.get("slant") == "italic":
                styles.append("italic")
            if current.get("underline"):
                styles.append("underline")
            if current.get("overstrike"):
                styles.append("overstrike")
            widget.configure(font=(current.get("family", "Segoe UI"), new_size, *styles))
    except Exception:
        pass

    try:
        for child in widget.winfo_children():
            _scale_widget_fonts(child, scale)
    except Exception:
        pass


def _fit_popup_to_content(win, mon, preferred_w, preferred_h, centered=False):
    """Make a popup large enough for its actual rendered content.

    This is used after a monitor/DPI change.  Tk can report a new window size
    before labels/buttons have recomputed their requested size, which used to
    leave large text clipped inside a too-small fixed popup.
    """
    mon_w = max(1, mon.right - mon.left)
    mon_h = max(1, mon.bottom - mon.top)

    # Leave a small visible border so even a very large dialog remains findable.
    max_w = max(360, int(mon_w * 0.96))
    max_h = max(280, int(mon_h * 0.94))

    try:
        win.update_idletasks()
        req_w = max(1, int(win.winfo_reqwidth()))
        req_h = max(1, int(win.winfo_reqheight()))
    except Exception:
        req_w, req_h = preferred_w, preferred_h

    # Extra breathing room prevents wrapped labels/button borders from touching
    # the fixed window edge after a DPI/monitor change.
    pad_w = 36
    pad_h = 36
    target_w = max(int(preferred_w), req_w + pad_w)
    target_h = max(int(preferred_h), req_h + pad_h)

    # If content still cannot fit on the new monitor, progressively reduce only
    # the popup's fonts until its requested size fits.  This is a last resort;
    # normally the window simply grows instead.
    if target_w > max_w or target_h > max_h:
        for _ in range(8):
            try:
                win.update_idletasks()
                req_w = max(1, int(win.winfo_reqwidth()))
                req_h = max(1, int(win.winfo_reqheight()))
            except Exception:
                break
            if req_w + pad_w <= max_w and req_h + pad_h <= max_h:
                break
            _scale_widget_fonts(win, 0.92)

        try:
            win.update_idletasks()
            req_w = max(1, int(win.winfo_reqwidth()))
            req_h = max(1, int(win.winfo_reqheight()))
        except Exception:
            pass
        target_w = max(int(preferred_w), req_w + pad_w)
        target_h = max(int(preferred_h), req_h + pad_h)

    target_w = min(max_w, max(320, int(target_w)))
    target_h = min(max_h, max(240, int(target_h)))

    if centered:
        x = mon.left + (mon_w - target_w) // 2
        y = mon.top + (mon_h - target_h) // 2
    else:
        # Preserve the window's current relative location as much as possible,
        # but clamp it so the complete popup remains on the active monitor.
        try:
            x = int(win.winfo_x())
            y = int(win.winfo_y())
        except Exception:
            x, y = mon.left, mon.top
        x = min(max(x, mon.left), mon.right - target_w)
        y = min(max(y, mon.top), mon.bottom - target_h)

    try:
        win.geometry(f"{target_w}x{target_h}+{x}+{y}")
        win.update_idletasks()
    except Exception:
        pass

    return target_w, target_h


def _reflow_open_toplevels(old_mon, new_mon):
    """Rebuild open-window geometry safely after the duplicated monitor changes.

    Important order:
      1) release stale fixed-size constraints,
      2) map the window to the new monitor,
      3) rescale fonts,
      4) ask Tk how much space the rendered content actually needs,
      5) expand/recenter the popup to that required size,
      6) only then restore fixed-size constraints.

    This prevents labels and buttons from being cut off after HDMI hot-plug or
    Windows re-detects the cloned display at a different effective resolution.
    """
    if old_mon is None or new_mon is None:
        return

    old_w = max(1, old_mon.right - old_mon.left)
    old_h = max(1, old_mon.bottom - old_mon.top)
    new_w = max(1, new_mon.right - new_mon.left)
    new_h = max(1, new_mon.bottom - new_mon.top)

    scale_x = new_w / old_w
    scale_y = new_h / old_h

    # Do not aggressively enlarge fonts merely because the newly detected monitor
    # reports a larger pixel work area.  Shrinking when necessary is useful, but
    # uncontrolled enlargement is what commonly caused content clipping.
    font_scale = min(1.0, scale_x, scale_y)
    font_scale = max(0.75, font_scale)

    for win in _iter_toplevels(root):
        try:
            if not win.winfo_exists() or not win.winfo_viewable():
                continue

            win.update_idletasks()

            # Fullscreen pages simply remain fullscreen.
            try:
                if bool(win.attributes("-fullscreen")):
                    win.attributes("-fullscreen", True)
                    win.update_idletasks()
                    continue
            except Exception:
                pass

            # Remember whether this was designed as a non-resizable popup BEFORE
            # releasing old min/max constraints.
            try:
                was_fixed = (not win.resizable()[0] and not win.resizable()[1])
            except Exception:
                was_fixed = False

            x = win.winfo_x()
            y = win.winfo_y()
            w = max(1, win.winfo_width())
            h = max(1, win.winfo_height())

            rel_x = (x - old_mon.left) / old_w
            rel_y = (y - old_mon.top) / old_h
            rel_w = w / old_w
            rel_h = h / old_h

            preferred_w = max(320, int(round(rel_w * new_w)))
            preferred_h = max(240, int(round(rel_h * new_h)))

            centered = bool(getattr(win, "_keep_centered_on_monitor", False))
            if centered:
                new_x = new_mon.left + (new_w - preferred_w) // 2
                new_y = new_mon.top + (new_h - preferred_h) // 2
            else:
                new_x = int(round(new_mon.left + rel_x * new_w))
                new_y = int(round(new_mon.top + rel_y * new_h))

            # Release constraints inherited from the previous display.
            try:
                win.minsize(1, 1)
                win.maxsize(100000, 100000)
            except Exception:
                pass

            # First place it approximately where it belongs on the new monitor.
            win.geometry(f"{preferred_w}x{preferred_h}+{new_x}+{new_y}")
            win.update_idletasks()

            if abs(font_scale - 1.0) > 0.03:
                _scale_widget_fonts(win, font_scale)
                win.update_idletasks()

            # Crucial: measure AFTER font/layout changes, then make the actual
            # window large enough for its content.
            final_w, final_h = _fit_popup_to_content(
                win,
                new_mon,
                preferred_w,
                preferred_h,
                centered=centered,
            )

            # Re-lock fixed dialogs only after their final content-safe dimensions
            # are known.
            if was_fixed:
                try:
                    win.minsize(final_w, final_h)
                    win.maxsize(final_w, final_h)
                except Exception:
                    pass

        except Exception:
            pass


# =========================================================
# ========================= SMALL HELPERS =================
# =========================================================
def clamp01(x):
    return max(0.0, min(1.0, float(x)))


def rgb_to_hex(rgb):
    r, g, b = rgb
    return "#{:02X}{:02X}{:02X}".format(int(r), int(g), int(b))


def hex_to_rgb(hx):
    hx = hx.strip().lstrip("#")
    if len(hx) != 6:
        return (0, 0, 0)
    return (int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))


def hsv_to_rgb(h, s, v):
    h = (h % 1.0)
    s = clamp01(s)
    v = clamp01(v)
    i = int(h * 6.0)
    f = (h * 6.0) - i
    p = v * (1.0 - s)
    q = v * (1.0 - f * s)
    t = v * (1.0 - (1.0 - f) * s)
    i = i % 6
    if i == 0:
        r, g, b = v, t, p
    elif i == 1:
        r, g, b = q, v, p
    elif i == 2:
        r, g, b = p, v, t
    elif i == 3:
        r, g, b = p, q, v
    elif i == 4:
        r, g, b = t, p, v
    else:
        r, g, b = v, p, q
    return (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))


def rgb_dist(a, b):
    return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2)


# =========================================================
# ========================= LANDOLT-C SIZING ==============
# =========================================================
def landolt_base_diameter_mm_2020(viewing_distance_m):
    arcmin_rad = math.radians(1.0 / 60.0)  # 1 arcmin
    stroke_mm = viewing_distance_m * math.tan(arcmin_rad) * 1000.0
    return 5.0 * stroke_mm


def diameter_mm_for_20_over_den(viewing_distance_m, den):
    return landolt_base_diameter_mm_2020(viewing_distance_m) * (den / 20.0)


def draw_landolt_c(canvas_, dpi, viewing_distance_m, acuity_denominator, orientation):
    canvas_.update_idletasks()
    w = canvas_.winfo_width()
    h = canvas_.winfo_height()
    if w <= 1 or h <= 1:
        # Window not shown yet: use the canvas's configured size instead of 1 px.
        w, h = canvas_.winfo_reqwidth(), canvas_.winfo_reqheight()
    cx, cy = w / 2.0, h / 2.0

    diam_mm = diameter_mm_for_20_over_den(viewing_distance_m, acuity_denominator)
    stroke_mm = diam_mm / 5.0

    diam_px = (diam_mm / 25.4) * dpi
    stroke_px = (stroke_mm / 25.4) * dpi

    r_outer = diam_px / 2.0
    r_inner = max(0.1, r_outer - stroke_px)

    canvas_.delete("all")
    canvas_.configure(bg=BG_SOFT)
    canvas_.create_oval(cx - r_outer, cy - r_outer, cx + r_outer, cy + r_outer, fill="black", outline="")
    canvas_.create_oval(cx - r_inner, cy - r_inner, cx + r_inner, cy + r_inner, fill=BG_SOFT, outline="")

    half_gap = stroke_px / 2.0
    if orientation == "right":
        x1, y1, x2, y2 = cx, cy - half_gap, cx + r_outer + 2, cy + half_gap
    elif orientation == "left":
        x1, y1, x2, y2 = cx - r_outer - 2, cy - half_gap, cx, cy + half_gap
    elif orientation == "up":
        x1, y1, x2, y2 = cx - half_gap, cy - r_outer - 2, cx + half_gap, cy
    elif orientation == "down":
        x1, y1, x2, y2 = cx - half_gap, cy, cx + half_gap, cy + r_outer + 2
    else:
        x1, y1, x2, y2 = cx, cy - half_gap, cx + r_outer + 2, cy + half_gap

    canvas_.create_rectangle(x1, y1, x2, y2, fill=BG_SOFT, outline="")


# =========================================================
# ========================= CONTRAST HELPERS ==============
# =========================================================
def triplet_index_from_letter_index(letter_index):
    return int(letter_index) // 3


def hex_for_triplet(triplet_idx):
    triplet_idx = max(0, min(triplet_idx, len(CONTRAST_HEX_SEQUENCE) - 1))
    return CONTRAST_HEX_SEQUENCE[triplet_idx]


def logcs_for_triplet(triplet_idx):
    triplet_idx = max(0, min(triplet_idx, len(CONTRAST_LOGCS_BY_TRIPLET) - 1))
    return CONTRAST_LOGCS_BY_TRIPLET[triplet_idx]


# --- PIL font path ---
def _pick_windows_font_path():
    if sys.platform != "win32":
        return None
    win_fonts = Path(r"C:\Windows\Fonts")
    candidates = [
        win_fonts / "seguisb.ttf",
        win_fonts / "segoeuib.ttf",
        win_fonts / "arialbd.ttf",
        win_fonts / "calibrib.ttf",
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    return None


_PIL_FONT_PATH = _pick_windows_font_path()


def _make_letter_photo_exact_box(letter, box_px, color_hex, font_path=None, pad_ratio=0.06,
                                 x_scale=1.0, y_scale=1.0, weight_px=0):
    if not PIL_OK:
        return None

    box_px = int(max(8, round(box_px)))
    oversample = 4
    big = box_px * oversample

    img = Image.new("RGBA", (big, big), (255, 255, 255, 0))
    draw = ImageDraw.Draw(img)

    fs = int(big * 0.85)
    try:
        if font_path:
            font = ImageFont.truetype(font_path, fs)
        else:
            font = ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()

    draw.text((big // 2, big // 2), letter, font=font, fill=color_hex, anchor="mm")

    bbox = img.getbbox()
    if not bbox:
        return None

    glyph = img.crop(bbox)

    if weight_px and weight_px > 0 and ImageFilter is not None:
        a = glyph.split()[-1]
        a = a.filter(ImageFilter.MaxFilter(size=weight_px * 2 + 1))
        r, g, b = hex_to_rgb(color_hex)
        colored_pixels = Image.new("RGBA", glyph.size, (r, g, b, 255))
        blank = Image.new("RGBA", glyph.size, (0, 0, 0, 0))
        glyph = Image.composite(colored_pixels, blank, a)

    x_scale = float(x_scale)
    y_scale = float(y_scale)
    if abs(x_scale - 1.0) > 1e-6 or abs(y_scale - 1.0) > 1e-6:
        w, h = glyph.size
        new_w = max(1, int(round(w * x_scale)))
        new_h = max(1, int(round(h * y_scale)))
        try:
            resample = Image.Resampling.LANCZOS
        except Exception:
            resample = Image.LANCZOS
        glyph = glyph.resize((new_w, new_h), resample=resample)

    pad = max(1, int(round(box_px * oversample * pad_ratio)))
    glyph = ImageOps.expand(glyph, border=pad, fill=(255, 255, 255, 0))

    try:
        resample = Image.Resampling.LANCZOS
    except Exception:
        resample = Image.LANCZOS
    glyph = glyph.resize((box_px, box_px), resample=resample)

    return ImageTk.PhotoImage(glyph)


def draw_contrast_triplet(canvas_, dpi, box_mm, letters3, start_letter_index):
    assert len(letters3) == 3

    canvas_.update_idletasks()
    W = canvas_.winfo_width()
    H = canvas_.winfo_height()
    if W <= 1 or H <= 1:
        # Window not shown yet: use the canvas's configured size instead of 1 px.
        W, H = canvas_.winfo_reqwidth(), canvas_.winfo_reqheight()
    cy = H / 2.0

    box_px = max(10.0, (box_mm / 25.4) * dpi)
    gap_px = max(6.0, box_px * 0.12)

    total_w = 3 * box_px + 2 * gap_px
    left = (W - total_w) / 2.0

    cxs = [
        left + box_px / 2.0,
        left + box_px + gap_px + box_px / 2.0,
        left + 2 * (box_px + gap_px) + box_px / 2.0
    ]

    canvas_.delete("all")
    canvas_.configure(bg=BG_SOFT)
    # Remember where the letters sit so the typed answer can be drawn below them.
    canvas_._contrast_layout = (cy, box_px, W, H)

    if SHOW_CONTRAST_BOX:
        for cx in cxs:
            half = box_px / 2.0
            canvas_.create_rectangle(cx - half, cy - half, cx + half, cy + half, outline="gray60", width=2)

    t_idx = triplet_index_from_letter_index(start_letter_index)
    hex_color = hex_for_triplet(t_idx)

    if PIL_OK:
        photos = []
        for j, cx in enumerate(cxs):
            ph = _make_letter_photo_exact_box(
                letter=letters3[j],
                box_px=box_px,
                color_hex=hex_color,
                font_path=_PIL_FONT_PATH,
                x_scale=LETTER_WIDTH_SCALE,
                y_scale=LETTER_HEIGHT_SCALE,
                weight_px=LETTER_WEIGHT_PX
            )
            photos.append(ph)
            if ph is not None:
                canvas_.create_image(cx, cy, image=ph)
            else:
                canvas_.create_text(cx, cy, text=letters3[j], fill=hex_color, font=("Segoe UI", 120, "bold"))
        canvas_._contrast_letter_photos = photos  # keep refs
        canvas_.update_idletasks()
        return {"box_px": round(box_px, 2), "triplet_idx": t_idx, "hex": hex_color}

    # fallback
    for j, cx in enumerate(cxs):
        canvas_.create_text(cx, cy, text=letters3[j], fill=hex_color, font=("Segoe UI", 120, "bold"))

    canvas_.update_idletasks()
    return {"box_px": round(box_px, 2), "triplet_idx": t_idx, "hex": hex_color}


def draw_contrast_input(canvas_, typed):
    """Show the participant's typed answer directly below the target letters.

    Three slots, e.g. "V   N   _", so the researcher and participant can see
    exactly which letters have been registered for the current triplet.
    """
    canvas_.delete("contrast_input")
    layout = getattr(canvas_, "_contrast_layout", None)
    if not layout:
        return
    cy, box_px, W, H = layout
    letters_bottom = cy + box_px / 2.0
    space = H - letters_bottom
    if space < 30:
        return
    # Size the text from the free space under the letters so it always fits on screen.
    font_px = int(max(24, min(110, space * 0.42)))
    y = letters_bottom + space * 0.45
    slots = [typed[i] if i < len(typed) else "_" for i in range(3)]
    canvas_.create_text(
        W / 2.0, y,
        text="   ".join(slots),
        font=("Segoe UI", -font_px, "bold"),
        fill="#333333",
        anchor="center",
        tags=("contrast_input",),
    )


# =========================================================
# ========================= EXCEL (ONLY YOUR COLUMNS) =====
# =========================================================
def build_headers():
    return [
        "participant_id",
        "sequence_step_order",
        "raw_full_sequence_position",
        "condition",
        "full_sequence",
        "participant_email",
        "H_illuminance_lux",
        "V_illuminance_lux",
        "Sun_position_inch",
        "screening_visual_acuity_20_20_correct_of_2",
        "screening_visual_acuity_20_20_pass",
        "screening_ishihara_1_response",
        "screening_ishihara_1_pass",
        "screening_ishihara_2_response",
        "screening_ishihara_2_pass",
        "start_time",
        "end_time",
        "Visual acuity achieved",
        "Visual acuity failed at",
        "Contrast sensitivity achieved_log(CS)",
        "Contrast sensitivity failed at_log(CS)",
        "Contrast sensitivity achieved_hexcode",
        "Contrast sensitivity failed at_hexcode",
        "color_target_hex",
        "color_target_rgb",
        "color_selected_hex",
        "color_selected_rgb",
        "color_error_rgb_dist",
        "survey_clarity_of_view_7pt",
        "survey_visual_privacy_7pt",
        "survey_reflections_mirror_effect_7pt",
        "survey_visual_comfort_glare_7pt",
    ]


HEADERS = build_headers()


def ensure_workbook(path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = load_workbook(path) if path.exists() else Workbook()
    if SHEET in wb.sheetnames:
        ws = wb[SHEET]
    else:
        ws = wb.active
        ws.title = SHEET

    existing = [ws.cell(1, c).value for c in range(1, len(HEADERS) + 1)]
    if existing != HEADERS:
        ws.delete_rows(1, ws.max_row)
        ws.append(HEADERS)
    wb.save(path)


def write_participant_row(path, row_dict, row_number=None):
    """Write one result row: append a new row, or overwrite `row_number`.

    Returns the row number written (a condition reopened with Alt+Left keeps its row).
    """
    wb = load_workbook(path)
    ws = wb[SHEET]
    row = [row_dict.get(h, "") for h in HEADERS]
    if row_number is None:
        ws.append(row)
        row_number = ws.max_row
    else:
        for col, value in enumerate(row, start=1):
            ws.cell(row=row_number, column=col, value=value)
    wb.save(path)
    return row_number


# =========================================================
# ================= WINDOWS DISPLAY MODE ==================
# =========================================================
# Windows DISPLAY_DEVICE structure/flags used to notice HDMI/monitor hot-plug events.
# We watch the set of active desktop-attached display devices.  If that set changes,
# the experiment automatically asks Windows for Duplicate/Clone mode again.
_DISPLAY_DEVICE_ATTACHED_TO_DESKTOP = 0x00000001


class DISPLAY_DEVICEW(ctypes.Structure):
    _fields_ = [
        ("cb", wintypes.DWORD),
        ("DeviceName", wintypes.WCHAR * 32),
        ("DeviceString", wintypes.WCHAR * 128),
        ("StateFlags", wintypes.DWORD),
        ("DeviceID", wintypes.WCHAR * 128),
        ("DeviceKey", wintypes.WCHAR * 128),
    ]


def get_active_display_signature():
    """Return a stable tuple describing active Windows display devices.

    This is used only to detect that a monitor/HDMI configuration changed.
    It does not alter the display settings.
    """
    if sys.platform != "win32":
        return ()

    user32 = ctypes.windll.user32
    devices = []
    index = 0

    while True:
        dd = DISPLAY_DEVICEW()
        dd.cb = ctypes.sizeof(DISPLAY_DEVICEW)
        ok = user32.EnumDisplayDevicesW(None, index, ctypes.byref(dd), 0)
        if not ok:
            break

        if dd.StateFlags & _DISPLAY_DEVICE_ATTACHED_TO_DESKTOP:
            devices.append((str(dd.DeviceName), str(dd.DeviceString)))

        index += 1

    return tuple(sorted(devices))


def force_duplicate_display():
    """Force Windows to use Duplicate/Clone mode for connected displays.

    This is equivalent to pressing Win + P and choosing "Duplicate".
    The function is called once when the experiment program starts.
    """
    if sys.platform != "win32":
        return True, "Duplicate display is only applicable on Windows."

    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(
            ["DisplaySwitch.exe", "/clone"],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )

        # Give Windows a moment to apply the display topology before Tkinter
        # measures the monitor and creates the experiment windows.
        time.sleep(2.0)

        if result.returncode == 0:
            return True, "Windows was switched to Duplicate display mode."
        return False, f"DisplaySwitch.exe returned code {result.returncode}."

    except Exception as e:
        return False, str(e)


# ----- Screens for two-screen color matching -----
class MONITORINFOEXW(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", RECT), ("rcWork", RECT),
                ("dwFlags", wintypes.DWORD), ("szDevice", wintypes.WCHAR * 32)]


class _DC_LUID(ctypes.Structure):
    _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]


class _DC_PATH_SOURCE_INFO(ctypes.Structure):
    _fields_ = [("adapterId", _DC_LUID), ("id", ctypes.c_uint32),
                ("modeInfoIdx", ctypes.c_uint32), ("statusFlags", ctypes.c_uint32)]


class _DC_RATIONAL(ctypes.Structure):
    _fields_ = [("Numerator", ctypes.c_uint32), ("Denominator", ctypes.c_uint32)]


class _DC_PATH_TARGET_INFO(ctypes.Structure):
    _fields_ = [("adapterId", _DC_LUID), ("id", ctypes.c_uint32), ("modeInfoIdx", ctypes.c_uint32),
                ("outputTechnology", ctypes.c_uint32), ("rotation", ctypes.c_uint32),
                ("scaling", ctypes.c_uint32), ("refreshRate", _DC_RATIONAL),
                ("scanLineOrdering", ctypes.c_uint32), ("targetAvailable", wintypes.BOOL),
                ("statusFlags", ctypes.c_uint32)]


class _DC_PATH_INFO(ctypes.Structure):
    _fields_ = [("sourceInfo", _DC_PATH_SOURCE_INFO), ("targetInfo", _DC_PATH_TARGET_INFO),
                ("flags", ctypes.c_uint32)]


class _DC_MODE_INFO(ctypes.Structure):
    _fields_ = [("infoType", ctypes.c_uint32), ("id", ctypes.c_uint32),
                ("adapterId", _DC_LUID), ("data", ctypes.c_byte * 48)]


class _DC_DEVICE_INFO_HEADER(ctypes.Structure):
    _fields_ = [("type", ctypes.c_uint32), ("size", ctypes.c_uint32),
                ("adapterId", _DC_LUID), ("id", ctypes.c_uint32)]


class _DC_SOURCE_DEVICE_NAME(ctypes.Structure):
    _fields_ = [("header", _DC_DEVICE_INFO_HEADER), ("viewGdiDeviceName", wintypes.WCHAR * 32)]


def list_monitors():
    """Every active screen: dicts with 'rect' and 'work' (RECT), 'primary' and GDI 'device'."""
    if sys.platform != "win32":
        return []
    user32 = ctypes.windll.user32
    found = []
    proc_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMONITOR, wintypes.HDC,
                                   ctypes.POINTER(RECT), wintypes.LPARAM)

    def callback(hmon, _hdc, _rect, _lparam):
        mi = MONITORINFOEXW()
        mi.cbSize = ctypes.sizeof(MONITORINFOEXW)
        # GetMonitorInfoW is declared elsewhere to take a MONITORINFO pointer; the
        # extended struct starts with the same fields, and cbSize tells Windows to
        # also fill in the device name.
        if user32.GetMonitorInfoW(hmon, ctypes.cast(ctypes.pointer(mi), ctypes.POINTER(MONITORINFO))):
            found.append({
                "rect": RECT(mi.rcMonitor.left, mi.rcMonitor.top, mi.rcMonitor.right, mi.rcMonitor.bottom),
                "work": RECT(mi.rcWork.left, mi.rcWork.top, mi.rcWork.right, mi.rcWork.bottom),
                "primary": bool(mi.dwFlags & 1),
                "device": str(mi.szDevice),
            })
        return True

    try:
        user32.EnumDisplayMonitors(None, None, proc_type(callback), 0)
    except Exception:
        return []
    return found


def _active_display_paths():
    """Windows' active display paths: one per connected screen in use (any mode)."""
    if sys.platform != "win32":
        return []
    user32 = ctypes.windll.user32
    try:
        n_paths, n_modes = ctypes.c_uint32(), ctypes.c_uint32()
        QDC_ONLY_ACTIVE_PATHS = 2
        if user32.GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, ctypes.byref(n_paths), ctypes.byref(n_modes)):
            return []
        paths = (_DC_PATH_INFO * n_paths.value)()
        modes = (_DC_MODE_INFO * n_modes.value)()
        if user32.QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ctypes.byref(n_paths), paths,
                                     ctypes.byref(n_modes), modes, None):
            return []
        return list(paths[:n_paths.value])
    except Exception:
        return []


def active_display_count():
    """How many screens are connected and in use (2 with the TV plugged in, even in Duplicate)."""
    return len(_active_display_paths())


def internal_display_devices():
    """GDI device names (e.g. \\\\.\\DISPLAY1) of built-in laptop panels."""
    if sys.platform != "win32":
        return set()
    INTERNAL_TECHNOLOGIES = (0x80000000, 11, 13)  # internal, embedded DisplayPort, embedded UDI
    user32 = ctypes.windll.user32
    names = set()
    try:
        for path in _active_display_paths():
            if path.targetInfo.outputTechnology not in INTERNAL_TECHNOLOGIES:
                continue
            req = _DC_SOURCE_DEVICE_NAME()
            req.header.type = 1  # DISPLAYCONFIG_DEVICE_INFO_GET_SOURCE_NAME
            req.header.size = ctypes.sizeof(_DC_SOURCE_DEVICE_NAME)
            req.header.adapterId = path.sourceInfo.adapterId
            req.header.id = path.sourceInfo.id
            if user32.DisplayConfigGetDeviceInfo(ctypes.byref(req)) == 0:
                names.add(str(req.viewGdiDeviceName))
    except Exception:
        pass
    return names


def find_color_screens():
    """Return (tv_monitor, laptop_monitor) when two screens are active, else None.

    The laptop is the built-in panel; if Windows cannot tell, the main (primary)
    display is taken as the laptop and the other one as the TV.
    """
    monitors = list_monitors()
    if len(monitors) < 2:
        return None
    internal = internal_display_devices()
    laptops = [m for m in monitors if m["device"] in internal]
    others = [m for m in monitors if m["device"] not in internal]
    if len(laptops) == 1 and others:
        return others[0], laptops[0]
    primary = [m for m in monitors if m["primary"]]
    rest = [m for m in monitors if not m["primary"]]
    if primary and rest:
        return rest[0], primary[0]
    return None


def request_display_mode(mode):
    """Ask Windows for 'extend' or 'clone' (= Duplicate) without blocking the UI."""
    if sys.platform != "win32":
        return False
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["DisplaySwitch.exe", "/extend" if mode == "extend" else "/clone"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )
        return True
    except Exception:
        return False


def request_duplicate_display_nonblocking():
    """Ask Windows for Duplicate mode without blocking the Tkinter event loop."""
    if sys.platform != "win32":
        return True
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.Popen(
            ["DisplaySwitch.exe", "/clone"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=creationflags,
        )
        return True
    except Exception:
        return False



# =========================================================
# ============ ACCESSIBLE FOCUS / LARGE CONFIRM ===========
# =========================================================
FOCUS_RING_COLOR = "#0067C0"      # strong Windows-style blue focus ring
FOCUS_RING_WIDTH = 7
FOCUS_BG = "#DCEEFF"              # pale blue fill for the currently focused button


def _focus_in_button(event):
    """Make the currently keyboard-selected button unmistakably visible."""
    w = event.widget
    try:
        if not hasattr(w, "_normal_bg"):
            w._normal_bg = w.cget("bg")
        w.configure(
            highlightthickness=FOCUS_RING_WIDTH,
            highlightbackground=FOCUS_RING_COLOR,
            highlightcolor=FOCUS_RING_COLOR,
            relief="solid",
            bd=3,
            bg=FOCUS_BG,
        )
    except Exception:
        pass


def _focus_out_button(event):
    w = event.widget
    try:
        normal_bg = getattr(w, "_normal_bg", w.cget("bg"))
        w.configure(
            highlightthickness=2,
            highlightbackground="black",
            highlightcolor="black",
            bd=2,
            bg=normal_bg,
        )
    except Exception:
        pass


def _focus_in_entry(event):
    """Also make the active text-entry box obvious to the researcher."""
    try:
        event.widget.configure(
            highlightthickness=FOCUS_RING_WIDTH,
            highlightbackground=FOCUS_RING_COLOR,
            highlightcolor=FOCUS_RING_COLOR,
        )
    except Exception:
        pass


def _focus_out_entry(event):
    try:
        event.widget.configure(
            highlightthickness=1,
            highlightbackground="gray55",
            highlightcolor="gray55",
        )
    except Exception:
        pass


def _button_focus_previous(event):
    """Left/Up arrow: move keyboard focus to the previous selectable control."""
    try:
        prev_w = event.widget.tk_focusPrev()
        if prev_w is not None:
            prev_w.focus_set()
    except Exception:
        pass
    return "break"


def _button_focus_next(event):
    """Right/Down arrow: move keyboard focus to the next selectable control."""
    try:
        next_w = event.widget.tk_focusNext()
        if next_w is not None:
            next_w.focus_set()
    except Exception:
        pass
    return "break"


def ask_large_skip_confirmation(parent, title, message, default="no"):
    """
    Large ESC skip confirmation dialog.

    Returns:
        should_skip: bool

    Keyboard:
      - Tab / Shift+Tab: move through controls
      - Left / Right / Up / Down arrows: move between Yes/No buttons
      - Enter: activate the focused button
      - Esc: cancel the skip and continue the test
    """
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(parent)
    # Keep this confirmation page centered when a different monitor is detected.
    dialog._keep_centered_on_monitor = True

    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    try:
        mon = get_monitor_rect_for_window(parent.winfo_id())
        mon_w = mon.right - mon.left
        mon_h = mon.bottom - mon.top
        win_w = int(mon_w * 0.68)
        win_h = int(mon_h * 0.58)
        x = mon.left + (mon_w - win_w) // 2
        y = mon.top + (mon_h - win_h) // 2
        dialog.geometry(f"{win_w}x{win_h}+{x}+{y}")
    except Exception:
        win_w, win_h = 1100, 650
        dialog.geometry(f"{win_w}x{win_h}")

    dialog.minsize(win_w, win_h)
    dialog.maxsize(win_w, win_h)

    result = {"skip": False}

    wrap = tk.Frame(dialog, bg=BG_SOFT, padx=55, pady=45)
    wrap.pack(fill="both", expand=True)

    tk.Label(
        wrap, text=title, bg=BG_SOFT, fg="black",
        font=("Segoe UI", 42, "bold"), justify="center", anchor="center",
    ).pack(fill="x", pady=(20, 22))

    tk.Label(
        wrap, text=message, bg=BG_SOFT, fg="black",
        font=("Segoe UI", 34), wraplength=max(750, int(win_w * 0.80)),
        justify="center", anchor="center",
    ).pack(fill="x", pady=(0, 34))

    tk.Label(
        wrap,
        text="Tab / Shift+Tab or Arrow keys = move selection     Enter = confirm",
        bg=BG_SOFT, fg="gray25", font=("Segoe UI", 18),
    ).pack(pady=(0, 24))

    btns = tk.Frame(wrap, bg=BG_SOFT)
    btns.pack(anchor="center", pady=(0, 8))

    def choose(value):
        result["skip"] = bool(value)
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    yes_btn = tk.Button(
        btns, text="Yes — Skip", font=("Segoe UI", 28, "bold"),
        width=14, height=2, takefocus=True, command=lambda: choose(True),
    )
    yes_btn.pack(side="left", padx=22)

    no_btn = tk.Button(
        btns, text="No — Continue", font=("Segoe UI", 28, "bold"),
        width=14, height=2, takefocus=True, command=lambda: choose(False),
    )
    no_btn.pack(side="left", padx=22)

    # Plain informational text at the bottom of the Skip Test dialog.
    tk.Label(
        wrap,
        text="Participants cannot read",
        bg=BG_SOFT,
        fg="gray25",
        font=("Segoe UI", 20),
        justify="center",
        anchor="center",
    ).pack(fill="x", pady=(28, 0))

    def activate_focused(_event=None):
        w = dialog.focus_get()
        if isinstance(w, tk.Button):
            w.invoke()
            return "break"
        return None

    dialog.bind("<Return>", activate_focused)
    dialog.bind("<Escape>", lambda e: (choose(False), "break")[1])
    dialog.protocol("WM_DELETE_WINDOW", lambda: choose(False))

    # Arrow navigation between the two action buttons.
    for key in ("<Left>", "<Up>"):
        yes_btn.bind(key, lambda e: (no_btn.focus_set(), "break")[1])
        no_btn.bind(key, lambda e: (yes_btn.focus_set(), "break")[1])
    for key in ("<Right>", "<Down>"):
        yes_btn.bind(key, lambda e: (no_btn.focus_set(), "break")[1])
        no_btn.bind(key, lambda e: (yes_btn.focus_set(), "break")[1])

    default_btn = yes_btn if str(default).lower() == "yes" else no_btn
    dialog.after(100, default_btn.focus_force)

    dialog.lift()
    try:
        dialog.attributes("-topmost", True)
        dialog.after(180, lambda: dialog.attributes("-topmost", False))
    except Exception:
        pass

    parent.wait_window(dialog)
    return result["skip"]


# =========================================================
# ============== PAGE NAVIGATION (Alt+← / Alt+→) ==========
# =========================================================
# Alt+Left  = go back to the previous page and redo it.
# Alt+Right = go on to the next page.
# A "page" is every screen of the session: the screening pages, the participant
# setup, each instruction, each condition's illuminance screen, each test (acuity,
# contrast, color matching) and each survey question.
# Leaving a page without finishing it never changes its saved result, except that
# going forward past a test that was never finished records it as "SKIPPED".
# A redone page's new answer replaces the old one only once the redo is finished.
NAV_BACK = "back"
NAV_NEXT = "next"
NAV_REPEAT_GUARD_S = 0.30  # ignore key auto-repeat so one press moves one page

_nav_handler = None        # callable(direction) for the page on screen, or None
_nav_page_windows = set()  # windows that belong to the page on screen
_nav_last_time = 0.0


def set_nav_page(handler, *windows):
    """Register the page now on screen: `handler(direction)` and its window(s)."""
    global _nav_handler, _nav_page_windows
    _nav_handler = handler
    _nav_page_windows = set(w for w in windows if w is not None)


def clear_nav_page():
    set_nav_page(None)


def _on_nav_key(event, direction):
    global _nav_last_time
    handler = _nav_handler
    if handler is None:
        return "break"
    try:
        top = event.widget.winfo_toplevel()
    except Exception:
        return "break"
    # Only the page on screen reacts (not e.g. a Skip confirmation or the progress list).
    if top not in _nav_page_windows:
        return "break"
    now = time.monotonic()
    if now - _nav_last_time < NAV_REPEAT_GUARD_S:
        return "break"
    _nav_last_time = now
    handler(direction)
    return "break"


def bind_nav_keys(widget):
    """Bind Alt+Left / Alt+Right on `widget` so they win over its plain arrow-key bindings.

    Without this, a plain <Left>/<Right> binding also fires for Alt+Left/Alt+Right
    (e.g. it would count as a "left" answer in the acuity test).
    """
    widget.bind("<Alt-Left>", lambda e: _on_nav_key(e, NAV_BACK))
    widget.bind("<Alt-Right>", lambda e: _on_nav_key(e, NAV_NEXT))


# =========================================================
# ========================= MAIN APP BOOT =================
# =========================================================
# Force laptop + HDMI display to mirror each other before creating the UI.
_duplicate_ok, _duplicate_message = force_duplicate_display()

enable_per_monitor_dpi_awareness()

root = tk.Tk()

# Render text at a fixed 100% size (96 px per inch) instead of following Windows
# display scaling. Every screen was laid out for a 1920x1080 display at 100%; in
# Duplicate mode Windows can apply a laptop's 125-150% scaling to the TV, which made
# text and buttons spill off popups. Stimulus sizes are unaffected (they use PPI).
root.tk.call("tk", "scaling", 96.0 / 72.0)

# If Windows could not switch automatically, warn the researcher but allow
# the program to continue so the display can be set manually with Win + P.
if not _duplicate_ok and sys.platform == "win32":
    messagebox.showwarning(
        "Duplicate display",
        "The program could not automatically switch Windows to Duplicate display mode.\n\n"
        f"Details: {_duplicate_message}\n\n"
        "Please press Win + P and choose Duplicate before continuing.",
        parent=root,
    )
root.title('Landolt-C + Contrast Triplets + Color Match + Survey (3.048 m / 10 ft)')
root.resizable(False, False)
root.configure(bg=BG_SOFT)

# Make keyboard focus visible everywhere. Tab / Shift+Tab now navigate controls.
root.option_add("*Button.takeFocus", 1)
root.option_add("*Entry.takeFocus", 1)
root.bind_class("Button", "<FocusIn>", _focus_in_button, add="+")
root.bind_class("Button", "<FocusOut>", _focus_out_button, add="+")
root.bind_class("Button", "<Left>", _button_focus_previous, add="+")
root.bind_class("Button", "<Up>", _button_focus_previous, add="+")
root.bind_class("Button", "<Right>", _button_focus_next, add="+")
root.bind_class("Button", "<Down>", _button_focus_next, add="+")
root.bind_class("Entry", "<FocusIn>", _focus_in_entry, add="+")
root.bind_class("Entry", "<FocusOut>", _focus_out_entry, add="+")

# Page navigation (Alt+Left / Alt+Right) works in every window. The more specific
# bindings make sure Alt+arrows are never also treated as a plain arrow press
# (main window = acuity answers, buttons = focus moves, entries = cursor moves).
root.bind_all("<Alt-Left>", lambda e: _on_nav_key(e, NAV_BACK))
root.bind_all("<Alt-Right>", lambda e: _on_nav_key(e, NAV_NEXT))
bind_nav_keys(root)
for _cls in ("Button", "Entry"):
    root.bind_class(_cls, "<Alt-Left>", lambda e: _on_nav_key(e, NAV_BACK))
    root.bind_class(_cls, "<Alt-Right>", lambda e: _on_nav_key(e, NAV_NEXT))

# HUD
hud = tk.Frame(root, bg=BG_SOFT)
hud.pack(fill="x", padx=10, pady=(8, 0))

stage_var = tk.StringVar(value="")
progress_var = tk.StringVar(value="")
instr_var = tk.StringVar(value="")

row1 = tk.Frame(hud, bg=BG_SOFT)
row1.pack(fill="x")

tk.Label(row1, textvariable=stage_var, bg=BG_SOFT, font=HUD_TITLE_FONT, anchor="w").pack(side="left")
tk.Label(row1, textvariable=progress_var, bg=BG_SOFT, font=HUD_TEXT_FONT, fg="gray25", anchor="e").pack(side="right")
tk.Label(hud, textvariable=instr_var, bg=BG_SOFT, font=HUD_INSTR_FONT, fg="black", anchor="w").pack(fill="x", pady=(4, 6))

# Quit button (ESC)
def on_escape(event=None):
    global stage
    # Esc means "skip the current step". Only offer it on stages that actually
    # have a skippable step; anywhere else Esc does nothing.
    if stage not in (STAGE_ACUITY, STAGE_CONTRAST, STAGE_COLOR, STAGE_SURVEY):
        return

    # Always confirm before skipping.
    should_skip = ask_large_skip_confirmation(
        root,
        title="Skip test?",
        message="Are you sure you want to skip this test?",
        default="no",
    )
    if not should_skip:
        return  # user chose No — stay on the current test

    if stage == STAGE_ACUITY:
        # synchronized behavior: ESC uses the same finalize logic as "2 wrong"
        handle_acuity_failure(esc=True)
        return
    if stage == STAGE_CONTRAST:
        handle_contrast_failure(esc=True)
        return
    if stage == STAGE_COLOR:
        _finish_color_and_move_to_survey()
        return
    if stage == STAGE_SURVEY:
        # Esc no longer quits the session: skip the whole survey and continue
        # to the next condition. Exit is only via the window's X button.
        if _survey_skip_all is not None:
            _survey_skip_all()
        return


quit_btn = tk.Button(
    hud,
    text="Skip (ESC)",
    command=on_escape,
    bg=QUIT_BG,
    fg="black",
    activebackground=QUIT_BG,
    relief="solid",
    bd=2,
    highlightthickness=2,
    highlightbackground="black",
    font=QUIT_FONT,
    padx=10,
    pady=4
)
quit_btn.pack(anchor="w", pady=(6, 2))

# Exit button: the on-screen way to end the whole run during the fullscreen
# stages (acuity / contrast), where the window has no titlebar X. Same action
# as clicking a window's close (X) button.
exit_btn = tk.Button(
    hud,
    text="Exit ✕",
    command=lambda: save_and_exit(),
    bg="#C00000",
    fg="white",
    activebackground="#C00000",
    activeforeground="white",
    relief="solid",
    bd=2,
    highlightthickness=2,
    highlightbackground="black",
    font=QUIT_FONT,
    padx=10,
    pady=4
)
exit_btn.pack(anchor="w", pady=(0, 6))

# main canvas
canvas = tk.Canvas(root, bg=BG_SOFT, highlightthickness=0)
canvas.pack(fill="both", expand=True)


def force_initial_geometry():
    # Fullscreen grey field. The stimulus is still drawn at true physical size,
    # centered in the canvas; fullscreen only hides the desktop/taskbar/borders.
    root.update_idletasks()
    try:
        root.minsize(1, 1)
        root.maxsize(100000, 100000)
        root.attributes("-fullscreen", True)
    except Exception:
        pass
    root.configure(bg=BG_SOFT)
    root.update_idletasks()


force_initial_geometry()


def activate_stimulus_window():
    try:
        root.deiconify()
    except Exception:
        pass
    try:
        root.lift()
        root.attributes("-topmost", True)
        root.after(40, lambda: root.attributes("-topmost", False))
    except Exception:
        pass
    try:
        root.focus_force()
        canvas.focus_set()
    except Exception:
        pass


# Transition overlay
_transition_id = None
def show_transition(title, subtitle="", ms=650, after_fn=None):
    global _transition_id
    canvas.delete("transition")

    canvas.update_idletasks()
    W = max(10, canvas.winfo_width())
    H = max(10, canvas.winfo_height())

    canvas.create_rectangle(0, 0, W, H, fill=BG_SOFT, outline="", tags=("transition",))
    canvas.create_text(W/2, H*0.42, text=title, font=("Segoe UI", 26, "bold"),
                       fill="black", tags=("transition",))
    if subtitle:
        canvas.create_text(W/2, H*0.55, text=subtitle, font=("Segoe UI", 14),
                           fill="gray25", tags=("transition",))
    canvas.create_text(W/2, H*0.70, text="(Please wait…)", font=("Segoe UI", 12, "italic"),
                       fill="gray35", tags=("transition",))

    if _transition_id is not None:
        try:
            root.after_cancel(_transition_id)
        except Exception:
            pass
        _transition_id = None

    def done():
        global _transition_id
        _transition_id = None
        canvas.delete("transition")
        if callable(after_fn):
            after_fn()

    _transition_id = root.after(ms, done)


def cancel_transition():
    """Drop a pending transition screen without starting the stage it leads to."""
    global _transition_id
    if _transition_id is not None:
        try:
            root.after_cancel(_transition_id)
        except Exception:
            pass
        _transition_id = None
    canvas.delete("transition")


# =========================================================
# ============ PARTICIPANT SEQUENCE SETUP ==================
# =========================================================
# This version reads Sheet 3 ("Full sequence") and runs all testable
# conditions in order. Instructions such as "Visit Cell 1", "Install ...",
# and "END of GLASS ..." are shown as popups only. Each actual condition
# gets one full acuity + contrast + color + survey test, and the result row
# saves the condition name.

participant_id = None
participant_email = ""
full_sequence_text = ""
sequence_items = []
current_sequence_index = -1
current_cell = ""
current_glass = ""
completed_test_count = 0  # counts only real TEST conditions, not instructions
record = {h: "" for h in HEADERS}
progress_win = None
progress_listbox = None


def _normalize_email(local_or_email):
    local_or_email = (local_or_email or "").strip()
    if not local_or_email:
        return ""
    if "@" in local_or_email:
        return local_or_email
    return local_or_email + DOMAIN


def _design_file_missing_message():
    """A plain-language 'how to fix it' message for non-technical operators."""
    return (
        "Could not find the experiment design file:\n\n"
        f"    {DESIGN_XLSX_NAME}\n\n"
        "This file must be in the SAME folder as the program.\n"
        "The program looked here:\n\n"
        f"    {DESIGN_XLSX_PATH.parent}\n\n"
        "HOW TO FIX:\n"
        f"  1. Make sure '{DESIGN_XLSX_NAME}' is in the folder shown\n"
        "     above, right next to 'Run Experiment.bat'.\n"
        "  2. This file normally downloads with the rest of the project.\n"
        "     If it is missing, re-download the project (see\n"
        "     SETUP_INSTRUCTIONS.md) or ask the study coordinator for it.\n"
        "  3. Do not rename the file.\n\n"
        "TIP: You can still use 'Practice (Demo)' on the start screen\n"
        "without this file."
    )


def _read_full_sequence_from_design(participant_id_value):
    """Read one participant's Full sequence from the Excel design file."""
    if not DESIGN_XLSX_PATH.exists():
        raise FileNotFoundError(_design_file_missing_message())

    wb = load_workbook(DESIGN_XLSX_PATH, data_only=True)
    if DESIGN_SHEET not in wb.sheetnames:
        raise ValueError(
            f"The design file '{DESIGN_XLSX_NAME}' is missing the required sheet "
            f"named '{DESIGN_SHEET}'.\n\n"
            "HOW TO FIX: make sure you are using the correct, unmodified design\n"
            "file from the project (do not rename its sheets), or ask the study\n"
            "coordinator for a fresh copy."
        )

    ws = wb[DESIGN_SHEET]
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    if DESIGN_PARTICIPANT_COL not in headers or DESIGN_SEQUENCE_COL not in headers:
        raise ValueError(
            f"The design sheet must contain columns '{DESIGN_PARTICIPANT_COL}' and '{DESIGN_SEQUENCE_COL}'."
        )

    pid_col = headers.index(DESIGN_PARTICIPANT_COL) + 1
    seq_col = headers.index(DESIGN_SEQUENCE_COL) + 1

    for r in range(2, ws.max_row + 1):
        pid_val = ws.cell(r, pid_col).value
        if str(pid_val).strip() == str(participant_id_value).strip():
            full_seq = ws.cell(r, seq_col).value
            if not full_seq:
                raise ValueError(f"Full sequence is blank for participant {participant_id_value}.")
            return str(full_seq)

    raise ValueError(f"Participant ID {participant_id_value} was not found in Sheet 3.")


def _split_full_sequence(full_sequence):
    return [p.strip() for p in str(full_sequence).split("->") if p and p.strip()]


def is_instruction_item(item):
    item = (item or "").strip()
    return (
        item.startswith("Visit Cell") or
        item.startswith("Install ") or
        item.startswith("END of GLASS")
    )


def _update_context_from_instruction(item):
    """Track current cell/glass while walking through Full sequence."""
    global current_cell, current_glass
    item = (item or "").strip()
    if item.startswith("Visit Cell"):
        current_cell = item.replace("Visit ", "").strip()
        if current_cell == "Cell 2":
            current_glass = "glass only"
    elif item.startswith("Install "):
        current_glass = item.replace("Install ", "").strip()
    elif item.startswith("END of GLASS") and "return to" in item:
        current_glass = item.split("return to", 1)[1].strip()


_setup_typed = {}  # participant setup text kept across Alt+Left / Alt+Right


def prompt_participant_sequence_setup(root_):
    """Ask participant ID and email once at the start.

    Returns the setup dict, None if cancelled, or NAV_BACK for Alt+Left.
    """
    dialog = tk.Toplevel(root_)
    dialog.title("Participant Setup")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root_)

    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    dpi = get_layout_dpi(root_.winfo_id())
    win_w, win_h = set_dialog_half_screen(dialog, root_, frac=0.72)

    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.after(150, lambda: dialog.attributes("-topmost", False))

    label_font = ("Segoe UI", 44)
    entry_font = ("Segoe UI", 44)

    frame = tk.Frame(dialog, bg=BG_SOFT, padx=40, pady=40)
    frame.place(relx=0.5, rely=0.5, anchor="center")

    tk.Label(frame, text="Participant ID", font=label_font, bg=BG_SOFT,
             wraplength=int(9.0 * dpi), justify="center")        .pack(anchor="center", pady=(0, 12))
    pid_entry = tk.Entry(frame, font=entry_font, justify="center")
    pid_entry.pack(anchor="center", pady=(0, 34), ipady=6)

    tk.Label(frame, text="Berkeley email ID", font=label_font, bg=BG_SOFT,
             wraplength=int(9.0 * dpi), justify="center")        .pack(anchor="center", pady=(0, 12))
    email_entry = tk.Entry(frame, font=entry_font, justify="center")
    email_entry.pack(anchor="center", pady=(0, 34), ipady=6)

    # Keep what was typed if the researcher comes back to this page with Alt+Left.
    pid_entry.insert(0, _setup_typed.get("pid", ""))
    email_entry.insert(0, _setup_typed.get("email", ""))

    result = {"data": None}

    def submit():
        try:
            pid = int(pid_entry.get().strip())
        except ValueError:
            messagebox.showerror("Invalid", "Participant ID must be a number.", parent=dialog)
            return

        email = _normalize_email(email_entry.get())
        if not email:
            messagebox.showerror("Missing", "Enter Berkeley email ID.", parent=dialog)
            return

        try:
            full_seq = _read_full_sequence_from_design(pid)
        except Exception as e:
            messagebox.showerror("Design file error", str(e), parent=dialog)
            return

        result["data"] = {
            "participant_id": pid,
            "email": email,
            "full_sequence": full_seq,
            "items": _split_full_sequence(full_seq),
        }
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    def cancel():
        result["data"] = None
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    btns = tk.Frame(frame, bg=BG_SOFT)
    btns.pack(anchor="center", pady=(12, 0))
    tk.Button(btns, text="Start", font=("Segoe UI", 22, "bold"), width=10, command=submit)        .pack(side="left", padx=(0, 14))
    tk.Button(btns, text="Cancel", font=("Segoe UI", 22), width=10, command=cancel)        .pack(side="left")

    def on_nav(direction):
        if direction == NAV_NEXT:
            submit()
            return
        _setup_typed["pid"] = pid_entry.get()
        _setup_typed["email"] = email_entry.get()
        result["data"] = NAV_BACK
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    set_nav_page(on_nav, dialog)
    dialog.bind("<Return>", lambda e: submit())
    dialog.bind("<Escape>", lambda e: cancel())
    dialog.after(120, lambda: pid_entry.focus_force())
    root_.wait_window(dialog)
    clear_nav_page()
    return result["data"]


def _entry_text(value):
    """Format a stored value for an entry box (1000.0 -> "1000")."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def prompt_illuminance_for_condition(root_, condition_text, prefill=None, can_go_back=True):
    """Ask H/V illuminance before each condition. Email is not asked again.

    `prefill` = (H, V, Sun) to show when the condition is revisited with Alt+Left.
    Returns {"H", "V", "Sun"}, None for Exit, or NAV_BACK for Alt+Left.
    """
    dialog = tk.Toplevel(root_)
    dialog.title("Next Condition")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root_)

    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    dpi = get_layout_dpi(root_.winfo_id())

    # Keep the Next Condition window unmistakably centered on the current
    # duplicated display.  Do not inherit stale x/y coordinates from a monitor
    # that was previously connected.
    root_.update_idletasks()
    dialog.update_idletasks()
    try:
        mon = get_monitor_rect_for_window(root_.winfo_id())
        mon_w = max(1, mon.right - mon.left)
        mon_h = max(1, mon.bottom - mon.top)

        # Use 90% of the current work area so the dialog is easy to locate and
        # there is a visible margin around all four sides.
        win_w = int(mon_w * 0.90)
        win_h = int(mon_h * 0.90)
        x = mon.left + (mon_w - win_w) // 2
        y = mon.top + (mon_h - win_h) // 2

        dialog.minsize(1, 1)
        dialog.maxsize(100000, 100000)
        dialog.geometry(f"{win_w}x{win_h}+{x}+{y}")
        dialog.update_idletasks()
        dialog.minsize(win_w, win_h)
        dialog.maxsize(win_w, win_h)

        # Marker used by the monitor hot-plug reflow logic below.
        dialog._keep_centered_on_monitor = True
    except Exception:
        win_w, win_h = 1100, 760
        sw = max(1, dialog.winfo_screenwidth())
        sh = max(1, dialog.winfo_screenheight())
        x = max(0, (sw - win_w) // 2)
        y = max(0, (sh - win_h) // 2)
        dialog.geometry(f"{win_w}x{win_h}+{x}+{y}")
        dialog._keep_centered_on_monitor = True

    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.after(150, lambda: dialog.attributes("-topmost", False))

    label_font = ("Segoe UI", 38)
    entry_font = ("Segoe UI", 38)

    frame = tk.Frame(dialog, bg=BG_SOFT, padx=26, pady=22)
    frame.place(relx=0.5, rely=0.5, anchor="center")
    frame.columnconfigure(0, weight=0)
    frame.columnconfigure(1, weight=1)

    tk.Label(frame, text="Condition:", font=("Segoe UI", 38, "bold"), bg=BG_SOFT)        .grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
    tk.Label(frame, text=condition_text, font=("Segoe UI", 38, "bold"), bg=CURRENT_YELLOW,
             wraplength=int(9.5 * dpi), justify="left")        .grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, 24))

    tk.Label(frame, text="H. illuminance lux", font=label_font, bg=BG_SOFT)        .grid(row=2, column=0, sticky="w", pady=16, padx=(0, 22))
    h_entry = tk.Entry(frame, font=entry_font)
    h_entry.grid(row=2, column=1, sticky="ew", pady=16)

    tk.Label(frame, text="V. illuminance lux", font=label_font, bg=BG_SOFT)        .grid(row=3, column=0, sticky="w", pady=16, padx=(0, 22))
    v_entry = tk.Entry(frame, font=entry_font)
    v_entry.grid(row=3, column=1, sticky="ew", pady=16)

    tk.Label(frame, text="Sun position (inch)", font=label_font, bg=BG_SOFT)        .grid(row=4, column=0, sticky="w", pady=16, padx=(0, 22))

    def _sun_text_allowed(proposed):
        # Sun position: only a number (digits, one decimal point) or "NA".
        return (re.fullmatch(r"\d*\.?\d*", proposed) is not None
                or proposed.lower() in ("n", "na"))

    sun_entry = tk.Entry(frame, font=entry_font, validate="key",
                         validatecommand=(dialog.register(_sun_text_allowed), "%P"))
    sun_entry.grid(row=4, column=1, sticky="ew", pady=16)

    if prefill:
        for entry, value in zip((h_entry, v_entry, sun_entry), prefill):
            entry.insert(0, _entry_text(value))

    result = {"data": None}

    def parse_float(s):
        s = (s or "").strip()
        if s == "":
            return None
        return float(s)

    def parse_sun(s):
        # Sun position is a number or "NA" (e.g. overcast sky). Anything else -> ValueError.
        s = (s or "").strip()
        if s == "":
            return None
        if s.lower() == "na":
            return "NA"
        return float(s)

    def submit():
        try:
            h_val = parse_float(h_entry.get())
            v_val = parse_float(v_entry.get())
        except ValueError:
            messagebox.showerror("Invalid", "Illuminance must be numeric.", parent=dialog)
            return
        try:
            sun_val = parse_sun(sun_entry.get())
        except ValueError:
            messagebox.showerror("Invalid", "Sun position must be a number or NA.", parent=dialog)
            return
        if h_val is None or v_val is None or sun_val is None:
            messagebox.showerror(
                "Missing",
                "Enter H illuminance, V illuminance, and Sun position.\n"
                "Sun position can be a number or NA.",
                parent=dialog,
            )
            return
        result["data"] = {"H": h_val, "V": v_val, "Sun": sun_val}
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    def cancel():
        result["data"] = None
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    btns = tk.Frame(frame, bg=BG_SOFT)
    btns.grid(row=5, column=1, sticky="w", pady=(30, 0))
    tk.Button(btns, text="Start test", font=("Segoe UI", 26, "bold"), width=12, command=submit)        .pack(side="left", padx=(0, 14))
    tk.Button(btns, text="Exit ✕", font=("Segoe UI", 26, "bold"), width=10, command=cancel,
              bg="#C00000", fg="white", activebackground="#C00000", activeforeground="white")        .pack(side="left")

    # Esc does NOT quit here — it is ignored on the illuminance screen. Exit only
    # via the red "Exit ✕" button (or Alt+F4).
    def on_escape(_e=None):
        messagebox.showinfo(
            "Esc disabled",
            "Esc does not exit here. Use the red \"Exit ✕\" button to leave.",
            parent=dialog,
        )
        return "break"

    def on_nav(direction):
        if direction == NAV_NEXT:
            submit()
            return
        if not can_go_back:
            return
        result["data"] = NAV_BACK
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    set_nav_page(on_nav, dialog)
    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.bind("<Return>", lambda e: submit())
    dialog.bind("<Escape>", on_escape)  # show a message instead of quitting
    dialog.after(120, lambda: h_entry.focus_force())
    root_.wait_window(dialog)
    clear_nav_page()
    return result["data"]


def show_instruction_popup(item, can_go_back=True, show_glass=True):
    """Show one instruction page. Returns NAV_NEXT (Continue) or NAV_BACK (Alt+Left).

    show_glass=False hides the "Current glass should be" line (researcher reminders).
    """
    _update_context_from_instruction(item)
    nav = {"result": NAV_NEXT}

    dialog = tk.Toplevel(root)
    dialog.title("Instruction")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root)
    # Instruction pages (e.g. Visit Cell C) must remain centered across monitor swaps.
    dialog._keep_centered_on_monitor = True

    try:
        dialog.grab_set()
    except Exception:
        pass

    # ===== Large centered window =====
    try:
        mon = get_monitor_rect_for_window(root.winfo_id())
        mon_w = mon.right - mon.left
        mon_h = mon.bottom - mon.top

        win_w = int(mon_w * 0.8)
        win_h = int(mon_h * 0.8)

        x = mon.left + (mon_w - win_w) // 2
        y = mon.top + (mon_h - win_h) // 2

        dialog.geometry(f"{win_w}x{win_h}+{x}+{y}")
    except Exception:
        win_w, win_h = 1200, 800
        dialog.geometry(f"{win_w}x{win_h}")

    dialog.minsize(win_w, win_h)
    dialog.maxsize(win_w, win_h)

    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.after(200, lambda: dialog.attributes("-topmost", False))

    # ===== FONT SCALING (≈ 3x) =====
    scale = win_h / 800  # base reference

    title_size = int(52 * scale)
    text_size = int(46 * scale)
    sub_size = int(40 * scale)
    glass_size = int(48 * scale)
    button_size = int(30 * scale)

    def close():
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    frame = tk.Frame(dialog, bg=BG_SOFT, padx=50, pady=50)
    frame.pack(fill="both", expand=True)

    # Center all instruction text in the middle of the screen.
    inner = tk.Frame(frame, bg=BG_SOFT)
    inner.place(relx=0.5, rely=0.5, anchor="center")

    # Baseline-glass installs get a single clean line: "Install Baseline Glass <name>".
    is_baseline = "[BASELINE GLASS" in item
    # "Visit Cell X" moves to a new cell: don't show a stale "current glass" line.
    is_visit_cell = item.strip().startswith("Visit Cell")
    if is_baseline:
        _name = item.split("[BASELINE GLASS", 1)[0].strip()
        if _name.startswith("Install "):
            _name = _name[len("Install "):].strip()
        main_text = "Install Baseline Glass " + _name
    else:
        main_text = item

    # ===== Title (hidden for baseline-glass installs) =====
    if not is_baseline:
        tk.Label(
            inner,
            text="INSTRUCTION",
            font=("Segoe UI", title_size, "bold"),
            bg=BG_SOFT,
            justify="center",
            anchor="center"
        ).pack(anchor="center", pady=(0, int(40 * scale)))

    # ===== Main message =====
    tk.Label(
        inner,
        text=main_text,
        font=("Segoe UI", text_size),
        bg=BG_SOFT,
        wraplength=int(win_w * 0.75),
        justify="center",
        anchor="center"
    ).pack(anchor="center", pady=(0, int(40 * scale)))

    # ===== Current glass (hidden for baseline-glass installs and cell visits) =====
    if show_glass and current_glass and not is_baseline and not is_visit_cell:
        tk.Label(
            inner,
            text="Current glass should be:",
            font=("Segoe UI", sub_size),
            bg=BG_SOFT,
            justify="center",
            anchor="center"
        ).pack(anchor="center", pady=(10, int(15 * scale)))

        tk.Label(
            inner,
            text=current_glass,
            font=("Segoe UI", glass_size, "bold"),
            fg="#C00000",
            bg="#FFF2CC",
            padx=int(30 * scale),
            pady=int(15 * scale),
            justify="center",
            anchor="center"
        ).pack(anchor="center", pady=(0, int(30 * scale)))

    # ===== Button (fixed bottom) =====
    button_frame = tk.Frame(dialog, bg=BG_SOFT)
    button_frame.pack(side="bottom", fill="x", pady=int(35 * scale))

    continue_btn = tk.Button(
        button_frame,
        text="Continue  (Press Enter)",
        font=("Segoe UI", button_size, "bold"),
        width=22,
        command=close
    )
    continue_btn.pack(anchor="center")

    # Default focus
    continue_btn.focus_force()
    continue_btn.configure(default="active")

    def on_nav(direction):
        if direction == NAV_BACK and not can_go_back:
            return
        nav["result"] = direction
        close()

    set_nav_page(on_nav, dialog)
    dialog.bind("<Return>", lambda e: close())
    dialog.bind("<Escape>", lambda e: close())

    root.wait_window(dialog)
    clear_nav_page()
    return nav["result"]


def create_progress_window():
    global progress_win, progress_listbox
    progress_win = tk.Toplevel(root)
    progress_win.title("Experiment sequence progress")
    progress_win.configure(bg=BG_SOFT)
    progress_win.geometry("720x520+50+50")
    progress_win.protocol("WM_DELETE_WINDOW", lambda: save_and_exit())

    tk.Label(progress_win, text=f"Participant {participant_id} sequence", bg=BG_SOFT,
             font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=12, pady=(10, 4))
    progress_listbox = tk.Listbox(progress_win, font=("Segoe UI", 12), height=20)
    progress_listbox.pack(fill="both", expand=True, padx=12, pady=12)

    for idx, item in enumerate(sequence_items, start=1):
        label = "INSTRUCTION" if is_instruction_item(item) else "TEST"
        progress_listbox.insert("end", f"{idx:02d}. [{label}] {item}")


def mark_sequence_done(index0):
    if progress_listbox is None:
        return
    try:
        progress_listbox.itemconfig(index0, bg=DONE_GRAY)
        progress_listbox.see(index0)
    except Exception:
        pass


def mark_sequence_current(index0):
    if progress_listbox is None:
        return
    try:
        progress_listbox.itemconfig(index0, bg=CURRENT_YELLOW)
        progress_listbox.see(index0)
    except Exception:
        pass


def reset_condition_state():
    global stage, current_level_index, current_attempt, current_true_orientation
    global contrast_triplet_start_index, contrast_true_triplet, contrast_user_buffer
    global _contrast_triplet_n_correct, _contrast_finished, _color_finished, _survey_finished
    global _key_debounce

    stage = STAGE_ACUITY
    current_level_index = 0
    current_attempt = 1
    current_true_orientation = None
    contrast_triplet_start_index = 0
    contrast_true_triplet = ""
    contrast_user_buffer = ""
    _contrast_triplet_n_correct = None
    _contrast_finished = False
    _color_finished = False
    _survey_finished = False
    _key_debounce = False
    canvas.delete("all")


# Each test condition keeps its own result record, so Alt+Left can reopen a condition
# that was already finished and its Excel row is updated instead of duplicated.
records_by_index = {}     # Full-sequence index -> that condition's result record
excel_rows_by_index = {}  # Full-sequence index -> its row number in the results workbook
record_seq_idx = None     # Full-sequence index the current `record` belongs to


def mark_sequence_left(index0):
    """Item `index0` was left with Alt+Left: saved items stay gray, others go back to white."""
    if progress_listbox is None:
        return
    try:
        progress_listbox.itemconfig(index0, bg=DONE_GRAY if index0 in excel_rows_by_index else "white")
    except Exception:
        pass


def _rebuild_context_before(index0):
    """Recompute the current cell / glass from every item before `index0`.

    Needed because Alt+Left / Alt+Right can jump over the instructions that set them.
    """
    global current_cell, current_glass
    current_cell = ""
    current_glass = ""
    for it in sequence_items[:index0]:
        if is_instruction_item(it):
            _update_context_from_instruction(it)
        elif it == "glass only" and not current_glass:
            current_glass = "glass only"


def _test_order_for_index(index0):
    """1-based order of this condition among the TEST items (instructions not counted)."""
    return sum(1 for it in sequence_items[:index0 + 1] if not is_instruction_item(it))


def start_next_sequence_item():
    """Continue with the item after the current one in the Full sequence."""
    go_to_sequence_item(current_sequence_index + 1)


def go_to_sequence_item(index0, from_back=False):
    """Show item `index0` of the Full sequence.

    from_back=True means we arrived with Alt+Left from the item after it, so a
    test condition that was already finished reopens on its LAST page (the final
    survey question) rather than starting over.
    """
    global current_sequence_index

    current_sequence_index = max(0, index0)

    if current_sequence_index >= len(sequence_items):
        messagebox.showinfo("Complete", f"All conditions are complete for participant {participant_id}.", parent=root)
        try:
            if progress_win is not None:
                progress_win.destroy()
        except Exception:
            pass
        try:
            arrow_win.destroy()
        except Exception:
            pass
        root.destroy()
        return

    item = sequence_items[current_sequence_index]
    _rebuild_context_before(current_sequence_index)
    mark_sequence_current(current_sequence_index)

    if is_instruction_item(item):
        _run_instruction_page(current_sequence_index)
    elif from_back and current_sequence_index in records_by_index:
        _resume_condition_at_end(current_sequence_index)
    else:
        _run_illuminance_page(current_sequence_index)


def _run_instruction_page(index0):
    nav = show_instruction_popup(sequence_items[index0], can_go_back=index0 > 0)
    if nav == NAV_BACK:
        mark_sequence_left(index0)
        root.after(100, lambda: go_to_sequence_item(index0 - 1, from_back=True))
        return
    mark_sequence_done(index0)
    root.after(100, lambda: go_to_sequence_item(index0 + 1))


def _run_illuminance_page(index0):
    """First page of a test condition: illuminance / sun entry, then the tests."""
    global record, record_seq_idx, current_glass, completed_test_count
    global H_ill, V_ill, Sun_position_inch

    item = sequence_items[index0]
    if item == "glass only" and not current_glass:
        current_glass = "glass only"

    existing = records_by_index.get(index0)
    prefill = None
    if existing is not None:
        prefill = (existing.get("H_illuminance_lux"), existing.get("V_illuminance_lux"),
                   existing.get("Sun_position_inch"))

    inputs = prompt_illuminance_for_condition(root, item, prefill=prefill, can_go_back=index0 > 0)
    if inputs == NAV_BACK:
        mark_sequence_left(index0)
        root.after(100, lambda: go_to_sequence_item(index0 - 1, from_back=True))
        return
    if not inputs:
        root.destroy()
        return

    H_ill = inputs["H"]
    V_ill = inputs["V"]
    Sun_position_inch = inputs["Sun"]

    reset_condition_state()

    if existing is None:
        # Count only actual test conditions.
        # Instructions such as "Visit Cell 1", "Install glass C",
        # and "END of GLASS" are excluded from this order.
        rec = {h: "" for h in HEADERS}
        rec["participant_id"] = participant_id
        rec["sequence_step_order"] = _test_order_for_index(index0)
        rec["condition"] = item
        rec["raw_full_sequence_position"] = index0 + 1
        rec["full_sequence"] = full_sequence_text
        rec["participant_email"] = participant_email
        rec["screening_visual_acuity_20_20_correct_of_2"] = screening_results.get("va_correct", "")
        rec["screening_visual_acuity_20_20_pass"] = screening_results.get("va_pass", "")
        for _i in range(2):
            _r = screening_results.get("ishihara", [{}, {}])[_i]
            rec[f"screening_ishihara_{_i+1}_response"] = _r.get("response", "")
            rec[f"screening_ishihara_{_i+1}_pass"] = _r.get("pass", "")
        rec["start_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        records_by_index[index0] = rec

    record = records_by_index[index0]
    record_seq_idx = index0
    completed_test_count = record["sequence_step_order"]
    record["H_illuminance_lux"] = H_ill
    record["V_illuminance_lux"] = V_ill
    record["Sun_position_inch"] = Sun_position_inch

    enter_acuity()


def _resume_condition_at_end(index0):
    """Reopen a finished condition (reached with Alt+Left) on its last survey question."""
    global record, record_seq_idx, completed_test_count
    reset_condition_state()
    record = records_by_index[index0]
    record_seq_idx = index0
    completed_test_count = record["sequence_step_order"]
    set_condition_hud(SURVEY_HUD)
    enter_survey(start_q=len(SURVEY_QUESTIONS) - 1)


# =========================================================
# ================= PRE-SESSION SCREENING ==================
# =========================================================
screening_results = {
    "va_correct": "",
    "va_pass": "",
    "ishihara": [
        {"response": "", "pass": ""},
        {"response": "", "pass": ""},
    ],
}


def _script_folder():
    try:
        return Path(__file__).resolve().parent
    except Exception:
        return Path.cwd()


def _show_screening_intro(parent):
    """Returns NAV_NEXT (Start Screening) or NAV_BACK (Alt+Left: back to the start screen)."""
    dialog = tk.Toplevel(parent)
    dialog.title("Screening")
    dialog.configure(bg=BG_SOFT)
    dialog.transient(parent)
    dialog.resizable(False, False)
    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    win_w, win_h = set_dialog_half_screen(dialog, parent, frac=0.80)
    wrap = tk.Frame(dialog, bg=BG_SOFT, padx=55, pady=45)
    wrap.pack(fill="both", expand=True)

    tk.Label(wrap, text="Screening", bg=BG_SOFT,
             font=("Segoe UI", 44, "bold")).pack(anchor="center", pady=(30, 18))
    tk.Label(
        wrap,
        text=f"1. 20/{SCREENING_ACUITY_DENOMINATOR} visual acuity screening (2 trials)\n2. Two Ishihara color-vision plates",
        bg=BG_SOFT, font=("Segoe UI", 28), justify="center"
    ).pack(anchor="center", pady=(0, 30))

    nav = {"result": NAV_NEXT}

    def on_nav(direction):
        nav["result"] = direction
        dialog.destroy()

    start_btn = tk.Button(wrap, text="Start Screening", font=("Segoe UI", 28, "bold"),
                          width=18, height=2, takefocus=True, command=dialog.destroy)
    start_btn.pack(anchor="center", pady=20)
    set_nav_page(on_nav, dialog)
    dialog.bind("<Return>", lambda e: start_btn.invoke())
    dialog.bind("<Escape>", lambda e: "break")
    dialog.after(100, start_btn.focus_force)
    parent.wait_window(dialog)
    clear_nav_page()
    return nav["result"]


def _run_20_20_screening(parent):
    """Two Landolt-C presentations at 20/SCREENING_ACUITY_DENOMINATOR; show feedback after each.

    Returns (nav, correct, passed). correct/passed are None when the page was
    left with Alt+Left / Alt+Right before both presentations were answered.
    """
    dialog = tk.Toplevel(parent)
    dialog.title(f"Screening - 20/{SCREENING_ACUITY_DENOMINATOR} Visual Acuity")
    dialog.configure(bg=BG_SOFT)
    try:
        dialog.attributes("-fullscreen", True)
    except Exception:
        pass
    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    top = tk.Frame(dialog, bg=BG_SOFT)
    top.pack(fill="x", padx=24, pady=18)
    title_var = tk.StringVar(value=f"Screening: 20/{SCREENING_ACUITY_DENOMINATOR} Visual Acuity")
    progress = tk.StringVar(value="")
    tk.Label(top, textvariable=title_var, bg=BG_SOFT,
             font=("Segoe UI", 25, "bold")).pack(side="left")
    tk.Label(top, textvariable=progress, bg=BG_SOFT,
             font=("Segoe UI", 20)).pack(side="right")
    tk.Label(dialog,
             text="Press the arrow key matching the Landolt-C gap direction.",
             bg=BG_SOFT, font=("Segoe UI", 20)).pack(pady=(0, 8))

    c = tk.Canvas(dialog, bg=BG_SOFT, highlightthickness=0)
    c.pack(fill="both", expand=True)

    state = {"trial": 0, "correct": 0, "orientation": None, "done": False}

    def draw_trial():
        if not dialog.winfo_exists():
            return  # page was left with Alt+arrows
        if state["trial"] >= SCREENING_ACUITY_TRIALS:
            state["done"] = True
            passed = state["correct"] >= SCREENING_ACUITY_PASS_N
            result_text = (
                f"20/{SCREENING_ACUITY_DENOMINATOR} screening complete\n\n"
                f"Correct: {state['correct']} / {SCREENING_ACUITY_TRIALS}\n"
                f"Result: {'PASS' if passed else 'DID NOT PASS'}"
            )
            c.delete("all")
            c.create_text(c.winfo_width()/2, c.winfo_height()*0.42,
                          text=result_text, font=("Segoe UI", 34, "bold"),
                          fill="black", justify="center")
            c.create_text(c.winfo_width()/2, c.winfo_height()*0.67,
                          text="Press Enter to continue to Ishihara screening.",
                          font=("Segoe UI", 22), fill="gray25")
            return

        state["orientation"] = random.choice(SCREENING_ORIENTATIONS)
        progress.set(f"Trial {state['trial'] + 1} / {SCREENING_ACUITY_TRIALS}")
        dpi = get_dpi_for_window(dialog.winfo_id())
        draw_landolt_c(c, dpi, VIEWING_DISTANCE_M, SCREENING_ACUITY_DENOMINATOR, state["orientation"])

    feedback_active = {"value": False}

    def show_feedback(is_correct):
        feedback_active["value"] = True
        c.delete("all")
        c.create_text(
            c.winfo_width() / 2,
            c.winfo_height() / 2,
            text="Correct" if is_correct else "Incorrect",
            font=("Segoe UI", 54, "bold"),
            fill="black",
            justify="center"
        )

        def continue_after_feedback():
            if not dialog.winfo_exists():
                return  # page was left with Alt+arrows during the feedback
            feedback_active["value"] = False
            # After the final screening response, go directly to Ishihara.
            # No intermediate acuity-complete/transition page is shown.
            if state["trial"] >= SCREENING_ACUITY_TRIALS:
                dialog.destroy()
            else:
                draw_trial()

        dialog.after(700, continue_after_feedback)

    def on_key(event):
        if feedback_active["value"]:
            return "break"
        if state["done"]:
            if event.keysym in ("Return", "KP_Enter"):
                dialog.destroy()
                return "break"
            return "break"
        if event.keysym not in SCREENING_KEYSYM_TO_ORIENTATION:
            return None
        ans = SCREENING_KEYSYM_TO_ORIENTATION[event.keysym]
        is_correct = (ans == state["orientation"])
        if is_correct:
            state["correct"] += 1
        state["trial"] += 1
        show_feedback(is_correct)
        return "break"

    nav = {"result": None}

    def on_nav(direction):
        nav["result"] = direction
        dialog.destroy()

    for key in ("<Up>", "<Down>", "<Left>", "<Right>", "<Return>"):
        dialog.bind(key, on_key)
    bind_nav_keys(dialog)  # Alt+Left / Alt+Right must not count as an answer
    set_nav_page(on_nav, dialog)
    dialog.bind("<Escape>", lambda e: "break")
    dialog.protocol("WM_DELETE_WINDOW", lambda: None)
    dialog.after(200, draw_trial)
    parent.wait_window(dialog)
    clear_nav_page()

    if nav["result"] is not None:
        return nav["result"], None, None
    correct = state["correct"]
    return NAV_NEXT, correct, (correct >= SCREENING_ACUITY_PASS_N)



def _resolve_ishihara_image(filename):
    """Resolve an Ishihara image from the configured Downloads folder."""
    base = ISHIHARA_DIR / str(filename)
    if base.exists():
        return base
    if base.suffix:
        return base
    for ext in (".png", ".jpg", ".jpeg", ".bmp"):
        candidate = base.with_suffix(ext)
        if candidate.exists():
            return candidate
    return base

def _load_ishihara_photo(image_path, max_w, max_h):
    if not PIL_OK or not image_path.exists():
        return None
    img = Image.open(image_path).convert("RGB")
    img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS)
    return ImageTk.PhotoImage(img)


def _run_one_ishihara(parent, plate_index, config, previous=None):
    """One Ishihara plate page. Returns (nav, result).

    result is None when the page is left without a new answer (the previous answer
    is kept), otherwise {"response", "pass"}.
    """
    dialog = tk.Toplevel(parent)
    dialog.title(f"Screening - Ishihara {plate_index + 1}/{len(ISHIHARA_PLATES)}")
    dialog.configure(bg=BG_SOFT)
    dialog.transient(parent)
    dialog._keep_centered_on_monitor = True
    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    try:
        mon = get_monitor_rect_for_window(parent.winfo_id())
        w = int((mon.right-mon.left)*0.88)
        h = int((mon.bottom-mon.top)*0.88)
        x = mon.left + ((mon.right-mon.left)-w)//2
        y = mon.top + ((mon.bottom-mon.top)-h)//2
        dialog.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        w,h=1400,900
        dialog.geometry(f"{w}x{h}")

    wrap = tk.Frame(dialog, bg=BG_SOFT, padx=35, pady=25)
    wrap.pack(fill="both", expand=True)
    tk.Label(wrap, text=f"Ishihara Screening  {plate_index + 1} / {len(ISHIHARA_PLATES)}",
             bg=BG_SOFT, font=("Segoe UI", 32, "bold")).pack(pady=(0, 12))
    tk.Label(wrap, text="What number do you see?",
             bg=BG_SOFT, font=("Segoe UI", 25)).pack(pady=(0, 12))

    image_path = _resolve_ishihara_image(config.get("filename", ""))
    image_label = None
    if PIL_OK and image_path.exists():
        # The plate is sized below, once the space left for it is known.
        image_label = tk.Label(wrap, bg=BG_SOFT)
        image_label.pack(pady=(0, 16))
    else:
        tk.Label(
            wrap,
            text=f"Plate image not found:\n{image_path}\n\nPlace this image in the 'ishihara' folder next to the program.",
            bg="#FFF2CC", fg="#C00000", font=("Segoe UI", 19, "bold"),
            justify="center", padx=20, pady=20
        ).pack(pady=(10, 22))

    prev_response = str((previous or {}).get("response", "") or "")
    answer_var = tk.StringVar(value="" if prev_response in ("", "SKIPPED") else prev_response)
    entry = tk.Entry(wrap, textvariable=answer_var, font=("Segoe UI", 30), justify="center", width=18)
    entry.pack(ipady=8, pady=(0, 16))

    result = {"response": "", "pass": ""}
    nav = {"result": NAV_NEXT, "new_answer": False}

    def submit():
        ans = answer_var.get().strip()
        if not ans:
            messagebox.showerror("Missing response", "Enter the participant's response.", parent=dialog)
            return
        expected = config.get("expected_answer")
        result["response"] = ans
        if expected is None or str(expected).strip() == "":
            result["pass"] = "Not graded"
        else:
            result["pass"] = "PASS" if ans.casefold() == str(expected).strip().casefold() else "DID NOT PASS"
        nav["new_answer"] = True
        dialog.destroy()

    def on_nav(direction):
        if direction == NAV_BACK:
            nav["result"] = NAV_BACK
            dialog.destroy()
            return
        if answer_var.get().strip():
            submit()
            return
        if not prev_response:
            # Moving past a plate that was never answered.
            result["response"] = "SKIPPED"
            result["pass"] = "SKIPPED"
            nav["new_answer"] = True
        dialog.destroy()

    submit_btn = tk.Button(wrap, text="Continue", font=("Segoe UI", 24, "bold"),
                           width=14, height=2, command=submit, takefocus=True)
    submit_btn.pack(pady=(4, 0))

    if image_label is not None:
        # Give the plate the height left after the title, prompt, answer box and
        # Continue button, so the button is never pushed off the bottom of the window.
        dialog.update_idletasks()
        others_h = wrap.winfo_reqheight() - image_label.winfo_reqheight()
        # 24 px covers the label's own border/padding around the picture plus a margin.
        max_img_h = max(120, min(int(h * 0.55), h - others_h - 24))
        photo = _load_ishihara_photo(image_path, int(w * 0.52), max_img_h)
        if photo is not None:
            image_label.configure(image=photo)
            image_label.image = photo
    set_nav_page(on_nav, dialog)
    dialog.bind("<Return>", lambda e: submit())
    dialog.bind("<Escape>", lambda e: "break")
    dialog.protocol("WM_DELETE_WINDOW", lambda: None)
    dialog.after(120, entry.focus_force)
    parent.wait_window(dialog)
    clear_nav_page()
    return nav["result"], (dict(result) if nav["new_answer"] else None)


# Screening pages: 0 = intro, 1 = acuity screening, 2.. = Ishihara plates, last = summary.
SCREENING_SUMMARY_PAGE = 2 + len(ISHIHARA_PLATES)


def run_screening(parent, start_page=0):
    """Run the pre-session screening pages; Alt+Left / Alt+Right move between them.

    Returns NAV_BACK if Alt+Left is pressed on the first page, otherwise NAV_NEXT.
    """
    page = max(0, min(start_page, SCREENING_SUMMARY_PAGE))
    while True:
        if page == 0:
            nav = _show_screening_intro(parent)
        elif page == 1:
            nav, correct, passed = _run_20_20_screening(parent)
            if correct is not None:
                screening_results["va_correct"] = correct
                screening_results["va_pass"] = "PASS" if passed else "DID NOT PASS"
            elif nav == NAV_NEXT and screening_results["va_correct"] in ("", None):
                # Moving past a screening that was never completed.
                screening_results["va_correct"] = "SKIPPED"
                screening_results["va_pass"] = "SKIPPED"
        elif page < SCREENING_SUMMARY_PAGE:
            i = page - 2
            nav, res = _run_one_ishihara(parent, i, ISHIHARA_PLATES[i],
                                         previous=screening_results["ishihara"][i])
            if res is not None:
                screening_results["ishihara"][i] = res
        else:
            nav = _show_screening_summary(parent)

        if nav == NAV_BACK:
            if page == 0:
                return NAV_BACK
            page -= 1
        else:
            if page == SCREENING_SUMMARY_PAGE:
                return NAV_NEXT
            page += 1


def _show_screening_summary(parent):
    """Returns NAV_NEXT (Continue) or NAV_BACK (Alt+Left: back to the last plate)."""
    summary = tk.Toplevel(parent)
    summary.title("Screening Complete")
    summary.configure(bg=BG_SOFT)
    summary.transient(parent)
    try:
        summary.grab_set()
    except tk.TclError:
        pass
    set_dialog_half_screen(summary, parent, frac=0.72)
    wrap = tk.Frame(summary, bg=BG_SOFT, padx=45, pady=38)
    wrap.pack(fill="both", expand=True)
    tk.Label(wrap, text="Screening Complete", bg=BG_SOFT,
             font=("Segoe UI", 38, "bold")).pack(pady=(30, 20))
    if screening_results["va_correct"] == "SKIPPED":
        lines = [f"20/{SCREENING_ACUITY_DENOMINATOR} visual acuity: SKIPPED"]
    else:
        lines = [f"20/{SCREENING_ACUITY_DENOMINATOR} visual acuity: {screening_results['va_correct']}/{SCREENING_ACUITY_TRIALS} — {screening_results['va_pass']}"]
    for i, r in enumerate(screening_results["ishihara"], start=1):
        lines.append(f"Ishihara {i}: response = {r['response']}   ({r['pass']})")
    tk.Label(wrap, text="\n".join(lines), bg=BG_SOFT,
             font=("Segoe UI", 21), justify="left").pack(pady=(0, 25))
    nav = {"result": NAV_NEXT}

    def on_nav(direction):
        nav["result"] = direction
        summary.destroy()

    btn = tk.Button(wrap, text="Continue", font=("Segoe UI", 25, "bold"),
                    width=14, height=2, command=summary.destroy, takefocus=True)
    btn.pack()
    set_nav_page(on_nav, summary)
    summary.bind("<Return>", lambda e: btn.invoke())
    summary.bind("<Escape>", lambda e: "break")
    summary.protocol("WM_DELETE_WINDOW", lambda: None)
    summary.after(100, btn.focus_force)
    parent.wait_window(summary)
    clear_nav_page()
    return nav["result"]


# =========================================================
# ================ LAUNCHER / DEMO (RESEARCHER UI) ========
# =========================================================
def _demo_setup():
    """A self-contained practice run: no Excel design file or real participant needed."""
    items = [
        "Visit Cell A  (DEMO — practice run, results are not used)",
        "DEMO material 1 - clear glass",
        "DEMO material 2 - tinted glass",
    ]
    return {
        "participant_id": "DEMO",
        "email": "demo" + DOMAIN,
        "full_sequence": " -> ".join(items),
        "items": items,
    }


def prompt_launcher(root_):
    """Start screen for researchers. Returns 'demo', 'real', or None (quit)."""
    dialog = tk.Toplevel(root_)
    dialog.title("ASHRAE 1925-TRP - Start")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root_)
    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    # Larger start page: use 88% of the active monitor in both dimensions (82% left the
    # bottom of the hotkey list cut off on a 1920x1080 screen).
    set_dialog_half_screen(dialog, root_, frac=0.88)
    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.after(150, lambda: dialog.attributes("-topmost", False))

    choice = {"val": None}

    def choose(v):
        choice["val"] = v
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    wrap = tk.Frame(dialog, bg=BG_SOFT, padx=60, pady=44)
    wrap.pack(fill="both", expand=True)

    tk.Label(wrap, text="Visual Experiment", bg=BG_SOFT,
             font=("Segoe UI", 52, "bold")).pack(anchor="w")
    tk.Label(wrap, text="ASHRAE 1925-TRP   ·   Acuity → Contrast → Color → Survey",
             bg=BG_SOFT, fg="gray25", font=("Segoe UI", 24)).pack(anchor="w", pady=(4, 34))

    btns = tk.Frame(wrap, bg=BG_SOFT)
    btns.pack(anchor="w", pady=(0, 34))
    demo_btn = tk.Button(btns, text="Practice  (Demo)", font=("Segoe UI", 27, "bold"),
                         width=22, height=2, padx=12, pady=8, command=lambda: choose("demo"), takefocus=True)
    demo_btn.pack(side="left", padx=(0, 24))
    real_btn = tk.Button(btns, text="Start Real Session", font=("Segoe UI", 27, "bold"),
                         width=22, height=2, padx=12, pady=8, command=lambda: choose("real"), takefocus=True)
    real_btn.pack(side="left")
    # Quit sits in the same row: below the hotkey list it was pushed off a 1080-px screen.
    tk.Button(btns, text="Quit", font=("Segoe UI", 20), width=12, height=2,
              command=lambda: choose(None)).pack(side="left", padx=(24, 0))

    cheat = tk.Frame(wrap, bg="white", bd=3, relief="solid", padx=28, pady=22)
    cheat.pack(anchor="w", fill="x")
    tk.Label(cheat, text="Hotkeys (the same in every stage)", bg="white",
             font=("Segoe UI", 21, "bold")).pack(anchor="w", pady=(0, 12))
    for k, v in [
        ("Acuity", "Arrow keys  ← ↑ → ↓  =  direction of the gap"),
        ("Contrast", "Type the 3 letters (A-Z)    ·    Backspace = clear    ·    Enter = submit fewer"),
        ("Color", "Click / drag the wheel    ·    Enter = confirm"),
        ("Survey", "← / →  = move selection    ·    Enter = confirm"),
        ("Pages", "Alt+←  =  previous page    ·    Alt+→  =  next page"),
        ("SKIP", "F2  =  skip the current task (for those who cannot see)"),
        ("Esc", "Esc  =  skip the whole current task and move on (never quits)"),
        ("Exit", "Close (X) button on the window  =  end the session"),
    ]:
        line = tk.Frame(cheat, bg="white")
        line.pack(anchor="w", fill="x", pady=3)
        tk.Label(line, text=k, bg="white", fg="#C00000", width=12, anchor="w",
                 font=("Segoe UI", 18, "bold")).pack(side="left")
        tk.Label(line, text=v, bg="white", anchor="w",
                 font=("Segoe UI", 18)).pack(side="left")


    dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
    dialog.bind("<Escape>", lambda e: choose(None))

    # Start Real Session is the default selection. Pressing Enter immediately starts it.
    def _launcher_enter(_event=None):
        focused = dialog.focus_get()
        if isinstance(focused, tk.Button):
            focused.invoke()
        else:
            real_btn.invoke()
        return "break"

    def on_nav(direction):
        if direction == NAV_NEXT:  # the start screen is the first page: no previous page
            _launcher_enter()

    set_nav_page(on_nav, dialog)
    dialog.bind("<Return>", _launcher_enter)
    dialog.after(120, real_btn.focus_force)
    root_.wait_window(dialog)
    clear_nav_page()
    return choice["val"]


def run_pre_session():
    """Start screen -> screening -> participant setup, with Alt+arrow navigation.

    Returns the session setup dict, or None to quit.
    """
    while True:
        mode = prompt_launcher(root)
        if mode is None:
            return None
        if mode == "demo":
            return _demo_setup()
        screening_page = 0
        while True:
            if run_screening(root, start_page=screening_page) == NAV_BACK:
                break  # back to the start screen
            setup_ = prompt_participant_sequence_setup(root)
            if setup_ == NAV_BACK:
                screening_page = SCREENING_SUMMARY_PAGE
                continue
            return setup_


setup = run_pre_session()

if not setup:
    root.destroy()
    raise SystemExit

participant_id = setup["participant_id"]
participant_email = setup["email"]
full_sequence_text = setup["full_sequence"]
sequence_items = setup["items"]

# Save one result file per participant.
XLSX_PATH = DOWNLOADS_DIR / f"participant_{participant_id}_results.xlsx"
ensure_workbook(XLSX_PATH)
create_progress_window()

# =========================================================
# ========================= STAGES / STATE =================
# =========================================================
STAGE_ACUITY = "acuity"
STAGE_CONTRAST = "contrast"
STAGE_COLOR = "color"
STAGE_SURVEY = "survey"
stage = STAGE_ACUITY

current_level_index = 0
current_attempt = 1
current_true_orientation = None

contrast_triplet_start_index = 0
contrast_true_triplet = ""
contrast_user_buffer = ""
_contrast_triplet_n_correct = None  # last triplet correctness
_contrast_finished = False
_color_finished = False
_survey_finished = False

# for bind_all management
_contrast_bind_all = None

ORIENTATIONS = ["up", "down", "left", "right"]
KEYSYM_TO_ORIENTATION = {"Up": "up", "Down": "down", "Left": "left", "Right": "right"}


def set_hud(stage_name, progress_text="", instruction=""):
    stage_var.set(stage_name)
    progress_var.set(progress_text)
    instr_var.set(instruction)


def set_hud_instruction(instruction):
    """Change only the instruction line, keeping the condition and step shown above it."""
    instr_var.set(instruction)


# =========================================================
# ========================= DEBUG OVERLAY ==================
# =========================================================
def _contrast_current_triplet_idx():
    return triplet_index_from_letter_index(contrast_triplet_start_index)


def draw_debug_overlay():
    if not TEST:
        return
    canvas.delete("debug")

    lines = []

    if stage == STAGE_ACUITY:
        den = ACUITY_LEVELS[current_level_index] if 0 <= current_level_index < len(ACUITY_LEVELS) else None
        lines.append(f"TEST: Acuity stage = 20/{den}")
        lines.append(f"attempt={current_attempt}  level_index={current_level_index}")

    if stage == STAGE_CONTRAST:
        tidx = _contrast_current_triplet_idx()
        lines.append(f"TEST: Contrast triplet_idx={tidx}")
        lines.append(f"hex={hex_for_triplet(tidx)}  logCS={logcs_for_triplet(tidx)}")
        lines.append(f"start_index={contrast_triplet_start_index} buffer='{contrast_user_buffer}'")

    if stage == STAGE_COLOR:
        lines.append(f"TEST: Color target={record.get('color_target_hex','')}")
        lines.append(f"selected={record.get('color_selected_hex','')}")

    # draw box
    text = "\n".join(lines)
    if not text:
        return

    canvas.update_idletasks()
    W = max(10, canvas.winfo_width())

    pad = 10
    x0, y0 = pad, pad
    box_id = canvas.create_rectangle(x0, y0, x0 + 10, y0 + 10,
                                    fill="#FFFFFF", outline="black", width=1, tags=("debug",))
    tid = canvas.create_text(x0 + 8, y0 + 6, text=text, anchor="nw",
                            font=("Segoe UI", 10, "bold"), fill="black", tags=("debug",))
    bb = canvas.bbox(tid)
    if bb:
        canvas.coords(box_id, bb[0] - 6, bb[1] - 6, bb[2] + 6, bb[3] + 6)


# =========================================================
# ========================= CONTRAST BIND_ALL ==============
# =========================================================
def enable_contrast_typing_bind_all():
    global _contrast_bind_all
    if _contrast_bind_all is not None:
        return
    _contrast_bind_all = root.bind_all("<KeyPress>", on_key_contrast_triplets)


def disable_contrast_typing_bind_all():
    global _contrast_bind_all
    if _contrast_bind_all is None:
        return
    try:
        root.unbind_all("<KeyPress>")
    except Exception:
        pass
    _contrast_bind_all = None


# =========================================================
# ========================= ARROW WINDOW ===================
# =========================================================
arrow_win = tk.Toplevel(root)
arrow_win.title("Arrows")
arrow_win.configure(bg=BG_SOFT)
arrow_win.resizable(False, False)
arrow_win.transient(root)
arrow_win.attributes("-topmost", True)
arrow_win.after(200, lambda: arrow_win.attributes("-topmost", False))
arrow_win.protocol("WM_DELETE_WINDOW", lambda: save_and_exit())

arrow_frame = tk.Frame(
    arrow_win,
    bg=BG_SOFT,
    bd=2,
    relief="solid",
    highlightthickness=2,
    highlightbackground="black",
    padx=10,
    pady=8
)
arrow_frame.pack(fill="both", expand=True)

arrow_title = tk.Label(arrow_frame, text="Use arrow keys", bg=BG_SOFT, fg="black", font=ARROW_TITLE_FONT)
arrow_title.grid(row=0, column=0, columnspan=3, padx=6, pady=(2, 6))


def _arrow_label(txt):
    return tk.Label(
        arrow_frame,
        text=txt,
        bg=ARROW_BG,
        fg="black",
        width=4,
        height=2,
        font=ARROW_FONT,
        relief="solid",
        bd=2
    )


arrow_up = _arrow_label("↑")
arrow_left = _arrow_label("←")
arrow_down = _arrow_label("↓")
arrow_right = _arrow_label("→")

arrow_up.grid(row=1, column=1, padx=6, pady=6)
arrow_left.grid(row=2, column=0, padx=6, pady=6)
arrow_down.grid(row=2, column=1, padx=6, pady=6)
arrow_right.grid(row=2, column=2, padx=6, pady=6)

_arrow_map = {"up": arrow_up, "down": arrow_down, "left": arrow_left, "right": arrow_right}
_arrow_reset_after_id = None

arrow_win.update_idletasks()
ARROW_WIN_W = arrow_win.winfo_width()
ARROW_WIN_H = arrow_win.winfo_height()
arrow_win.withdraw()


def flash_arrow(direction, ms=220):
    global _arrow_reset_after_id
    if direction not in _arrow_map:
        return

    if _arrow_reset_after_id is not None:
        try:
            root.after_cancel(_arrow_reset_after_id)
        except Exception:
            pass
        _arrow_reset_after_id = None

    for lbl in _arrow_map.values():
        lbl.configure(bg=ARROW_BG)
    _arrow_map[direction].configure(bg=ARROW_HILITE)

    def reset():
        for lbl in _arrow_map.values():
            lbl.configure(bg=ARROW_BG)

    _arrow_reset_after_id = root.after(ms, reset)


def set_arrow_panel_visible(visible):
    if visible:
        arrow_win.deiconify()
        arrow_win.lift()
        activate_stimulus_window()
    else:
        arrow_win.withdraw()
        activate_stimulus_window()


def position_arrow_window_attached():
    try:
        root.update_idletasks()
        mon = get_monitor_rect_for_window(root.winfo_id())
        aw = ARROW_WIN_W or 260
        ah = ARROW_WIN_H or 220
        margin = 60
        x = mon.left + margin
        y = mon.top + max(0, ((mon.bottom - mon.top) - ah) // 2)
        arrow_win.geometry(f"+{x}+{y}")
    except Exception:
        pass


# =========================================================
# ========================= SAVE / EXIT ====================
# =========================================================
def finalize_and_save():
    record["end_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row = write_participant_row(XLSX_PATH, record, excel_rows_by_index.get(record_seq_idx))
    if record_seq_idx is not None:
        excel_rows_by_index[record_seq_idx] = row


def finish_current_condition_and_continue():
    # Called after Survey is complete. Save this condition, gray it out,
    # then automatically continue to the next item in Full sequence.
    finalize_and_save()
    mark_sequence_done(current_sequence_index)
    set_arrow_panel_visible(False)
    root.after(250, lambda: remind_laptop_on_table(start_next_sequence_item))


def save_and_exit():
    # Used only for early quit / ESC. Save the active condition if it started, then exit.
    if _screens_extended:
        request_display_mode("clone")  # leave Windows in Duplicate, as at the start
    try:
        if record and record.get("start_time"):
            finalize_and_save()
            if current_sequence_index >= 0:
                mark_sequence_done(current_sequence_index)
    except Exception:
        pass
    try:
        arrow_win.destroy()
    except Exception:
        pass
    try:
        if progress_win is not None:
            progress_win.destroy()
    except Exception:
        pass
    try:
        root.destroy()
    except Exception:
        pass


# =========================================================
# ========================= ACUITY: ACHIEVED/FAILED AT =====
# =========================================================
def _set_acuity_outcome(failed_at_den: int):
    """
    If failed at first stage: achieved = 20/firststage and failed_at = 20/firststage (per your request)
    Else: achieved = previous stage, failed_at = current stage
    """
    failed_at_den = int(failed_at_den)
    record["Visual acuity failed at"] = f"20/{failed_at_den}"

    if failed_at_den == ACUITY_LEVELS[0]:
        record["Visual acuity achieved"] = f"20/{ACUITY_LEVELS[0]}"
    else:
        try:
            idx = ACUITY_LEVELS.index(failed_at_den)
            prev_den = ACUITY_LEVELS[idx - 1]
        except Exception:
            prev_den = ACUITY_LEVELS[0]
        record["Visual acuity achieved"] = f"20/{prev_den}"


def end_acuity_and_move_on():
    enter_contrast()


def handle_acuity_failure(esc=False):
    """
    Unified failure handler used by:
      - 2 wrong answers at a level
      - ESC pressed in acuity stage

    Policy:
      - If ESC at first stage => fail at first stage
      - If ESC at later stages => fail at current stage (same as two-wrong),
        and achieved is previous stage.
    """
    global stage

    if stage != STAGE_ACUITY:
        return

    failed_at_den = ACUITY_LEVELS[current_level_index]
    _set_acuity_outcome(failed_at_den)

    # proceed
    stage = STAGE_CONTRAST
    set_arrow_panel_visible(False)
    end_acuity_and_move_on()


# =========================================================
# ========================= ACUITY STAGE ===================
# =========================================================
_key_debounce = False

def _show_acuity_stimulus():
    """Show the Landolt-C after the pre-stimulus blank interval, then enable input."""
    global current_true_orientation, _key_debounce, _acuity_show_after_id

    _acuity_show_after_id = None
    if stage != STAGE_ACUITY:
        return

    current_true_orientation = random.choice(ORIENTATIONS)
    dpi = get_dpi_for_window(root.winfo_id())
    den = ACUITY_LEVELS[current_level_index]
    draw_landolt_c(stimulus_canvas(), dpi, VIEWING_DISTANCE_M, den, current_true_orientation)
    draw_debug_overlay()

    # Arrow-key responses are accepted immediately once the Landolt-C is visible.
    _key_debounce = False


_acuity_show_after_id = None  # pending "show the Landolt-C" callback


def cancel_acuity_trial():
    global _acuity_show_after_id
    if _acuity_show_after_id is not None:
        try:
            root.after_cancel(_acuity_show_after_id)
        except Exception:
            pass
        _acuity_show_after_id = None


def new_trial_acuity():
    """Start an acuity trial: 0.5-s blank screen, then show Landolt-C and enable input."""
    global _key_debounce, _acuity_show_after_id

    if stage != STAGE_ACUITY:
        return

    activate_stimulus_window()

    # Block responses while the screen is blank so a held/repeated arrow key
    # from the previous trial cannot answer the new trial accidentally.
    _key_debounce = True
    stim = stimulus_canvas()  # the TV in two-screen mode (the laptop keeps its arrows)
    stim.delete("all")
    stim.configure(bg=BG_SOFT)

    cancel_acuity_trial()  # never let two trials be pending at once
    _acuity_show_after_id = root.after(ACUITY_PRE_STIMULUS_DELAY_MS, _show_acuity_stimulus)


def on_arrow(event):
    global current_attempt, current_level_index, _key_debounce, stage
    if stage != STAGE_ACUITY:
        return
    if _key_debounce:
        return

    key = event.keysym
    if key not in KEYSYM_TO_ORIENTATION:
        return

    _key_debounce = True

    user_orientation = KEYSYM_TO_ORIENTATION[key]
    flash_arrow(user_orientation)
    flash_laptop_arrow(user_orientation)

    den = ACUITY_LEVELS[current_level_index]
    correct = (user_orientation == current_true_orientation)

    if correct:
        if den == LAST_LEVEL:
            # achieved last, failed at none
            record["Visual acuity achieved"] = f"20/{LAST_LEVEL}"
            record["Visual acuity failed at"] = ""

            # IMPORTANT: leave the acuity stage immediately before showing the
            # contrast transition. This prevents a held/repeated arrow-key event
            # from being interpreted as another 20/20 acuity response while the
            # transition screen is visible.
            stage = STAGE_CONTRAST
            _key_debounce = True
            canvas.delete("all")
            set_arrow_panel_visible(False)
            end_acuity_and_move_on()
            return

        current_level_index += 1
        current_attempt = 1
        new_trial_acuity()
        return

    # wrong
    if current_attempt == 1:
        current_attempt = 2
        new_trial_acuity()
    else:
        # 2nd wrong => fail at current level (unified)
        handle_acuity_failure(esc=False)


# =========================================================
# ========================= CONTRAST: ACHIEVED/FAILED AT ===
# =========================================================
def _set_contrast_outcome_by_failure(failed_at_triplet_idx: int):
    """
    Policy:
      - failed_at always recorded (hex + logCS for that triplet)
      - achieved:
          if failed_at == 0 -> "Failed"
          else -> previous triplet value (hex/logCS)
    """
    failed_at_triplet_idx = int(max(0, min(failed_at_triplet_idx, len(CONTRAST_HEX_SEQUENCE)-1)))

    # failed at
    record["Contrast sensitivity failed at_hexcode"] = hex_for_triplet(failed_at_triplet_idx)
    record["Contrast sensitivity failed at_log(CS)"] = logcs_for_triplet(failed_at_triplet_idx)

    if failed_at_triplet_idx == 0:
        record["Contrast sensitivity achieved_hexcode"] = "Failed"
        record["Contrast sensitivity achieved_log(CS)"] = "Failed"
    else:
        achieved_idx = failed_at_triplet_idx - 1
        record["Contrast sensitivity achieved_hexcode"] = hex_for_triplet(achieved_idx)
        record["Contrast sensitivity achieved_log(CS)"] = logcs_for_triplet(achieved_idx)


def _set_contrast_outcome_by_completion():
    """
    If user completes (reaches end) without failing:
      - achieved = last triplet shown (or last index)
      - failed_at = ""
    """
    # current_triplet_idx refers to the next to be shown; last shown is current-1
    current_triplet_idx = triplet_index_from_letter_index(contrast_triplet_start_index)
    last_shown = max(0, min(current_triplet_idx, len(CONTRAST_HEX_SEQUENCE)-1))

    record["Contrast sensitivity achieved_hexcode"] = hex_for_triplet(last_shown)
    record["Contrast sensitivity achieved_log(CS)"] = logcs_for_triplet(last_shown)
    record["Contrast sensitivity failed at_hexcode"] = ""
    record["Contrast sensitivity failed at_log(CS)"] = ""


def finish_contrast_and_move_to_color(failure=False, esc=False):
    global _contrast_finished
    if _contrast_finished:
        return
    _contrast_finished = True
    _cancel_contrast_eval()

    # stop global typing capture
    disable_contrast_typing_bind_all()

    # compute outcome
    failed_at_triplet_idx = _contrast_current_triplet_idx()

    if failure:
        _set_contrast_outcome_by_failure(failed_at_triplet_idx)
    else:
        _set_contrast_outcome_by_completion()

    clear_tv_window()  # keep the TV blank until the color target covers it
    enter_color()


def handle_contrast_failure(esc=False):
    """
    Unified handler used by:
      - typing result <2 correct (failure=True)
      - ESC pressed during contrast (failure=True)
    """
    if stage != STAGE_CONTRAST:
        return
    finish_contrast_and_move_to_color(failure=True, esc=esc)


_contrast_eval_after_id = None  # pending "score the typed triplet" callback


def _cancel_contrast_eval():
    global _contrast_eval_after_id
    if _contrast_eval_after_id is not None:
        try:
            root.after_cancel(_contrast_eval_after_id)
        except Exception:
            pass
        _contrast_eval_after_id = None


# ----- Two-screen contrast: target letters on the TV, typed letters on the laptop -----
_tv_win = None
_tv_canvas = None
_contrast_start_token = 0  # stops a delayed start if the page was left while switching


def open_tv_window(tv):
    """Borderless window covering the TV, used to show the target letters."""
    global _tv_win, _tv_canvas
    if _tv_win is not None and _tv_win.winfo_exists():
        _tv_canvas.delete("all")
        return
    r = tv["rect"]
    win = tk.Toplevel(root)
    win.overrideredirect(True)  # Tk "-fullscreen" would land on the laptop, not the TV
    win.configure(bg=BG_SOFT)
    tv_w, tv_h = r.right - r.left, r.bottom - r.top
    win.geometry(f"{tv_w}x{tv_h}+{r.left}+{r.top}")
    # Give the canvas the TV's exact size up front: the letters are laid out from the
    # canvas size, and a brand-new window can still report 1 px before it is shown.
    c = tk.Canvas(win, bg=BG_SOFT, highlightthickness=0, width=tv_w, height=tv_h)
    c.pack(fill="both", expand=True)
    win.attributes("-topmost", True)  # stay above the TV's taskbar
    win.update_idletasks()
    _tv_win, _tv_canvas = win, c


def clear_tv_window():
    """Blank the TV (keeps the window so the desktop never shows through)."""
    if _tv_canvas is not None:
        try:
            _tv_canvas.delete("all")
        except Exception:
            pass


def close_tv_window():
    global _tv_win, _tv_canvas
    if _tv_win is not None:
        try:
            _tv_win.destroy()
        except Exception:
            pass
    _tv_win = None
    _tv_canvas = None


def stimulus_canvas():
    """Where the test stimulus is drawn: the TV in two-screen mode, else the main window."""
    return _tv_canvas if _tv_canvas is not None else canvas


# ----- Two-screen acuity: Landolt C on the TV, answer arrows on the laptop -----
_acuity_start_token = 0  # stops a delayed start if the page was left while switching
_LAPTOP_ARROW_GLYPHS = {"up": "↑", "down": "↓", "left": "←", "right": "→"}


def draw_laptop_arrows(highlight=None):
    """Four answer arrows on the laptop; `highlight` lights up the one just pressed."""
    canvas.delete("laptop_arrows")
    canvas.update_idletasks()
    W = max(10, canvas.winfo_width())
    H = max(10, canvas.winfo_height())
    b = int(min(W, H) * 0.20)  # arrow box size; the whole cross is 3.3 boxes tall/wide
    d = int(b * 1.15)
    cx, cy = W / 2.0, H / 2.0
    positions = {"up": (cx, cy - d), "down": (cx, cy + d), "left": (cx - d, cy), "right": (cx + d, cy)}
    for key, (x, y) in positions.items():
        canvas.create_rectangle(x - b / 2, y - b / 2, x + b / 2, y + b / 2,
                                fill=ARROW_HILITE if key == highlight else "white",
                                outline="black", width=4, tags=("laptop_arrows",))
        canvas.create_text(x, y, text=_LAPTOP_ARROW_GLYPHS[key], font=("Segoe UI", -int(b * 0.6), "bold"),
                           fill="black", tags=("laptop_arrows",))


def flash_laptop_arrow(direction, ms=220):
    """Briefly light up the pressed arrow on the laptop (two-screen acuity only)."""
    if _tv_canvas is None:
        return
    draw_laptop_arrows(highlight=direction)

    def reset():
        if stage == STAGE_ACUITY and _tv_canvas is not None:
            draw_laptop_arrows()
    root.after(ms, reset)


def start_acuity_trials():
    """With the TV connected: switch to Extend, C on the TV, arrows on the laptop."""
    global _acuity_start_token
    _acuity_start_token += 1
    token = _acuity_start_token

    if not ACUITY_TWO_SCREENS:
        close_tv_window()
        new_trial_acuity()
        return

    def begin(screens):
        if token != _acuity_start_token or stage != STAGE_ACUITY:
            return  # the page was left while Windows was switching
        if screens is not None:
            open_tv_window(screens[0])
            draw_laptop_arrows()
        else:
            close_tv_window()
        new_trial_acuity()

    ensure_two_screens(begin)


def draw_contrast_typed_only(canvas_, typed):
    """Laptop view during two-screen contrast: only the participant's typed letters."""
    canvas_.delete("contrast_input")
    canvas_.update_idletasks()
    W = max(10, canvas_.winfo_width())
    H = max(10, canvas_.winfo_height())
    slots = [typed[i] if i < len(typed) else "_" for i in range(3)]
    font_px = int(max(40, min(220, H * 0.28, W * 0.09)))
    canvas_.create_text(
        W / 2.0, H / 2.0,
        text="   ".join(slots),
        font=("Segoe UI", -font_px, "bold"),
        fill="#333333",
        anchor="center",
        tags=("contrast_input",),
    )


def _redraw_contrast_input():
    if _tv_canvas is not None:
        draw_contrast_typed_only(canvas, contrast_user_buffer)  # laptop
    else:
        draw_contrast_input(canvas, contrast_user_buffer)       # under the letters


def _draw_contrast_screens():
    """Target letters on the TV (or the main window on one screen), then the typed letters."""
    target = _tv_canvas if _tv_canvas is not None else canvas
    dpi = get_dpi_for_window(root.winfo_id())
    draw_contrast_triplet(target, dpi, CONTRAST_BOX_MM, contrast_true_triplet, contrast_triplet_start_index)
    _redraw_contrast_input()


def start_contrast_triplets():
    global stage, _contrast_finished, _contrast_start_token
    _cancel_contrast_eval()
    _contrast_finished = False
    stage = STAGE_CONTRAST
    _contrast_start_token += 1
    token = _contrast_start_token

    if not CONTRAST_TWO_SCREENS:
        _begin_contrast_triplets(None)
        return

    def begin(screens):
        if token != _contrast_start_token or stage != STAGE_CONTRAST or _contrast_finished:
            return  # the page was left while Windows was switching
        _begin_contrast_triplets(screens)

    ensure_two_screens(begin)


def _begin_contrast_triplets(screens):
    """Start the triplets. `screens` = (tv, laptop) for the two-screen layout, else None."""
    global contrast_triplet_start_index, contrast_true_triplet, contrast_user_buffer

    if screens is not None:
        open_tv_window(screens[0])
    else:
        close_tv_window()

    set_hud_instruction(
        "Type the 3 letters (A-Z). Enter = submit when fewer than 3 are visible. Backspace = clear. "
        "F2 = skip this triplet; ESC = confirm skip; Alt+← / Alt+→ = previous / next page"
    )

    # hide arrow panel so it can't steal focus
    set_arrow_panel_visible(False)

    # capture typing regardless of focus window
    enable_contrast_typing_bind_all()

    contrast_triplet_start_index = 0
    contrast_user_buffer = ""
    contrast_true_triplet = "".join(random.choice(SLOAN_LETTERS) for _ in range(3))

    canvas.delete("all")
    _draw_contrast_screens()
    activate_stimulus_window()
    draw_debug_overlay()


def _advance_to_next_triplet():
    global contrast_triplet_start_index, contrast_true_triplet, contrast_user_buffer

    _cancel_contrast_eval()

    # If the participant has just passed the lowest-contrast level we want to
    # test (triplet_idx = 7), finish here.  Do NOT increment first, otherwise
    # the completion record could incorrectly refer to triplet_idx = 8.
    current_triplet_idx = triplet_index_from_letter_index(contrast_triplet_start_index)
    if current_triplet_idx >= MAX_CONTRAST_TRIPLET_IDX:
        finish_contrast_and_move_to_color(failure=False)
        return

    contrast_triplet_start_index += 3

    # Safety end conditions (normally the MAX_CONTRAST_TRIPLET_IDX check above
    # ends the task first).
    if triplet_index_from_letter_index(contrast_triplet_start_index) >= len(CONTRAST_HEX_SEQUENCE):
        finish_contrast_and_move_to_color(failure=False)
        return
    if contrast_triplet_start_index + 2 >= MAX_LETTER_INDEX:
        finish_contrast_and_move_to_color(failure=False)
        return

    contrast_user_buffer = ""
    contrast_true_triplet = "".join(random.choice(SLOAN_LETTERS) for _ in range(3))

    activate_stimulus_window()
    _draw_contrast_screens()
    draw_debug_overlay()


def _score_contrast_answer():
    """Score the typed letters for this triplet and schedule the next step.

    Letters that were not typed (answer submitted early with Enter) count as wrong.
    """
    global _contrast_eval_after_id
    user3 = contrast_user_buffer
    truth3 = contrast_true_triplet
    n_correct = sum(1 for i in range(min(3, len(user3))) if user3[i] == truth3[i])

    # Rule: ≥2 correct -> advance; else fail
    if n_correct >= 2:
        _contrast_eval_after_id = root.after(350, _advance_to_next_triplet)
    else:
        _contrast_eval_after_id = root.after(180, lambda: handle_contrast_failure(esc=False))


def on_key_contrast_triplets(event):
    global contrast_user_buffer
    if stage != STAGE_CONTRAST:
        return
    if _contrast_finished:
        return

    if event.keysym == "Escape":
        # Esc inside the main window is handled by on_escape (root binding), which asks
        # for confirmation first. Only handle it here if it came from another window.
        try:
            from_root = event.widget.winfo_toplevel() is root
        except Exception:
            from_root = False
        if not from_root:
            on_escape()
        return

    # This triplet's answer is already being scored; ignore extra keys meanwhile.
    if _contrast_eval_after_id is not None:
        return

    if event.keysym in ("BackSpace", "Delete"):
        contrast_user_buffer = ""
        _redraw_contrast_input()
        draw_debug_overlay()
        return

    # Enter submits fewer than 3 letters (e.g. the participant can only read two of
    # the faint final letters); untyped positions count as wrong.
    if event.keysym in ("Return", "KP_Enter"):
        if 1 <= len(contrast_user_buffer) < 3:
            _score_contrast_answer()
        return

    ch = (event.char or "").strip().upper()
    if not (len(ch) == 1 and "A" <= ch <= "Z"):
        return
    if len(contrast_user_buffer) >= 3:
        return

    contrast_user_buffer += ch
    _redraw_contrast_input()
    draw_debug_overlay()

    if len(contrast_user_buffer) < 3:
        return

    _score_contrast_answer()


# =========================================================
# ========================= COLOR MATCHING STAGE ===========
# =========================================================
target_win = None
picker_win = None

_target_hex = None
_target_rgb = None

_sel_h = 0.0
_sel_s = 0.0
_sel_rgb = (255, 0, 0)
_sel_hex = "#FF0000"

_wheel_img = None
_wheel_photo = None
_wheel_canvas = None
_pick_chip = None
_wheel_marker = None


def _random_target_color():
    h = random.random()
    s = random.uniform(0.20, 1.0)
    v = 1.0
    rgb = hsv_to_rgb(h, s, v)
    return rgb_to_hex(rgb), rgb


_wheel_image_cache = {}  # size_px -> PIL image; building the wheel pixel by pixel is slow


def _make_wheel_image(size_px):
    if not PIL_OK:
        return None, None

    cached = _wheel_image_cache.get(size_px)
    if cached is not None:
        return cached, ImageTk.PhotoImage(cached)

    v = 1.0
    img = Image.new("RGBA", (size_px, size_px), (255, 255, 255, 0))
    cx = cy = (size_px - 1) / 2.0
    r = (size_px - 2) / 2.0

    pix = img.load()
    for y in range(size_px):
        dy = y - cy
        for x in range(size_px):
            dx = x - cx
            dist = math.sqrt(dx * dx + dy * dy)
            if dist <= r:
                s = dist / r
                ang = math.atan2(dy, dx)
                h = (ang / (2.0 * math.pi)) % 1.0
                rgb = hsv_to_rgb(h, s, v)
                pix[x, y] = (rgb[0], rgb[1], rgb[2], 255)
            else:
                pix[x, y] = (255, 255, 255, 0)

    _wheel_image_cache[size_px] = img
    return img, ImageTk.PhotoImage(img)


def _refresh_wheel():
    global _wheel_img, _wheel_photo, _wheel_marker
    if not PIL_OK or _wheel_canvas is None:
        return

    _wheel_img, _wheel_photo = _make_wheel_image(COLOR_WHEEL_SIZE_PX)
    _wheel_canvas.delete("all")
    _wheel_canvas.create_image(COLOR_WHEEL_SIZE_PX // 2, COLOR_WHEEL_SIZE_PX // 2, image=_wheel_photo)
    _wheel_canvas.create_oval(2, 2, COLOR_WHEEL_SIZE_PX - 2, COLOR_WHEEL_SIZE_PX - 2,
                             outline="gray30", width=2)
    _wheel_marker = None


def _draw_wheel_marker(x, y):
    """Mark the currently selected point on the wheel (the large cursor ring is separate)."""
    global _wheel_marker
    if _wheel_canvas is None:
        return

    _wheel_canvas.delete("wheel_marker")

    r = int(WHEEL_MARKER_RADIUS_PX)
    outer_w = int(WHEEL_MARKER_STROKE_PX) + 5
    inner_w = int(WHEEL_MARKER_STROKE_PX)

    # Double-contrast ring: black outside + white inside stays visible on
    # bright, dark, and saturated parts of the color wheel.
    _wheel_canvas.create_oval(
        x - r, y - r, x + r, y + r,
        outline="black",
        width=outer_w,
        tags=("wheel_marker",)
    )
    _wheel_marker = _wheel_canvas.create_oval(
        x - r, y - r, x + r, y + r,
        outline="white",
        width=inner_w,
        tags=("wheel_marker",)
    )

    # High-visibility center crosshair.
    cross = max(10, r // 2)
    _wheel_canvas.create_line(
        x - cross, y, x + cross, y,
        fill="black", width=7, tags=("wheel_marker",)
    )
    _wheel_canvas.create_line(
        x, y - cross, x, y + cross,
        fill="black", width=7, tags=("wheel_marker",)
    )
    _wheel_canvas.create_line(
        x - cross, y, x + cross, y,
        fill="white", width=3, tags=("wheel_marker",)
    )
    _wheel_canvas.create_line(
        x, y - cross, x, y + cross,
        fill="white", width=3, tags=("wheel_marker",)
    )

    # Small center dot marks the exact sampled pixel.
    dot = 4
    _wheel_canvas.create_oval(
        x - dot, y - dot, x + dot, y + dot,
        fill="black", outline="white", width=2,
        tags=("wheel_marker",)
    )


def _wheel_xy_from_hs(h, s):
    """Convert the current HSV hue/saturation to wheel canvas coordinates."""
    size = COLOR_WHEEL_SIZE_PX
    cx = cy = (size - 1) / 2.0
    r = (size - 2) / 2.0
    ang = (float(h) % 1.0) * 2.0 * math.pi
    dist = clamp01(s) * r
    return cx + math.cos(ang) * dist, cy + math.sin(ang) * dist


# ----- Large, high-visibility mouse cursor for color matching -----
# Windows shrinks custom cursor files down to its own pointer size, so the large
# cursor is drawn in a small borderless, see-through, click-through window that
# follows the mouse everywhere while color matching is on screen. The real pointer
# stays at its center and every click passes straight through it.
CURSOR_RING_SIZE_PX = 150           # outer diameter (~55 mm on the 32" TV)
CURSOR_RING_KEY_COLOR = "#FF00FE"   # see-through key color; never used in the drawing
CURSOR_RING_FOLLOW_MS = 15

_cursor_ring_win = None
_cursor_ring_after_id = None
_cursor_ring_last_xy = None


def _make_window_click_through(win):
    """Let mouse clicks pass through `win` and keep it from taking focus (Windows)."""
    if sys.platform != "win32":
        return
    user32 = ctypes.windll.user32
    hwnd = user32.GetParent(win.winfo_id())
    GWL_EXSTYLE = -20
    WS_EX_TRANSPARENT = 0x00000020
    WS_EX_TOOLWINDOW = 0x00000080
    WS_EX_LAYERED = 0x00080000
    WS_EX_NOACTIVATE = 0x08000000
    user32.GetWindowLongW.restype = ctypes.c_long
    ex = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    user32.SetWindowLongW(hwnd, GWL_EXSTYLE,
                          ex | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_LAYERED | WS_EX_NOACTIVATE)


def _follow_cursor_ring():
    global _cursor_ring_after_id, _cursor_ring_last_xy
    _cursor_ring_after_id = None
    win = _cursor_ring_win
    if win is None:
        return
    try:
        if not win.winfo_exists():
            return
        x, y = root.winfo_pointerxy()
        if _cursor_clamp_rect is not None:
            # Two-screen color matching: keep the mouse on the laptop (wheel) screen.
            left, top, right, bottom = _cursor_clamp_rect
            cx, cy = min(max(x, left), right - 1), min(max(y, top), bottom - 1)
            if (cx, cy) != (x, y):
                ctypes.windll.user32.SetCursorPos(cx, cy)
                x, y = cx, cy
        if (x, y) != _cursor_ring_last_xy:
            s = CURSOR_RING_SIZE_PX
            win.geometry(f"{s}x{s}+{x - s // 2}+{y - s // 2}")
            # Tk only applies a window move in its idle pass; apply it right away so
            # the ring keeps up with the mouse.
            win.update_idletasks()
            win.lift()
            _cursor_ring_last_xy = (x, y)
    except Exception:
        pass
    _cursor_ring_after_id = root.after(CURSOR_RING_FOLLOW_MS, _follow_cursor_ring)


def start_cursor_ring():
    """Show the large cursor ring. Returns False if it could not be created."""
    global _cursor_ring_win, _cursor_ring_last_xy
    stop_cursor_ring()
    try:
        s = CURSOR_RING_SIZE_PX
        win = tk.Toplevel(root)
        win._is_cursor_overlay = True
        win.overrideredirect(True)
        win.configure(bg=CURSOR_RING_KEY_COLOR)
        win.attributes("-topmost", True)
        win.attributes("-transparentcolor", CURSOR_RING_KEY_COLOR)
        c = tk.Canvas(win, width=s, height=s, bg=CURSOR_RING_KEY_COLOR, highlightthickness=0)
        c.pack()
        # Black / yellow / black ring stays visible on light, dark and saturated colors.
        m = 7
        c.create_oval(m, m, s - m, s - m, outline="black", width=14)
        c.create_oval(m, m, s - m, s - m, outline="#FFD400", width=7)
        # Small center cross marks the exact point a click will pick.
        mid, arm = s / 2.0, 14
        for width, color in ((6, "black"), (2, "white")):
            c.create_line(mid - arm, mid, mid + arm, mid, fill=color, width=width)
            c.create_line(mid, mid - arm, mid, mid + arm, fill=color, width=width)
        win.update_idletasks()
        _make_window_click_through(win)
        _cursor_ring_win = win
        _cursor_ring_last_xy = None
        _follow_cursor_ring()
        return True
    except Exception:
        stop_cursor_ring()
        return False


def stop_cursor_ring():
    global _cursor_ring_win, _cursor_ring_after_id
    if _cursor_ring_after_id is not None:
        try:
            root.after_cancel(_cursor_ring_after_id)
        except Exception:
            pass
        _cursor_ring_after_id = None
    if _cursor_ring_win is not None:
        try:
            _cursor_ring_win.destroy()
        except Exception:
            pass
        _cursor_ring_win = None


def _warp_pointer_to_current_wheel_selection():
    """Start color matching with the real mouse pointer centered in the ring."""
    if _wheel_canvas is None:
        return
    try:
        x, y = _wheel_xy_from_hs(_sel_h, _sel_s)
        x_i, y_i = int(round(x)), int(round(y))
        _draw_wheel_marker(x_i, y_i)
        _wheel_canvas.focus_force()
        # Tk's warp=True physically moves the Windows pointer to this canvas point.
        _wheel_canvas.event_generate("<Motion>", warp=True, x=x_i, y=y_i)
    except Exception:
        # The custom ring still works even if pointer warping is unavailable.
        pass


def _update_selected_chip():
    global _sel_hex, _sel_rgb
    _sel_rgb = hsv_to_rgb(_sel_h, _sel_s, 1.0)
    _sel_hex = rgb_to_hex(_sel_rgb)

    if _pick_chip is not None:
        _pick_chip.configure(bg=_sel_hex)

    record["color_selected_hex"] = _sel_hex
    record["color_selected_rgb"] = str(_sel_rgb)
    if _target_rgb is not None:
        record["color_error_rgb_dist"] = round(rgb_dist(_target_rgb, _sel_rgb), 3)


def _on_wheel_pick(event):
    global _sel_h, _sel_s
    if not PIL_OK:
        return

    x = event.x
    y = event.y
    size = COLOR_WHEEL_SIZE_PX
    cx = cy = (size - 1) / 2.0
    r = (size - 2) / 2.0

    dx = x - cx
    dy = y - cy
    dist = math.sqrt(dx * dx + dy * dy)
    if dist > r:
        return

    _sel_s = clamp01(dist / r)
    ang = math.atan2(dy, dx)
    _sel_h = (ang / (2.0 * math.pi)) % 1.0

    _draw_wheel_marker(x, y)
    _update_selected_chip()


def _place_color_windows():
    # Fill the whole screen: two half-width, full-height windows side by side
    # (readability mode; color matching is not physically calibrated).
    root.update_idletasks()

    if _color_screens is not None:
        tv, laptop = _color_screens
        # Target color fills the TV: no title bar or taskbar for the participant to see.
        # (Tk's "-fullscreen" always goes to the MAIN screen on Windows, so instead use a
        # borderless window sized exactly to the TV, kept above the TV's taskbar.)
        r = tv["rect"]
        target_win.overrideredirect(True)
        target_win.geometry(f"{r.right - r.left}x{r.bottom - r.top}+{r.left}+{r.top}")
        target_win.update_idletasks()
        target_win.after(300, lambda: target_win.winfo_exists() and target_win.attributes("-topmost", True))
        # The color wheel fills the laptop's usable area (taskbar excluded). Tk sizes only
        # the inside of a window, so measure the title bar / border and fit the whole window.
        wk = laptop["work"]
        lap_w, lap_h = wk.right - wk.left, wk.bottom - wk.top
        picker_win.resizable(False, False)
        picker_win.geometry(f"{lap_w}x{lap_h}+{wk.left}+{wk.top}")
        try:
            root.update_idletasks()
            deco_top = picker_win.winfo_rooty() - picker_win.winfo_y()
            deco_side = picker_win.winfo_rootx() - picker_win.winfo_x()
            if 0 < deco_top < 200 and 0 <= deco_side < 60:
                lap_w, lap_h = max(200, lap_w - 2 * deco_side), max(200, lap_h - deco_top - deco_side)
                picker_win.geometry(f"{lap_w}x{lap_h}+{wk.left}+{wk.top}")
        except Exception:
            pass
        picker_win.minsize(lap_w, lap_h)
        picker_win.maxsize(lap_w, lap_h)
        return

    mon = get_monitor_rect_for_window(root.winfo_id())
    mw = mon.right - mon.left
    mh = mon.bottom - mon.top

    win_w = mw // 2
    win_h = mh

    def apply(client_w, client_h):
        for w in (target_win, picker_win):
            w.resizable(False, False)
            w.minsize(client_w, client_h)
            w.maxsize(client_w, client_h)
        target_win.geometry(f"{client_w}x{client_h}+{mon.left}+{mon.top}")
        picker_win.geometry(f"{client_w}x{client_h}+{mon.left + win_w}+{mon.top}")

    apply(win_w, win_h)

    # These windows keep their title bars, and Tk sizes only the inside of a window.
    # Measure the title bar / border and shrink the inside so the WHOLE window fits
    # on the screen; otherwise the bottom ~45 px (and the buttons) run off the screen.
    try:
        root.update_idletasks()
        deco_top = target_win.winfo_rooty() - target_win.winfo_y()
        deco_side = target_win.winfo_rootx() - target_win.winfo_x()
        if 0 < deco_top < 200 and 0 <= deco_side < 60:
            apply(max(200, win_w - 2 * deco_side), max(200, win_h - deco_top - deco_side))
    except Exception:
        pass


def _finish_color_and_move_to_survey():
    global _color_finished
    if _color_finished:
        return
    _color_finished = True

    close_color_windows()
    back_to_one_screen(enter_survey)  # back to Duplicate first (if it was switched)


def close_color_windows():
    stop_cursor_ring()
    try:
        if target_win is not None:
            target_win.destroy()
    except Exception:
        pass
    try:
        if picker_win is not None:
            picker_win.destroy()
    except Exception:
        pass


# ----- Two screens (TV + laptop) for contrast and color matching -----
# Windows is switched to Extend when contrast starts and stays there through color
# matching, then goes back to Duplicate for the survey.
_screens_extended = False        # Windows was asked for Extend; must go back to Duplicate
_color_screens = None            # (tv, laptop) monitors while the two-screen layout is up
_color_build_token = 0           # stops a delayed window build if the page was left
_display_switch_hold = False     # pause the Duplicate hot-plug watcher while we switch
_cursor_clamp_rect = None        # keep the mouse on the laptop screen during color matching


def _show_display_wait(text):
    canvas.delete("display_wait")
    canvas.update_idletasks()
    W = max(10, canvas.winfo_width())
    H = max(10, canvas.winfo_height())
    canvas.create_rectangle(0, 0, W, H, fill=BG_SOFT, outline="", tags=("display_wait",))
    canvas.create_text(W / 2, H * 0.45, text=text, font=("Segoe UI", 26, "bold"),
                       fill="black", tags=("display_wait",))
    canvas.create_text(W / 2, H * 0.58, text="(Please wait…)", font=("Segoe UI", 12, "italic"),
                       fill="gray35", tags=("display_wait",))
    canvas.update_idletasks()


def _hide_display_wait():
    canvas.delete("display_wait")


def _wait_for_screens(condition, then):
    """Poll the screen count until `condition(n)` holds (or the timeout), then `then(ok)`."""
    deadline = time.monotonic() + DISPLAY_SWITCH_TIMEOUT_MS / 1000.0

    def poll():
        if condition(len(list_monitors())):
            root.after(DISPLAY_SWITCH_SETTLE_MS, lambda: then(True))
            return
        if time.monotonic() > deadline:
            then(False)
            return
        root.after(200, poll)

    root.after(200, poll)


def ensure_two_screens(then):
    """Switch Windows to Extend (unless it already is) and call `then(screens)`.

    `screens` is (tv, laptop), or None when only one screen is available, in which
    case Windows is left in (or put back in) Duplicate.
    """
    global _screens_extended, _display_switch_hold
    if sys.platform != "win32" or active_display_count() < 2:
        then(None)
        return
    if _screens_extended:
        screens = find_color_screens()
        if screens is not None:
            then(screens)
            return
    _display_switch_hold = True
    _screens_extended = True
    _show_display_wait("Setting up the TV and laptop screens")
    request_display_mode("extend")

    def after_extend(ok):
        _hide_display_wait()
        screens = find_color_screens() if ok else None
        if screens is None:
            # Extend did not take effect: go back to Duplicate and use one screen.
            back_to_one_screen(lambda: then(None))
            return
        then(screens)

    _wait_for_screens(lambda n: n >= 2, after_extend)


def back_to_one_screen(then):
    """Put Windows back in Duplicate (if it was switched to Extend), then `then()`."""
    global _screens_extended, _color_screens, _cursor_clamp_rect
    _color_screens = None
    _cursor_clamp_rect = None
    close_tv_window()
    if not _screens_extended:
        then()
        return
    _screens_extended = False
    _show_display_wait("Returning to the main screen")
    request_display_mode("clone")

    def after_clone(_ok):
        global _display_switch_hold
        _hide_display_wait()
        try:
            _restore_experiment_layout_after_display_change()
        except Exception:
            pass
        _display_switch_hold = False
        then()

    _wait_for_screens(lambda n: n <= 1, after_clone)


def start_color_match():
    """Color-matching page.

    With the TV connected, Windows is in Extend so the target color is shown only on
    the TV and the color wheel only on the laptop.
    """
    global _color_finished, _color_build_token

    _color_finished = False
    _color_build_token += 1
    token = _color_build_token

    set_hud("Stage: Color matching", progress_text="", instruction=COLOR_HUD)

    if not COLOR_MATCH_TWO_SCREENS:
        close_tv_window()
        _build_color_windows(None)  # one screen: both windows side by side
        return

    def build(screens):
        if token != _color_build_token or stage != STAGE_COLOR or _color_finished:
            return  # the page was left while Windows was switching
        _build_color_windows(screens)
        close_tv_window()  # the color target now covers the TV

    ensure_two_screens(build)


def _build_color_windows(screens):
    """Create the target and picker windows. `screens` = (tv, laptop) or None for one screen."""
    global target_win, picker_win
    global _target_hex, _target_rgb
    global _sel_h, _sel_s
    global _wheel_canvas, _pick_chip
    global COLOR_WHEEL_SIZE_PX
    global _color_screens, _cursor_clamp_rect

    _color_screens = screens

    _target_hex, _target_rgb = _random_target_color()
    record["color_target_hex"] = _target_hex
    record["color_target_rgb"] = str(_target_rgb)

    _sel_h = random.random()
    _sel_s = 0.8

    # Readability sizing (color matching is NOT physically calibrated):
    # size chips + wheel from the monitor so they fill the screen and are easy to see.
    if screens is not None:
        tv, laptop = screens
        tv_w = tv["rect"].right - tv["rect"].left
        tv_h = tv["rect"].bottom - tv["rect"].top
        # Same target size as on a shared screen (that formula used half the width).
        tgt_chip_px = int(max(200, min(tv_w * 0.40, tv_h * 0.70)))
        # The picker has the laptop to itself: size it from the laptop's usable area
        # (taskbar excluded, ~60 px left for the window's title bar).
        lap_w = laptop["work"].right - laptop["work"].left
        lap_h = laptop["work"].bottom - laptop["work"].top - 60
        pick_chip_px = int(max(200, min(lap_w * 0.52, lap_h * 0.32)))
        COLOR_WHEEL_SIZE_PX = int(max(220, min(lap_w * 0.55, lap_h * 0.40)))
        r = laptop["rect"]
        _cursor_clamp_rect = (r.left, r.top, r.right, r.bottom)
    else:
        _mon = get_monitor_rect_for_window(root.winfo_id())
        _mw = _mon.right - _mon.left
        _mh = _mon.bottom - _mon.top
        _half_w = max(200, _mw // 2)
        tgt_chip_px = int(max(200, min(_half_w * 0.80, _mh * 0.70)))
        pick_chip_px = int(max(200, min(_half_w * 0.52, _mh * 0.32)))
        COLOR_WHEEL_SIZE_PX = int(max(220, min(_half_w * 0.55, _mh * 0.40)))
        _cursor_clamp_rect = None

    target_win = tk.Toplevel(root)
    target_win.title("Target Color")
    target_win.configure(bg=BG_SOFT)
    target_win.protocol("WM_DELETE_WINDOW", lambda: save_and_exit())

    picker_win = tk.Toplevel(root)
    picker_win.title("Pick Matching Color")
    picker_win.configure(bg=BG_SOFT)
    picker_win.protocol("WM_DELETE_WINDOW", lambda: save_and_exit())

    for w in (target_win, picker_win):
        try:
            w.attributes("-topmost", True)
            w.after(250, lambda ww=w: ww.attributes("-topmost", False))
        except Exception:
            pass

    twrap = tk.Frame(target_win, bg=BG_SOFT, padx=16, pady=16)
    twrap.pack(fill="both", expand=True)

    tk.Label(twrap, text="Target color (fixed)", bg=BG_SOFT,
             font=("Segoe UI", 40, "bold")).pack(pady=(0, 6))

    chip = tk.Frame(
        twrap,
        bg=_target_hex,
        width=tgt_chip_px,
        height=tgt_chip_px,
        relief="solid",
        bd=2
    )
    chip.pack(pady=6)
    chip.pack_propagate(False)

    pwrap = tk.Frame(picker_win, bg=BG_SOFT, padx=12, pady=12)
    pwrap.pack(fill="both", expand=True)

    tk.Label(pwrap, text="Pick the closest color", bg=BG_SOFT,
             font=("Segoe UI", 40, "bold")).pack(pady=(0, 2))

    if PIL_OK:
        _wheel_canvas = tk.Canvas(
            pwrap,
            width=COLOR_WHEEL_SIZE_PX,
            height=COLOR_WHEEL_SIZE_PX,
            bg=BG_SOFT,
            highlightthickness=0,
            cursor="none"  # hide the small Windows cursor; the large ring is the cursor
        )
        _wheel_canvas.pack(pady=(0, 8))
        _wheel_canvas.bind("<Button-1>", _on_wheel_pick)
        _wheel_canvas.bind("<B1-Motion>", _on_wheel_pick)
        _refresh_wheel()
        tk.Label(pwrap, text="(Click/drag on wheel)", bg=BG_SOFT, fg="gray25",
                 font=("Segoe UI", 16)).pack(pady=(0, 8))
    else:
        tk.Label(
            pwrap,
            text="Pillow not installed.\nInstall with: pip install pillow\n(Falling back to button picker)",
            bg=BG_SOFT,
            fg="gray25",
            font=("Segoe UI", 11),
            justify="center"
        ).pack(pady=(10, 10))

    row = tk.Frame(pwrap, bg=BG_SOFT)
    row.pack(pady=(0, 8))

    _pick_chip = tk.Frame(
        row,
        bg="#FF0000",
        width=pick_chip_px,
        height=pick_chip_px,
        relief="solid",
        bd=2
    )
    _pick_chip.pack(side="left")
    _pick_chip.pack_propagate(False)

    btns = tk.Frame(pwrap, bg=BG_SOFT)
    btns.pack(pady=(16, 0))

    tk.Button(btns, text="Done (Enter)", font=("Segoe UI", 16, "bold"),
              command=_finish_color_and_move_to_survey, width=14).pack(side="left", padx=6)

    tk.Button(btns, text="Quit (ESC)", font=("Segoe UI", 16),
              command=on_escape, width=12).pack(side="left", padx=6)

    _update_selected_chip()
    _place_color_windows()

    # Large cursor ring that follows the mouse over both color windows. If it cannot
    # be shown, fall back to a visible crosshair pointer over the wheel.
    if not start_cursor_ring() and _wheel_canvas is not None:
        _wheel_canvas.configure(cursor="crosshair")

    # After window geometry is finalized, move the actual mouse pointer onto the
    # current selection so the participant starts with the cursor on the wheel.
    picker_win.after(250, _warp_pointer_to_current_wheel_selection)

    # Alt+Left / Alt+Right also work while either color window has focus.
    set_nav_page(lambda d: _nav_from_stage(STAGE_COLOR, d), root, target_win, picker_win)

    picker_win.bind("<Return>", lambda e: _finish_color_and_move_to_survey())
    target_win.bind("<Return>", lambda e: _finish_color_and_move_to_survey())
    picker_win.bind("<Escape>", lambda e: on_escape())
    target_win.bind("<Escape>", lambda e: on_escape())
    picker_win.bind(f"<{SKIP_KEYSYM}>", lambda e: skip_current_task())
    target_win.bind(f"<{SKIP_KEYSYM}>", lambda e: skip_current_task())


# =========================================================
# ========================= SURVEY STAGE (TEXT OUTPUT) =====
# =========================================================
_survey_dialog = None
_survey_skip_question = None  # set by start_survey; skips ONE question (F2)
_survey_skip_all = None  # set by start_survey; skips ALL remaining questions (Esc)

def start_survey(start_q=0):
    """Show the survey, starting at question `start_q` (Alt+Left can reopen the last one)."""
    global _survey_dialog, _survey_finished, stage
    if _survey_finished:
        return
    stage = STAGE_SURVEY

    set_hud("Stage: Survey", progress_text="", instruction=SURVEY_HUD)

    dialog = tk.Toplevel(root)
    _survey_dialog = dialog
    dialog.title("Survey")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.protocol("WM_DELETE_WINDOW", lambda: save_and_exit())

    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    dpi = get_layout_dpi(root.winfo_id())
    # Survey is for reading, not acuity: use the full screen.
    try:
        dialog.attributes("-fullscreen", True)
    except Exception:
        pass
    dialog.update_idletasks()
    win_w, win_h = dialog.winfo_width(), dialog.winfo_height()

    dialog.lift()
    dialog.attributes("-topmost", True)
    dialog.after(150, lambda: dialog.attributes("-topmost", False))

    GAP_TITLE_TO_CATEGORY_IN = 0.55
    GAP_PX = max(16, int(round(GAP_TITLE_TO_CATEGORY_IN * dpi)))

    # Large type: the participant reads this from ~10 ft (3 m) away.
    title_font = ("Segoe UI", 96)
    cat_font = ("Segoe UI", 84)
    end_font = ("Segoe UI", 54)
    small_font = ("Segoe UI", 26)

    wrap = tk.Frame(dialog, bg=BG_SOFT, padx=18, pady=14)
    wrap.pack(fill="both", expand=True)

    # On-screen exit for this fullscreen stage (same action as the window X).
    tk.Button(
        dialog, text="Exit ✕", command=lambda: save_and_exit(),
        bg="#C00000", fg="white", activebackground="#C00000",
        activeforeground="white", relief="solid", bd=2,
        font=("Segoe UI", 18, "bold"), padx=10, pady=4
    ).place(relx=1.0, x=-18, y=14, anchor="ne")

    tk.Label(
        wrap,
        text="How satisfied are you with...",
        bg=BG_SOFT,
        font=title_font,
        anchor="w",
        justify="left"
    ).pack(fill="x", pady=(0, 0))

    tk.Frame(wrap, bg=BG_SOFT, height=GAP_PX).pack(fill="x")

    cat_var = tk.StringVar(value="")
    cat_label = tk.Label(
        wrap,
        textvariable=cat_var,
        bg=BG_SOFT,
        font=cat_font,
        anchor="center",
        justify="center",
        fg="black"
    )
    cat_label.pack(fill="x", pady=(0, 6))

    canvas_h = int(round(5.3 * dpi))
    c = tk.Canvas(wrap, bg=BG_SOFT, highlightthickness=0, height=canvas_h)
    c.pack(fill="x", pady=(0, 2))

    progress_var_local = tk.StringVar(value="")
    progress_lbl = tk.Label(
        wrap,
        textvariable=progress_var_local,
        bg=BG_SOFT,
        fg="gray25",
        font=small_font,
        anchor="w"
    )
    progress_lbl.pack(fill="x", pady=(0, 4))

    questions = SURVEY_QUESTIONS

    MARGIN_INCH = 0.65
    MARGIN_PX = max(18, int(round(MARGIN_INCH * dpi)))
    # Keep the 7 dots in a centered band (not full-width) so the endpoint
    # labels have room to be large.
    SCALE_MAX_SPAN_PX = 1500

    R = int(round(0.95 * dpi))
    R = max(70, min(R, 120))
    DOT = max(8, int(round(R * 0.48)))
    HIT = R + 14
    GAP_TEXT_TO_CIRCLE = int(round(0.10 * dpi))

    state = {
        "q_idx": 0,
        "value": 4,
        "centers": [],
        "circle_ids": [],
        "dot_ids": [],
        "blink_after_ids": [],
    }

    def _compute_centers():
        c.update_idletasks()
        W = max(320, c.winfo_width())
        H = max(120, c.winfo_height())
        # Push the dots (and the endpoint labels above them) lower so the
        # "Very dissatisfied / satisfied" text is not clipped at the top.
        y_circle = int(round(H * 0.70))
        max_span = min(W - 2 * MARGIN_PX, SCALE_MAX_SPAN_PX)
        left = (W - max_span) / 2.0
        step = max_span / 6.0
        return [(left + i * step, y_circle) for i in range(7)]

    def _update_visuals():
        thick = max(3, int(round(R * 0.16)))
        thin = max(2, int(round(R * 0.12)))
        for i in range(1, 8):
            cid = state["circle_ids"][i - 1]
            did = state["dot_ids"][i - 1]
            if i == state["value"]:
                c.itemconfigure(cid, outline="black", width=thick)
                c.itemconfigure(did, state="normal")
            else:
                c.itemconfigure(cid, outline="gray35", width=thin)
                c.itemconfigure(did, state="hidden")

    def _draw_endpoints_on_canvas():
        if not state["centers"]:
            return
        x1, y1 = state["centers"][0]
        x7, y7 = state["centers"][6]
        y_text = (y1 - R - GAP_TEXT_TO_CIRCLE)

        # Keep the endpoint labels fully on-screen: the words "dissatisfied" /
        # "satisfied" are centered on the end circles and can otherwise run off
        # the screen edge (clipping the leading "d" so it reads "lissatisfied").
        import tkinter.font as tkfont
        try:
            _ef = tkfont.Font(family=end_font[0], size=end_font[1])
            half_left = _ef.measure("dissatisfied") / 2.0
            half_right = _ef.measure("satisfied") / 2.0
        except Exception:
            half_left = half_right = 0.0
        Wc = max(320, c.winfo_width())
        x1 = max(MARGIN_PX + half_left, x1)
        x7 = min(Wc - MARGIN_PX - half_right, x7)

        c.create_text(
            x1, y_text,
            text="Very\ndissatisfied",
            font=end_font,
            fill="gray25",
            justify="center",
            anchor="s",
            tags=("endpoints",)
        )
        c.create_text(
            x7, y_text,
            text="Very\nsatisfied",
            font=end_font,
            fill="gray25",
            justify="center",
            anchor="s",
            tags=("endpoints",)
        )

    def _draw_scale():
        c.delete("all")
        state["centers"] = _compute_centers()
        state["circle_ids"] = []
        state["dot_ids"] = []

        _draw_endpoints_on_canvas()

        for i, (x, y) in enumerate(state["centers"], start=1):
            cid = c.create_oval(
                x - R, y - R, x + R, y + R,
                outline="gray35",
                width=max(2, int(round(R * 0.12))),
                fill="white",
                tags=("circle", f"v{i}")
            )
            state["circle_ids"].append(cid)

            did = c.create_oval(
                x - DOT, y - DOT, x + DOT, y + DOT,
                outline="",
                fill="black",
                state=("normal" if i == state["value"] else "hidden"),
                tags=("dot", f"v{i}")
            )
            state["dot_ids"].append(did)

        _update_visuals()

    def _clear_blink_jobs():
        for aid in state["blink_after_ids"]:
            try:
                dialog.after_cancel(aid)
            except Exception:
                pass
        state["blink_after_ids"].clear()

    def _blink_question():
        _clear_blink_jobs()
        base_fg = "black"
        flash_fg = "#C00000"
        base_bg = BG_SOFT
        flash_bg = "#FFF2F2"

        cat_label.configure(fg=base_fg)
        c.configure(bg=base_bg)

        steps = [(flash_fg, flash_bg), (base_fg, base_bg), (flash_fg, flash_bg), (base_fg, base_bg)]
        delays = [0, 110, 220, 330]

        def apply_step(k):
            fg, bg = steps[k]
            cat_label.configure(fg=fg)
            c.configure(bg=bg)

        for k, dly in enumerate(delays):
            state["blink_after_ids"].append(dialog.after(dly, lambda kk=k: apply_step(kk)))

    def _set_question(idx, do_blink=True):
        state["q_idx"] = idx
        label, _key = questions[idx]
        cat_var.set(label)
        progress_var_local.set(f"{idx + 1} / {len(questions)}")

        state["value"] = random.randint(1, 7)
        _draw_scale()

        if do_blink:
            _blink_question()

        dialog.focus_force()
        c.focus_set()

    def _close_dialog():
        global _survey_skip_question, _survey_skip_all
        _survey_skip_question = None
        _survey_skip_all = None
        clear_nav_page()
        _clear_blink_jobs()
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()

    def _finish_submit():
        global _survey_finished
        _survey_finished = True
        _close_dialog()
        finish_current_condition_and_continue()

    def _survey_nav(direction):
        # Each question is a page. Answers are only written when a question is
        # confirmed, so leaving one with Alt+arrows never changes its saved answer.
        q = state["q_idx"]
        _label, key = questions[q]
        if direction == NAV_BACK:
            if q > 0:
                _set_question(q - 1, do_blink=True)
            else:
                # Before the first question comes the color-matching page.
                _close_dialog()
                root.after(50, lambda: remind_laptop_on_table(enter_color))
            return
        if not str(record.get(key, "")).strip():
            record[key] = "SKIPPED"  # moving past a question that was never answered
        if q >= len(questions) - 1:
            _finish_submit()
        else:
            _set_question(q + 1, do_blink=True)

    def _commit_current_and_advance():
        _label, key = questions[state["q_idx"]]
        # ✅ save TEXT (not numbers)
        record[key] = SURVEY_VALUE_TO_TEXT.get(int(state["value"]), str(state["value"]))

        if state["q_idx"] >= len(questions) - 1:
            _finish_submit()
        else:
            _set_question(state["q_idx"] + 1, do_blink=True)

    def _skip_current_question():
        # F2: skip ONLY this question (record it "SKIPPED"), then advance.
        _label, key = questions[state["q_idx"]]
        record[key] = "SKIPPED"
        if state["q_idx"] >= len(questions) - 1:
            _finish_submit()
        else:
            _set_question(state["q_idx"] + 1, do_blink=True)

    def _skip_whole_survey():
        # Esc: skip ALL remaining questions (each marked "SKIPPED"), then finish
        # the survey and continue to the next condition. Never quits the session.
        for i in range(state["q_idx"], len(questions)):
            _label, key = questions[i]
            record[key] = "SKIPPED"
        _finish_submit()

    global _survey_skip_question, _survey_skip_all
    _survey_skip_question = _skip_current_question
    _survey_skip_all = _skip_whole_survey

    def _on_click(event):
        if not state["centers"]:
            return
        x, y = event.x, event.y
        best_i, best_d = None, 1e18
        for i, (cx, cy_) in enumerate(state["centers"], start=1):
            d = (x - cx) ** 2 + (y - cy_) ** 2
            if d < best_d:
                best_d = d
                best_i = i
        if best_i is None:
            return
        cx, cy_ = state["centers"][best_i - 1]
        if (x - cx) ** 2 + (y - cy_) ** 2 <= HIT ** 2:
            state["value"] = best_i
            _update_visuals()

    def _on_key(event):
        ks = event.keysym
        if ks in ("Left", "KP_Left"):
            state["value"] = max(1, state["value"] - 1)
            _update_visuals()
            return "break"
        if ks in ("Right", "KP_Right"):
            state["value"] = min(7, state["value"] + 1)
            _update_visuals()
            return "break"
        if ks in ("Return", "KP_Enter"):
            _commit_current_and_advance()
            return "break"
        if ks == "Escape":
            on_escape()
            return "break"
        return None

    c.bind("<Button-1>", _on_click)
    c.bind("<B1-Motion>", _on_click)

    dialog.bind("<Left>", _on_key)
    dialog.bind("<Right>", _on_key)
    dialog.bind("<Return>", _on_key)
    dialog.bind("<Escape>", _on_key)
    dialog.bind(f"<{SKIP_KEYSYM}>", lambda e: skip_current_task())
    bind_nav_keys(dialog)  # Alt+Left / Alt+Right must not also move the selection
    set_nav_page(_survey_nav, dialog)

    # No on-screen hotkey buttons here: the participant views this from ~10 ft
    # and the researcher drives it with the keyboard (arrows / Enter / Esc / F2; Tab moves focus).

    _set_question(max(0, min(start_q, len(questions) - 1)), do_blink=True)
    root.wait_window(dialog)


# =========================================================
# ============ TEST PAGES (enter / leave with Alt+arrows) ==
# =========================================================
# Result fields owned by each test page, restored if the page is left unfinished.
STAGE_RESULT_FIELDS = {
    STAGE_ACUITY: ["Visual acuity achieved", "Visual acuity failed at"],
    STAGE_CONTRAST: ["Contrast sensitivity achieved_log(CS)", "Contrast sensitivity failed at_log(CS)",
                     "Contrast sensitivity achieved_hexcode", "Contrast sensitivity failed at_hexcode"],
    STAGE_COLOR: ["color_target_hex", "color_target_rgb", "color_selected_hex",
                  "color_selected_rgb", "color_error_rgb_dist"],
}
_stage_snapshot = {}


def set_condition_hud(instruction):
    idx = record_seq_idx if record_seq_idx is not None else current_sequence_index
    set_hud(f"Condition: {sequence_items[idx]}",
            progress_text=f"Step {idx + 1} / {len(sequence_items)}",
            instruction=instruction)


def _begin_stage_page(stage_name):
    """Remember this test's current results and make Alt+arrows act on it."""
    global _stage_snapshot
    _stage_snapshot = {f: record.get(f, "") for f in STAGE_RESULT_FIELDS[stage_name]}
    set_nav_page(lambda d: _nav_from_stage(stage_name, d), root)


def enter_acuity():
    global stage, current_level_index, current_attempt, current_true_orientation, _key_debounce
    cancel_transition()
    cancel_acuity_trial()
    stage = STAGE_ACUITY
    current_level_index = 0
    current_attempt = 1
    current_true_orientation = None
    _key_debounce = True  # no answers until the first Landolt-C is on screen
    canvas.delete("all")
    _begin_stage_page(STAGE_ACUITY)

    # Keep the separate arrow-key guidance window hidden.
    # Participants still answer the Landolt-C with the keyboard arrow keys.
    set_arrow_panel_visible(False)
    set_condition_hud(ACUITY_HUD)
    show_transition("Stage: Acuity", f"Condition: {sequence_items[record_seq_idx]}", ms=650,
                    after_fn=start_acuity_trials)
    activate_stimulus_window()


def enter_contrast():
    global stage
    cancel_transition()
    stage = STAGE_CONTRAST
    canvas.delete("all")
    clear_tv_window()
    _begin_stage_page(STAGE_CONTRAST)
    set_arrow_panel_visible(False)
    show_transition("Next: Contrast letters", "Type 3 letters. Backspace clears. F2 = skip. Tab = move focus.",
                    ms=700, after_fn=start_contrast_triplets)


def enter_color():
    global stage
    cancel_transition()
    stage = STAGE_COLOR
    canvas.delete("all")
    _begin_stage_page(STAGE_COLOR)
    show_transition("Next: Color matching", "Click/drag the wheel. Enter when done.",
                    ms=700, after_fn=start_color_match)


# Researcher reminders: the laptop is hidden during the survey and on the table otherwise.
LAPTOP_HIDE_MESSAGE = "Hide the laptop screen now.\n\nThe participant should not see the laptop during the survey."
LAPTOP_TABLE_MESSAGE = "Place the laptop back on the table."


def remind_laptop_on_table(then):
    """After the survey: remind the researcher to put the laptop back on the table."""
    show_instruction_popup(LAPTOP_TABLE_MESSAGE, can_go_back=False, show_glass=False)
    then()


def enter_survey(start_q=0):
    global stage, _survey_finished
    cancel_transition()
    stage = STAGE_SURVEY
    _survey_finished = False
    canvas.delete("all")
    clear_nav_page()  # the survey window registers its own pages once it is open

    def after_hide_reminder():
        # Before the survey: remind the researcher to hide the laptop screen.
        if show_instruction_popup(LAPTOP_HIDE_MESSAGE, show_glass=False) == NAV_BACK:
            enter_color()  # Alt+Left: back to color matching
            return
        show_transition("Next: Survey", "Use ←/→ then Enter to confirm.", ms=650,
                        after_fn=lambda: start_survey(start_q))

    root.after(50, after_hide_reminder)


def _leave_stage_unfinished(stage_name, direction):
    """Tear down a test page that is left with Alt+arrows before it finished."""
    global _key_debounce, _contrast_finished, _color_finished
    cancel_transition()
    clear_nav_page()
    if stage_name == STAGE_ACUITY:
        cancel_acuity_trial()
        _key_debounce = True
        clear_tv_window()
    elif stage_name == STAGE_CONTRAST:
        _contrast_finished = True
        _cancel_contrast_eval()
        disable_contrast_typing_bind_all()
        clear_tv_window()
    elif stage_name == STAGE_COLOR:
        _color_finished = True
        close_color_windows()
    canvas.delete("all")

    fields = STAGE_RESULT_FIELDS[stage_name]
    had_result = any(str(_stage_snapshot.get(f, "")).strip() for f in fields)
    if direction == NAV_BACK or had_result:
        # An unfinished redo keeps whatever this test had before.
        for f in fields:
            record[f] = _stage_snapshot.get(f, "")
    elif stage_name == STAGE_COLOR:
        # Same as F2: keep the target that was shown, mark the choice SKIPPED.
        for f in ("color_selected_hex", "color_selected_rgb", "color_error_rgb_dist"):
            record[f] = "SKIPPED"
    else:
        for f in fields:
            record[f] = "SKIPPED"


def _nav_from_stage(stage_name, direction):
    if stage != stage_name:
        return
    _leave_stage_unfinished(stage_name, direction)
    # Acuity, contrast and color matching use the TV + laptop (Extend). Moving between them
    # keeps Extend; leaving them (back to the illuminance page, on to the survey) returns
    # to Duplicate first.
    leaves_two_screens = ((stage_name == STAGE_ACUITY and direction == NAV_BACK) or
                          (stage_name == STAGE_COLOR and direction == NAV_NEXT))
    if leaves_two_screens:
        back_to_one_screen(lambda: _route_after_stage(stage_name, direction))
        return
    _route_after_stage(stage_name, direction)


def _route_after_stage(stage_name, direction):
    if direction == NAV_BACK:
        if stage_name == STAGE_ACUITY:
            idx = record_seq_idx
            root.after(50, lambda: go_to_sequence_item(idx))  # this condition's illuminance page
        elif stage_name == STAGE_CONTRAST:
            enter_acuity()
        else:
            enter_contrast()
    else:
        if stage_name == STAGE_ACUITY:
            enter_contrast()
        elif stage_name == STAGE_CONTRAST:
            enter_color()
        else:
            enter_survey()


# =========================================================
# ========================= SKIP (ACCESSIBILITY) ==========
# =========================================================
def skip_current_task(event=None):
    """
    Skip only the CURRENT ITEM and move on to the next item in the same task:
      - Acuity   : show a fresh Landolt-C at the current level
      - Contrast : advance to the next letter set (triplet)
      - Survey   : advance to the next question (this one recorded "SKIPPED")
      - Color    : this task has a single item, so it finishes and moves on

    This is deliberately NOT a whole-task skip. To skip an entire task, use Esc.
    Bound to SKIP_KEYSYM (F2). Returns "break" so the key never doubles as a
    response, and never triggers Tkinter's default focus traversal mid-test.
    """
    global stage, current_attempt, _contrast_finished

    if stage == STAGE_ACUITY:
        # Skip just this one presentation: re-roll a fresh trial at the same level.
        current_attempt = 1
        new_trial_acuity()
        return "break"

    if stage == STAGE_CONTRAST:
        # Skip just this one letter set: go to the next triplet
        # (_advance_to_next_triplet ends the task on its own once triplets run out).
        if _contrast_finished or _contrast_bind_all is None:
            return "break"  # finished, or the first triplet is not on screen yet
        _advance_to_next_triplet()
        return "break"

    if stage == STAGE_COLOR:
        # Color matching is a single item, so skipping it finishes the task.
        if _color_finished:
            return "break"  # already finished (e.g. screens are switching back)
        record["color_selected_hex"] = "SKIPPED"
        record["color_selected_rgb"] = "SKIPPED"
        record["color_error_rgb_dist"] = "SKIPPED"
        _finish_color_and_move_to_survey()
        return "break"

    if stage == STAGE_SURVEY:
        # Skip just this one question (marked "SKIPPED"); finishes after the last.
        if _survey_skip_question is not None:
            _survey_skip_question()
        return "break"

    return "break"


# =========================================================
# ========================= BINDINGS =======================
# =========================================================
root.bind("<Up>", on_arrow)
root.bind("<Down>", on_arrow)
root.bind("<Left>", on_arrow)
root.bind("<Right>", on_arrow)
root.bind("<Escape>", on_escape)

# Exit the whole run ONLY via a window's close (X) button (or Alt+F4). Esc never quits.
root.protocol("WM_DELETE_WINDOW", lambda: save_and_exit())

# SKIP: same key in every stage. Bound on root (covers acuity + contrast, whose
# focus lives on the main canvas) and on the arrow panel. The color windows and
# the survey dialog are separate Toplevels, so they bind SKIP_KEYSYM themselves.
root.bind(f"<{SKIP_KEYSYM}>", skip_current_task)
arrow_win.bind(f"<{SKIP_KEYSYM}>", skip_current_task)
# NOTE: contrast typing is handled by bind_all during contrast stage; the root-level
# SKIP binding fires (and "break"s) before that bind_all sees the key.


# =========================================================
# ========================= MONITOR TICK ===================
# =========================================================
_last_dpi = None
_display_signature = get_active_display_signature()
_display_change_in_progress = False
_display_change_candidate = None
_display_change_candidate_since = None
_last_monitor_rect = get_monitor_rect_for_window(root.winfo_id())
_display_change_old_monitor_rect = None
DISPLAY_CHANGE_STABLE_MS = 750
DISPLAY_REAPPLY_SETTLE_MS = 1800


def _force_refresh_after_display_change():
    """Force Tk/Windows to repaint and reactivate the experiment after a monitor swap.

    This intentionally mimics the useful effect of manually Alt+Tabbing back to the
    experiment, without synthesizing keyboard input.  It refreshes geometry, raises
    visible windows, generates expose/configure events, briefly toggles top-most, and
    restores focus to the front-most visible popup (or the root if no popup is open).
    """
    windows = [root]
    try:
        windows.extend(_iter_toplevels(root))
    except Exception:
        pass

    visible = []
    for win in windows:
        try:
            if not win.winfo_exists() or not win.winfo_viewable():
                continue
            visible.append(win)

            # Ask Tk to recompute layout and Windows to repaint this window.
            win.update_idletasks()
            try:
                win.event_generate("<Configure>", when="tail")
                win.event_generate("<Expose>", when="tail")
            except Exception:
                pass

            try:
                win.lift()
            except Exception:
                pass

            # A brief top-most toggle forces the window manager to reactivate the
            # surface after DisplaySwitch without permanently pinning the UI.
            try:
                win.attributes("-topmost", True)
            except Exception:
                pass
        except Exception:
            pass

    def _finish_refresh():
        # Release temporary top-most state.
        for win in visible:
            try:
                if win.winfo_exists():
                    win.attributes("-topmost", False)
                    win.update_idletasks()
            except Exception:
                pass

        # Prefer the most recently created visible popup, because that is normally
        # the page the researcher/participant is currently interacting with.
        focus_target = None
        for win in reversed(visible):
            try:
                if isinstance(win, tk.Toplevel) and win.winfo_exists() and win.winfo_viewable():
                    focus_target = win
                    break
            except Exception:
                pass
        if focus_target is None:
            focus_target = root

        try:
            focus_target.lift()
            focus_target.focus_force()
        except Exception:
            pass

        # If a child control already owned focus, keep keyboard interaction natural.
        try:
            child = focus_target.focus_get()
            if child is not None:
                child.focus_set()
        except Exception:
            pass

    # Two short passes are deliberate: the first happens after Tk's geometry pass;
    # the second occurs after Windows has repainted the newly cloned display.
    root.after(120, _finish_refresh)
    root.after(360, _finish_refresh)


def _restore_experiment_layout_after_display_change():
    """Reassert the experiment layout after Windows finishes changing topology."""
    global _display_signature, _display_change_in_progress
    global _last_monitor_rect, _display_change_old_monitor_rect

    try:
        root.update_idletasks()
        root.minsize(1, 1)
        root.maxsize(100000, 100000)
        root.attributes("-fullscreen", True)
        root.configure(bg=BG_SOFT)
        root.update_idletasks()
    except Exception:
        pass

    # Reflow every currently-open popup to the newly active monitor geometry.
    # This preserves the same relative size/position and rescales text/buttons if
    # the new monitor causes Windows to choose a different effective resolution.
    try:
        new_mon = get_monitor_rect_for_window(root.winfo_id())
        _reflow_open_toplevels(_display_change_old_monitor_rect, new_mon)
        _last_monitor_rect = new_mon
    except Exception:
        pass
    _display_change_old_monitor_rect = None

    # Re-read the active displays *after* clone mode has settled so the monitor
    # watcher does not immediately retrigger on the topology change we caused.
    _display_signature = get_active_display_signature()
    _display_change_in_progress = False

    # Keep the currently running experiment visible/focused, then force a
    # repaint/reactivation pass. This removes the need to manually Alt+Tab once
    # after a monitor hot-plug.
    try:
        activate_stimulus_window()
    except Exception:
        pass
    try:
        _force_refresh_after_display_change()
    except Exception:
        pass


def _reapply_duplicate_after_monitor_change():
    """Hot-plug handler: restore Windows Duplicate mode, then restore Tk layout."""
    global _display_change_in_progress, _display_change_old_monitor_rect

    if _display_change_in_progress:
        return

    # Preserve the pre-change monitor rectangle so open popups can be mapped to
    # the same relative position/size on the new monitor after clone mode settles.
    _display_change_old_monitor_rect = _last_monitor_rect
    _display_change_in_progress = True
    request_duplicate_display_nonblocking()

    # DisplaySwitch can momentarily resize/reposition Tk windows.  Give Windows
    # time to settle, then force the experiment back to its fullscreen layout.
    root.after(DISPLAY_REAPPLY_SETTLE_MS, _restore_experiment_layout_after_display_change)


def monitor_tick():
    global _last_dpi, _display_signature, _last_monitor_rect
    global _display_change_candidate, _display_change_candidate_since

    root.update_idletasks()
    dpi = get_dpi_for_window(root.winfo_id())

    # Detect monitor/HDMI hot-plug.  Require the new device set to remain stable
    # briefly because Windows may report several intermediate states while a
    # monitor is being connected or replaced.
    # (Paused while color matching has Windows in Extend mode on purpose.)
    if sys.platform == "win32" and not _display_change_in_progress and not _display_switch_hold:
        sig = get_active_display_signature()

        if sig != _display_signature:
            now_ms = int(time.monotonic() * 1000)
            if sig != _display_change_candidate:
                _display_change_candidate = sig
                _display_change_candidate_since = now_ms
            elif (_display_change_candidate_since is not None and
                  now_ms - _display_change_candidate_since >= DISPLAY_CHANGE_STABLE_MS):
                _display_change_candidate = None
                _display_change_candidate_since = None
                _reapply_duplicate_after_monitor_change()
        else:
            _display_change_candidate = None
            _display_change_candidate_since = None

    # Keep the grey field fullscreen.  A display hot-plug can make Windows drop
    # fullscreen temporarily, so this also preserves the experiment layout.
    try:
        if not bool(root.attributes("-fullscreen")):
            root.attributes("-fullscreen", True)
    except Exception:
        pass

    if stage == STAGE_ACUITY:
        position_arrow_window_attached()

    # Keep the last stable monitor geometry for the next hot-plug event.
    if not _display_change_in_progress:
        try:
            _last_monitor_rect = get_monitor_rect_for_window(root.winfo_id())
        except Exception:
            pass

    _last_dpi = dpi
    root.after(POLL_MS, monitor_tick)


# =========================================================
# ========================= START ==========================
# =========================================================
set_arrow_panel_visible(False)
stage = STAGE_ACUITY

set_hud(
    "Ready",
    progress_text="",
    instruction="The experiment will follow every test condition in Sheet 3: Full sequence."
)

monitor_tick()
activate_stimulus_window()
root.after(300, start_next_sequence_item)
root.mainloop()

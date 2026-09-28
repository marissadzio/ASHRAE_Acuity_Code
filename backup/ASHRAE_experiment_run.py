import tkinter as tk
from tkinter import messagebox
import sys
import math
import ctypes
from ctypes import wintypes
from pathlib import Path
from datetime import datetime
import random

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


# =========================================================
# ========================= USER SETTINGS =================
# =========================================================

# ✅ Toggle for debugging annotations
TEST = True  # True => show debug overlay (acuity 20/den, contrast hex/logCS, etc.)

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
#   Tab = skip the CURRENT task and move on to the next task.
#     - Works in every stage: acuity -> contrast -> color -> survey -> next condition.
#     - Records the skipped task as "SKIPPED" (NOT the same as a wrong answer / Esc fail).
#     - Chosen because Tab does not overlap any answer / submit / quit key listed above.
SKIP_KEYSYM = "Tab"  # change here to use a different skip key (must not overlap the keys above)

VIEWING_DISTANCE_M = 3.048  # 10 ft back (was 1.0). Use 3.0 if you mark the distance in metric.

# ===== Monitor physical-size calibration (Dell S2409W: 24", 1920x1080) =====
# Windows reports ~96 DPI at 100% scaling, but this panel is physically 92 PPI.
# Render all real-world sizes (Landolt-C, contrast boxes, color chips, layout)
# using the TRUE physical PPI so on-screen targets match their intended mm size.
USE_PHYSICAL_PPI = True
MONITOR_PHYSICAL_PPI = 92.0  # Dell S2409W native (24" 1920x1080). Verify: 20/320 should read ~71 mm.
POLL_MS = 250

# ===== Contrast letter shape controls (PIL render) =====
# NOTE: use 1.0 for normal; >1.0 stretches. (Your snippet had 10/10 which will explode.)
LETTER_WIDTH_SCALE = 1.00
LETTER_HEIGHT_SCALE = 1.00
LETTER_WEIGHT_PX = 0  # 0=off, 1..3 => thicker strokes

# ✅ Acuity levels (Landolt-C)
ACUITY_LEVELS = [320, 160, 80, 40, 20]
LAST_LEVEL = ACUITY_LEVELS[-1]

# ✅ Contrast follow-up (Pelli-style triplets)
# Contrast letters: hit a target on-screen LETTER height (ruler-measured).
# The Sloan glyph renders ~0.83 of its box, so size the box from that fill ratio.
# (20/680 at 3.048 m is ~150 mm overall; measured glyph 125 mm in a 150.7 mm box -> 0.829.)
CONTRAST_TARGET_LETTER_MM = 148.0
CONTRAST_LETTER_FILL = 0.829
CONTRAST_BOX_MM = CONTRAST_TARGET_LETTER_MM / CONTRAST_LETTER_FILL
SHOW_CONTRAST_BOX = False
SLOAN_LETTERS = list("CDHKNORSVZ")

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
WHEEL_MARKER_RADIUS_PX = 16
WHEEL_MARKER_STROKE_PX = 8

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

DOMAIN = "@berkeley.edu"
DOWNLOADS_DIR = Path.home() / "Downloads"
# XLSX_PATH is set per-participant at startup (participant_<id>_results.xlsx).
SHEET = "participants"

# ===== Experiment design file =====
# The design/sequence spreadsheet must live in the SAME folder as this program.
# It ships with the project download, right next to this .py and the .bat files,
# so this resolves correctly on any machine (no hard-coded C:\Users\... path).
DESIGN_XLSX_NAME = "final_material_sequence_AorC.xlsx"
DESIGN_XLSX_PATH = Path(__file__).resolve().with_name(DESIGN_XLSX_NAME)
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


def get_dpi_for_window(hwnd):
    # If the panel's true physical PPI differs from what Windows reports
    # (e.g. the 92 PPI Dell S2409W), override so real-world sizes are correct.
    if USE_PHYSICAL_PPI:
        return float(MONITOR_PHYSICAL_PPI)
    if sys.platform != "win32":
        return 96.0
    user32 = ctypes.windll.user32
    try:
        user32.GetDpiForWindow.restype = ctypes.c_uint
        dpi = user32.GetDpiForWindow(hwnd)
        return float(dpi) if dpi else 96.0
    except Exception:
        return 96.0


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
    """Set a popup dialog to `frac` of the current monitor and center it."""
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


def append_participant_row(path, row_dict):
    wb = load_workbook(path)
    ws = wb[SHEET]
    row = [row_dict.get(h, "") for h in HEADERS]
    ws.append(row)
    wb.save(path)


# =========================================================
# ========================= MAIN APP BOOT =================
# =========================================================
enable_per_monitor_dpi_awareness()

root = tk.Tk()
root.title('Landolt-C + Contrast Triplets + Color Match + Survey (3.048 m / 10 ft)')
root.resizable(False, False)
root.configure(bg=BG_SOFT)

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

    # Always confirm before skipping, every time Esc (or the Skip button) fires.
    if not messagebox.askyesno(
        "Skip step?",
        "Are you sure you want to skip this step?",
        default="no",
        parent=root,
    ):
        return  # user chose No — stay on the current step

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
        canvas.delete("transition")
        if callable(after_fn):
            after_fn()

    _transition_id = root.after(ms, done)


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


def prompt_participant_sequence_setup(root_):
    """Ask participant ID and email once at the start."""
    dialog = tk.Toplevel(root_)
    dialog.title("Participant Setup")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root_)

    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    dpi = get_dpi_for_window(root_.winfo_id())
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

    dialog.bind("<Return>", lambda e: submit())
    dialog.bind("<Escape>", lambda e: cancel())
    dialog.after(120, lambda: pid_entry.focus_force())
    root_.wait_window(dialog)
    return result["data"]


def prompt_illuminance_for_condition(root_, condition_text):
    """Ask H/V illuminance before each condition. Email is not asked again."""
    dialog = tk.Toplevel(root_)
    dialog.title("Next Condition")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root_)

    try:
        dialog.grab_set()
    except tk.TclError:
        pass

    dpi = get_dpi_for_window(root_.winfo_id())
    win_w, win_h = set_dialog_half_screen(dialog, root_, frac=1.0)

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

    result = {"data": None}

    def parse_float(s):
        s = (s or "").strip()
        if s == "":
            return None
        return float(s)

    def submit():
        try:
            h_val = parse_float(h_entry.get())
            v_val = parse_float(v_entry.get())
        except ValueError:
            messagebox.showerror("Invalid", "Illuminance must be numeric.", parent=dialog)
            return
        if h_val is None or v_val is None:
            messagebox.showerror("Missing", "Enter both H and V illuminance.", parent=dialog)
            return
        result["data"] = {"H": h_val, "V": v_val}
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
    btns.grid(row=4, column=1, sticky="w", pady=(30, 0))
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

    dialog.protocol("WM_DELETE_WINDOW", cancel)
    dialog.bind("<Return>", lambda e: submit())
    dialog.bind("<Escape>", on_escape)  # show a message instead of quitting
    dialog.after(120, lambda: h_entry.focus_force())
    root_.wait_window(dialog)
    return result["data"]


def show_instruction_popup(item):
    _update_context_from_instruction(item)

    dialog = tk.Toplevel(root)
    dialog.title("Instruction")
    dialog.configure(bg=BG_SOFT)
    dialog.resizable(False, False)
    dialog.transient(root)

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
    if current_glass and not is_baseline and not is_visit_cell:
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

    dialog.bind("<Return>", lambda e: close())
    dialog.bind("<Escape>", lambda e: close())

    root.wait_window(dialog)
    
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


def start_next_sequence_item():
    global current_sequence_index, record, current_glass, completed_test_count
    global H_ill, V_ill

    current_sequence_index += 1

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

    if is_instruction_item(item):
        mark_sequence_current(current_sequence_index)
        show_instruction_popup(item)
        mark_sequence_done(current_sequence_index)
        root.after(100, start_next_sequence_item)
        return

    mark_sequence_current(current_sequence_index)

    if item == "glass only" and not current_glass:
        current_glass = "glass only"

    inputs = prompt_illuminance_for_condition(root, item)
    if not inputs:
        root.destroy()
        return

    H_ill = inputs["H"]
    V_ill = inputs["V"]

    reset_condition_state()

    # Count only actual test conditions.
    # Instructions such as "Visit Cell 1", "Install glass C",
    # and "END of GLASS" are excluded from this order.
    completed_test_count += 1

    record = {h: "" for h in HEADERS}
    record["participant_id"] = participant_id
    record["sequence_step_order"] = completed_test_count
    record["condition"] = item
    record["raw_full_sequence_position"] = current_sequence_index + 1
    record["full_sequence"] = full_sequence_text
    record["participant_email"] = participant_email
    record["H_illuminance_lux"] = H_ill
    record["V_illuminance_lux"] = V_ill
    record["start_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    set_arrow_panel_visible(True)
    set_hud(
        f"Condition: {item}",
        progress_text=f"Step {current_sequence_index + 1} / {len(sequence_items)}",
        instruction="Press the arrow key for the gap direction. (Tab = skip this task; ESC quits using same rule as 2 wrong)"
    )
    show_transition("Stage: Acuity", f"Condition: {item}", ms=650, after_fn=new_trial_acuity)
    activate_stimulus_window()


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

    set_dialog_half_screen(dialog, root_)
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

    wrap = tk.Frame(dialog, bg=BG_SOFT, padx=40, pady=28)
    wrap.pack(fill="both", expand=True)

    tk.Label(wrap, text="Visual Experiment", bg=BG_SOFT,
             font=("Segoe UI", 38, "bold")).pack(anchor="w")
    tk.Label(wrap, text="ASHRAE 1925-TRP   ·   Acuity → Contrast → Color → Survey",
             bg=BG_SOFT, fg="gray25", font=("Segoe UI", 17)).pack(anchor="w", pady=(0, 22))

    btns = tk.Frame(wrap, bg=BG_SOFT)
    btns.pack(anchor="w", pady=(0, 22))
    tk.Button(btns, text="Practice  (Demo)", font=("Segoe UI", 19, "bold"),
              width=22, height=2, command=lambda: choose("demo")).pack(side="left", padx=(0, 16))
    tk.Button(btns, text="Start Real Session", font=("Segoe UI", 19, "bold"),
              width=22, height=2, command=lambda: choose("real")).pack(side="left")

    cheat = tk.Frame(wrap, bg="white", bd=2, relief="solid", padx=20, pady=14)
    cheat.pack(anchor="w", fill="x")
    tk.Label(cheat, text="Hotkeys (the same in every stage)", bg="white",
             font=("Segoe UI", 15, "bold")).pack(anchor="w", pady=(0, 8))
    for k, v in [
        ("Acuity", "Arrow keys  ← ↑ → ↓  =  direction of the gap"),
        ("Contrast", "Type the 3 letters (A-Z)    ·    Backspace = clear"),
        ("Color", "Click / drag the wheel    ·    Enter = confirm"),
        ("Survey", "← / →  = move selection    ·    Enter = confirm"),
        ("SKIP", "Tab  =  skip the current task (for those who cannot see)"),
        ("Esc", "Esc  =  skip the whole current task and move on (never quits)"),
        ("Exit", "Close (X) button on the window  =  end the session"),
    ]:
        line = tk.Frame(cheat, bg="white")
        line.pack(anchor="w", fill="x", pady=1)
        tk.Label(line, text=k, bg="white", fg="#C00000", width=12, anchor="w",
                 font=("Segoe UI", 13, "bold")).pack(side="left")
        tk.Label(line, text=v, bg="white", anchor="w",
                 font=("Segoe UI", 13)).pack(side="left")

    tk.Button(wrap, text="Quit", font=("Segoe UI", 14), width=10,
              command=lambda: choose(None)).pack(anchor="w", pady=(20, 0))

    dialog.protocol("WM_DELETE_WINDOW", lambda: choose(None))
    dialog.bind("<Escape>", lambda e: choose(None))
    root_.wait_window(dialog)
    return choice["val"]


_mode = prompt_launcher(root)
if _mode is None:
    root.destroy()
    raise SystemExit

if _mode == "demo":
    setup = _demo_setup()
else:
    setup = prompt_participant_sequence_setup(root)

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
    append_participant_row(XLSX_PATH, record)


def finish_current_condition_and_continue():
    # Called after Survey is complete. Save this condition, gray it out,
    # then automatically continue to the next item in Full sequence.
    finalize_and_save()
    mark_sequence_done(current_sequence_index)
    set_arrow_panel_visible(False)
    root.after(250, start_next_sequence_item)


def save_and_exit():
    # Used only for early quit / ESC. Save the active condition if it started, then exit.
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
    def go():
        start_contrast_triplets()
    show_transition("Next: Contrast letters", "Type 3 letters. Backspace clears. Tab = skip.", ms=700, after_fn=go)


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

def new_trial_acuity():
    global current_true_orientation
    if stage != STAGE_ACUITY:
        return

    activate_stimulus_window()
    current_true_orientation = random.choice(ORIENTATIONS)

    dpi = get_dpi_for_window(root.winfo_id())
    den = ACUITY_LEVELS[current_level_index]
    draw_landolt_c(canvas, dpi, VIEWING_DISTANCE_M, den, current_true_orientation)
    draw_debug_overlay()


def on_arrow(event):
    global current_attempt, current_level_index, _key_debounce
    if stage != STAGE_ACUITY:
        return
    if _key_debounce:
        return

    key = event.keysym
    if key not in KEYSYM_TO_ORIENTATION:
        return

    _key_debounce = True
    root.after(150, lambda: globals().__setitem__("_key_debounce", False))

    user_orientation = KEYSYM_TO_ORIENTATION[key]
    flash_arrow(user_orientation)

    den = ACUITY_LEVELS[current_level_index]
    correct = (user_orientation == current_true_orientation)

    if correct:
        if den == LAST_LEVEL:
            # achieved last, failed at none
            record["Visual acuity achieved"] = f"20/{LAST_LEVEL}"
            record["Visual acuity failed at"] = ""
            set_arrow_panel_visible(False)
            end_acuity_and_move_on()
            return

        current_level_index += 1
        current_attempt = 1
        root.after(250, new_trial_acuity)
        draw_debug_overlay()
        return

    # wrong
    if current_attempt == 1:
        current_attempt = 2
        root.after(250, new_trial_acuity)
        draw_debug_overlay()
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
    global _contrast_finished, stage
    if _contrast_finished:
        return
    _contrast_finished = True

    # stop global typing capture
    disable_contrast_typing_bind_all()

    # compute outcome
    failed_at_triplet_idx = _contrast_current_triplet_idx()

    if failure:
        _set_contrast_outcome_by_failure(failed_at_triplet_idx)
    else:
        _set_contrast_outcome_by_completion()

    canvas.delete("all")
    stage = STAGE_COLOR

    def go():
        start_color_match()
    show_transition("Next: Color matching", "Click/drag the wheel. Enter when done.", ms=700, after_fn=go)


def handle_contrast_failure(esc=False):
    """
    Unified handler used by:
      - typing result <2 correct (failure=True)
      - ESC pressed during contrast (failure=True)
    """
    if stage != STAGE_CONTRAST:
        return
    finish_contrast_and_move_to_color(failure=True, esc=esc)


def start_contrast_triplets():
    global stage, contrast_triplet_start_index, contrast_true_triplet, contrast_user_buffer, _contrast_finished
    _contrast_finished = False
    stage = STAGE_CONTRAST

    # hide arrow panel so it can't steal focus
    set_arrow_panel_visible(False)

    # capture typing regardless of focus window
    enable_contrast_typing_bind_all()

    contrast_triplet_start_index = 0
    contrast_user_buffer = ""
    contrast_true_triplet = "".join(random.choice(SLOAN_LETTERS) for _ in range(3))

    dpi = get_dpi_for_window(root.winfo_id())
    draw_contrast_triplet(canvas, dpi, CONTRAST_BOX_MM, contrast_true_triplet, contrast_triplet_start_index)
    activate_stimulus_window()
    draw_debug_overlay()


def _advance_to_next_triplet():
    global contrast_triplet_start_index, contrast_true_triplet, contrast_user_buffer

    contrast_triplet_start_index += 3

    # end conditions
    if triplet_index_from_letter_index(contrast_triplet_start_index) >= len(CONTRAST_HEX_SEQUENCE):
        finish_contrast_and_move_to_color(failure=False)
        return
    if contrast_triplet_start_index + 2 >= MAX_LETTER_INDEX:
        finish_contrast_and_move_to_color(failure=False)
        return

    contrast_user_buffer = ""
    contrast_true_triplet = "".join(random.choice(SLOAN_LETTERS) for _ in range(3))

    activate_stimulus_window()
    dpi = get_dpi_for_window(root.winfo_id())
    draw_contrast_triplet(canvas, dpi, CONTRAST_BOX_MM, contrast_true_triplet, contrast_triplet_start_index)
    draw_debug_overlay()


def on_key_contrast_triplets(event):
    global contrast_user_buffer
    if stage != STAGE_CONTRAST:
        return
    if _contrast_finished:
        return

    # allow ESC here too (in case bind_all catches it)
    if event.keysym == "Escape":
        handle_contrast_failure(esc=True)
        return

    if event.keysym in ("BackSpace", "Delete"):
        contrast_user_buffer = ""
        draw_debug_overlay()
        return

    ch = (event.char or "").strip().upper()
    if not (len(ch) == 1 and "A" <= ch <= "Z"):
        return
    if len(contrast_user_buffer) >= 3:
        return

    contrast_user_buffer += ch
    draw_debug_overlay()

    if len(contrast_user_buffer) < 3:
        return

    user3 = contrast_user_buffer
    truth3 = contrast_true_triplet
    n_correct = sum(1 for i in range(3) if user3[i] == truth3[i])

    # Rule: ≥2 correct -> advance; else fail
    if n_correct >= 2:
        root.after(350, _advance_to_next_triplet)
    else:
        root.after(180, lambda: handle_contrast_failure(esc=False))


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


def _make_wheel_image(size_px):
    if not PIL_OK:
        return None, None

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
    global _wheel_marker
    if _wheel_canvas is None:
        return

    _wheel_canvas.delete("wheel_marker")

    r = int(WHEEL_MARKER_RADIUS_PX)
    w_outer = int(WHEEL_MARKER_STROKE_PX) + 2
    w_inner = int(WHEEL_MARKER_STROKE_PX)

    _wheel_canvas.create_oval(
        x - r, y - r, x + r, y + r,
        outline="white",
        width=w_outer,
        tags=("wheel_marker",)
    )
    _wheel_marker = _wheel_canvas.create_oval(
        x - r, y - r, x + r, y + r,
        outline="black",
        width=w_inner,
        tags=("wheel_marker",)
    )


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
    mon = get_monitor_rect_for_window(root.winfo_id())
    mw = mon.right - mon.left
    mh = mon.bottom - mon.top

    win_w = mw // 2
    win_h = mh

    for w in (target_win, picker_win):
        w.resizable(False, False)
        w.minsize(win_w, win_h)
        w.maxsize(win_w, win_h)

    target_win.geometry(f"{win_w}x{win_h}+{mon.left}+{mon.top}")
    picker_win.geometry(f"{win_w}x{win_h}+{mon.left + win_w}+{mon.top}")


def _finish_color_and_move_to_survey():
    global _color_finished, stage
    if _color_finished:
        return
    _color_finished = True

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

    stage = STAGE_SURVEY

    def go():
        start_survey()
    show_transition("Next: Survey", "Use ←/→ then Enter to confirm.", ms=650, after_fn=go)


def start_color_match():
    global target_win, picker_win
    global _target_hex, _target_rgb
    global _sel_h, _sel_s
    global _wheel_canvas, _pick_chip
    global _color_finished
    global COLOR_WHEEL_SIZE_PX

    _color_finished = False

    set_hud(
        "Stage: Color matching",
        progress_text="",
        instruction="Click/drag on the wheel to pick the closest color. Press Enter when done. (Tab = skip this task)"
    )

    _target_hex, _target_rgb = _random_target_color()
    record["color_target_hex"] = _target_hex
    record["color_target_rgb"] = str(_target_rgb)

    _sel_h = random.random()
    _sel_s = 0.8

    # Readability sizing (color matching is NOT physically calibrated):
    # size chips + wheel from the monitor so they fill the screen and are easy to see.
    _mon = get_monitor_rect_for_window(root.winfo_id())
    _mw = _mon.right - _mon.left
    _mh = _mon.bottom - _mon.top
    _half_w = max(200, _mw // 2)
    tgt_chip_px = int(max(200, min(_half_w * 0.80, _mh * 0.70)))
    pick_chip_px = int(max(200, min(_half_w * 0.52, _mh * 0.32)))
    COLOR_WHEEL_SIZE_PX = int(max(220, min(_half_w * 0.55, _mh * 0.40)))

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
            highlightthickness=0
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
_survey_skip_question = None  # set by start_survey; skips ONE question (Tab)
_survey_skip_all = None  # set by start_survey; skips ALL remaining questions (Esc)

def start_survey():
    global _survey_dialog, _survey_finished, stage
    if _survey_finished:
        return
    stage = STAGE_SURVEY

    set_hud(
        "Stage: Survey",
        progress_text="",
        instruction="Use ← / → to move selection. Press Enter to confirm. (Tab = skip the survey)"
    )

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

    dpi = get_dpi_for_window(root.winfo_id())
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

    questions = [
        ("Clarity of view", "survey_clarity_of_view_7pt"),
        ("Visual privacy", "survey_visual_privacy_7pt"),
        ("Reflections / mirror-effect", "survey_reflections_mirror_effect_7pt"),
        ("Visual comfort (glare)", "survey_visual_comfort_glare_7pt"),
    ]

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

    def _finish_submit():
        global _survey_finished
        _survey_finished = True
        try:
            dialog.grab_release()
        except Exception:
            pass
        dialog.destroy()
        finish_current_condition_and_continue()

    def _commit_current_and_advance():
        _label, key = questions[state["q_idx"]]
        # ✅ save TEXT (not numbers)
        record[key] = SURVEY_VALUE_TO_TEXT.get(int(state["value"]), str(state["value"]))

        if state["q_idx"] >= len(questions) - 1:
            _finish_submit()
        else:
            _set_question(state["q_idx"] + 1, do_blink=True)

    def _skip_current_question():
        # Tab: skip ONLY this question (record it "SKIPPED"), then advance.
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

    # No on-screen hotkey buttons here: the participant views this from ~10 ft
    # and the researcher drives it with the keyboard (arrows / Enter / Esc / Tab).

    _set_question(0, do_blink=True)
    root.wait_window(dialog)


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
    Bound to SKIP_KEYSYM (Tab). Returns "break" so the key never doubles as a
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
        if _contrast_finished:
            return "break"
        _advance_to_next_triplet()
        return "break"

    if stage == STAGE_COLOR:
        # Color matching is a single item, so skipping it finishes the task.
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

def monitor_tick():
    global _last_dpi
    root.update_idletasks()
    dpi = get_dpi_for_window(root.winfo_id())

    # Keep the grey field fullscreen; do NOT re-lock to a fixed 8in window.
    try:
        if not bool(root.attributes("-fullscreen")):
            root.attributes("-fullscreen", True)
    except Exception:
        pass

    if stage == STAGE_ACUITY:
        position_arrow_window_attached()

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

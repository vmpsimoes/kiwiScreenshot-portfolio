"""
Kiwi Tool
---------
Press HOTKEY to capture the foreground window and embed it into an Excel file.
Each capture is timestamped and placed below the previous one.

Requirements:  pip install -r requirements.txt
Run:           run.bat  (as Administrator)
"""

import io
import math
import os
import sys
import tempfile
import threading
from datetime import datetime

import tkinter as tk

import win32gui
import keyboard
import mss
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageTk

# ── User configuration ────────────────────────────────────────────────────────

HOTKEY        = "scroll lock"
EXCEL_FILE    = "kiwi_screenshots.xlsx"
IMAGES_FOLDER = "screenshots"
CELL_SPACING  = 2

# ── Internal constants ────────────────────────────────────────────────────────

_START_COL       = "B"
_LABEL_COL       = 1
_EXCEL_PT_PER_PX = 0.75
_PX_PER_ROW      = 15.0 / _EXCEL_PT_PER_PX
_ICON_PATH       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kiwi_icon.png")
_ICO_PATH        = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kiwi_icon.ico")

# ── Global state ──────────────────────────────────────────────────────────────

_lock           = threading.Lock()
_current_row    = 1
_screenshot_num = 0
_session_sheet  = None   # name of the sheet created for this run


# ─────────────────────────────────────────────────────────────────────────────
# Kiwi icon generator
# ─────────────────────────────────────────────────────────────────────────────

def _make_kiwi_icon(size=64):
    """Draw a kiwi fruit cross-section and save as PNG + ICO."""
    img = _draw_kiwi(size)
    img.save(_ICON_PATH, format="PNG")

    # Save ICO with multiple sizes for Windows
    sizes  = [16, 32, 48, 64]
    frames = [_draw_kiwi(s) for s in sizes]
    frames[0].save(_ICO_PATH, format="ICO", append_images=frames[1:],
                   sizes=[(s, s) for s in sizes])
    return img


def _draw_kiwi(size):
    """Helper that draws a kiwi at the given size (used for ICO generation)."""
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    cx = cy = size // 2
    r  = size // 2 - 1
    draw.ellipse([cx-r,  cy-r,  cx+r,  cy+r],  fill="#7B5230")
    r2 = r - int(r * 0.12)
    draw.ellipse([cx-r2, cy-r2, cx+r2, cy+r2], fill="#8DB640")
    r3 = int(r2 * 0.72)
    draw.ellipse([cx-r3, cy-r3, cx+r3, cy+r3], fill="#A8CC5A")
    r4 = int(r2 * 0.22)
    draw.ellipse([cx-r4, cy-r4, cx+r4, cy+r4], fill="#F5F0DC")
    num_seeds  = 10
    seed_dist  = r2 * 0.52
    sw, sh_seed = max(2, size // 20), max(3, size // 13)
    for i in range(num_seeds):
        angle = (2 * math.pi * i / num_seeds) - math.pi / 2
        sx = cx + seed_dist * math.cos(angle)
        sy = cy + seed_dist * math.sin(angle)
        draw.ellipse([sx-sw//2, sy-sh_seed//2, sx+sw//2, sy+sh_seed//2], fill="#3D2B1F")
    return img


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _px_to_pt(px):
    return px * _EXCEL_PT_PER_PX

def _rows_for_height(h):
    return max(1, round(h / _PX_PER_ROW))

def _init_excel_if_needed():
    global _session_sheet

    # Build a unique sheet name for this session (by minute)
    base_name = datetime.now().strftime("%Y-%m-%d %H.%M")
    _session_sheet = base_name

    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = _session_sheet
        print(f"  Created: {os.path.abspath(EXCEL_FILE)}")
    else:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        # Avoid duplicate sheet name if opened twice in the same minute
        existing = wb.sheetnames
        suffix = 2
        while _session_sheet in existing:
            _session_sheet = f"{base_name} ({suffix})"
            suffix += 1
        ws = wb.create_sheet(title=_session_sheet)
        print(f"  Appended new session sheet to: {os.path.abspath(EXCEL_FILE)}")

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions[_START_COL].width = 170
    wb.save(EXCEL_FILE)
    print(f"  Session sheet : {_session_sheet}")


# ─────────────────────────────────────────────────────────────────────────────
# Screenshot capture
# ─────────────────────────────────────────────────────────────────────────────

def _capture() -> Image.Image:
    hwnd = win32gui.GetForegroundWindow()
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width  = right - left
    height = bottom - top

    if width <= 0 or height <= 0:
        raise ValueError("Foreground window has zero size.")

    with mss.mss() as sct:
        region = {"top": top, "left": left, "width": width, "height": height}
        raw = sct.grab(region)
        img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")

    return img


# ─────────────────────────────────────────────────────────────────────────────
# Note dialog
# ─────────────────────────────────────────────────────────────────────────────

def _ask_note():
    result = [None]

    def on_confirm(event=None):
        result[0] = entry.get().strip()
        root.destroy()

    def on_skip(event=None):
        result[0] = None
        root.destroy()

    root = tk.Tk()
    root.title("Kiwi Tool — Note")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    # Set kiwi icon on the dialog window
    try:
        icon_img = ImageTk.PhotoImage(Image.open(_ICON_PATH))
        root.iconphoto(True, icon_img)
        root._icon_ref = icon_img   # prevent garbage collection
    except Exception:
        pass

    w, h = 520, 140
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    tk.Label(root, text="Add a note for this screenshot:",
             font=("Segoe UI", 11), anchor="w").pack(fill="x", padx=16, pady=(14, 4))

    entry = tk.Entry(root, font=("Segoe UI", 12), width=52)
    entry.pack(padx=16, pady=(0, 10), ipady=6)

    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=(0, 10))
    tk.Button(btn_frame, text="Save  (Enter)", font=("Segoe UI", 10),
              width=16, command=on_confirm).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Skip  (Esc)", font=("Segoe UI", 10),
              width=16, command=on_skip).pack(side="left", padx=6)

    root.bind("<Return>", on_confirm)
    root.bind("<Escape>", on_skip)

    root.update()
    root.lift()
    root.grab_set()
    root.after(100, lambda: (root.focus_force(), entry.focus_force(), entry.icursor("end")))

    root.mainloop()
    return result[0]


# ─────────────────────────────────────────────────────────────────────────────
# Excel insertion
# ─────────────────────────────────────────────────────────────────────────────

def _insert(img: Image.Image, note: str = ""):
    global _current_row, _screenshot_num

    with _lock:
        _screenshot_num += 1
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 1. Save PNG to screenshots folder
        os.makedirs(IMAGES_FOLDER, exist_ok=True)
        img_filename = f"kiwi_{_screenshot_num:04d}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        img_path = os.path.join(IMAGES_FOLDER, img_filename)
        img.save(img_path, format="PNG")

        # 2. Temp copy for Excel embedding
        tmp_png = os.path.join(tempfile.gettempdir(), f"kiwi_tmp_{_screenshot_num}.png")
        img.save(tmp_png, format="PNG")

        # 3. Load workbook fresh
        wb = openpyxl.load_workbook(EXCEL_FILE)
        ws = wb[_session_sheet]

        # 4. Label (2 blank rows before each entry)
        _current_row += 2
        label = f"#{_screenshot_num}  |  {timestamp}"
        if note:
            label += f"  —  {note}"
        cell = ws.cell(row=_current_row, column=_LABEL_COL, value=label)
        cell.font = openpyxl.styles.Font(bold=True, size=11)
        ws.row_dimensions[_current_row].height = 18

        # 5. Row heights for image
        actual_w, actual_h = img.size
        rows = _rows_for_height(actual_h)
        img_row = _current_row + 1
        for r in range(img_row, img_row + rows):
            ws.row_dimensions[r].height = _px_to_pt(actual_h / rows)

        # 6. Embed image
        xl_img = XLImage(tmp_png)
        xl_img.width  = actual_w
        xl_img.height = actual_h
        xl_img.anchor = f"{_START_COL}{img_row}"
        ws.add_image(xl_img)

        # 7. Advance cursor
        _current_row = img_row + rows + CELL_SPACING

        # 8. Atomic save
        tmp_xlsx = EXCEL_FILE + ".tmp"
        wb.save(tmp_xlsx)
        os.replace(tmp_xlsx, EXCEL_FILE)

        # 9. Clean up temp PNG
        try:
            os.remove(tmp_png)
        except OSError:
            pass

    print(f"  [{timestamp}] Screenshot #{_screenshot_num} saved  "
          f"(next row: {_current_row})")


# ─────────────────────────────────────────────────────────────────────────────
# Hotkey wiring
# ─────────────────────────────────────────────────────────────────────────────

def _on_hotkey():
    threading.Thread(target=_safe_run, daemon=True).start()

def _safe_run():
    try:
        img  = _capture()
        note = _ask_note()
        if note is None:
            print("  Screenshot discarded.")
            return
        _insert(img, note)
    except Exception as exc:
        print(f"  ERROR: {exc}")


# ─────────────────────────────────────────────────────────────────────────────
# Project picker dialog
# ─────────────────────────────────────────────────────────────────────────────

def _ask_project():
    """Startup dialog: pick an existing Excel file or create a new one."""
    import glob as _glob

    result = [None]

    script_dir = os.path.dirname(os.path.abspath(__file__))
    existing = sorted(
        os.path.basename(f)
        for f in _glob.glob(os.path.join(script_dir, "*.xlsx"))
    )

    def on_open():
        sel = listbox.curselection()
        if not sel:
            return
        result[0] = existing[sel[0]]
        root.destroy()

    def on_create():
        name = entry.get().strip()
        if not name:
            return
        safe = "".join(c for c in name if c not in r'\/:*?"<>|')
        result[0] = f"kiwi_{safe}.xlsx"
        root.destroy()

    def on_return(event=None):
        if entry.get().strip():
            on_create()
        else:
            on_open()

    def on_close():
        result[0] = None
        root.destroy()

    root = tk.Tk()
    root.title("Kiwi Tool — Select Test")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    root.protocol("WM_DELETE_WINDOW", on_close)

    try:
        icon_img = ImageTk.PhotoImage(Image.open(_ICON_PATH))
        root.iconphoto(True, icon_img)
        root._icon_ref = icon_img
    except Exception:
        pass

    w, h = 500, 370
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    # ── Existing files ────────────────────────────────────────────────────────
    tk.Label(root, text="Continue existing test:",
             font=("Segoe UI", 11, "bold"), anchor="w").pack(fill="x", padx=16, pady=(14, 4))

    list_frame = tk.Frame(root)
    list_frame.pack(fill="both", expand=True, padx=16)

    scrollbar = tk.Scrollbar(list_frame, orient="vertical")
    listbox = tk.Listbox(list_frame, font=("Segoe UI", 10), height=7,
                         yscrollcommand=scrollbar.set, selectmode="single",
                         activestyle="dotbox")
    scrollbar.config(command=listbox.yview)
    scrollbar.pack(side="right", fill="y")
    listbox.pack(side="left", fill="both", expand=True)

    for f in existing:
        listbox.insert("end", f)

    if not existing:
        listbox.insert("end", "(no existing test files found)")
        listbox.config(state="disabled")

    def on_listbox_select(event):
        entry.delete(0, "end")
    listbox.bind("<<ListboxSelect>>", on_listbox_select)
    listbox.bind("<Double-Button-1>", lambda e: on_open())

    # ── Create new ────────────────────────────────────────────────────────────
    tk.Label(root, text="── or create new ──",
             font=("Segoe UI", 9), fg="gray").pack(pady=(10, 2))

    name_frame = tk.Frame(root)
    name_frame.pack(fill="x", padx=16)
    tk.Label(name_frame, text="Name:", font=("Segoe UI", 10)).pack(side="left")
    entry = tk.Entry(name_frame, font=("Segoe UI", 11), width=36)
    entry.pack(side="left", padx=(8, 0), ipady=4)

    def on_entry_key(event):
        listbox.selection_clear(0, "end")
    entry.bind("<Key>", on_entry_key)

    # ── Buttons ───────────────────────────────────────────────────────────────
    btn_frame = tk.Frame(root)
    btn_frame.pack(pady=14)
    tk.Button(btn_frame, text="Open Selected", font=("Segoe UI", 10),
              width=16, command=on_open).pack(side="left", padx=6)
    tk.Button(btn_frame, text="Create New", font=("Segoe UI", 10),
              width=16, command=on_create).pack(side="left", padx=6)

    root.bind("<Return>", on_return)
    root.bind("<Escape>", lambda e: on_close())

    root.update()
    root.lift()
    root.grab_set()
    entry.focus_force()

    root.mainloop()
    return result[0]


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # Set terminal window title
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW("Kiwi Tool")
    except Exception:
        pass

    # Generate kiwi icon
    _make_kiwi_icon(size=64)

    # Project picker
    global EXCEL_FILE
    chosen = _ask_project()
    if chosen is None:
        print("  No test selected. Exiting.")
        sys.exit(0)
    EXCEL_FILE = chosen

    print("=" * 56)
    print("  Kiwi Tool")
    print("=" * 56)

    try:
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("  WARNING: Not running as Administrator.")
            print("  Global hotkeys may not work in all windows.")
            print()
    except Exception:
        pass

    _init_excel_if_needed()

    print(f"  Hotkey : {HOTKEY}")
    print(f"  Output : {os.path.abspath(EXCEL_FILE)}")
    print()
    print("  Press Scroll Lock at any time to capture.")
    print("  Press Ctrl+C here to stop.")
    print("=" * 56)

    keyboard.add_hotkey(HOTKEY, _on_hotkey, suppress=False)
    keyboard.wait()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n  Kiwi Tool stopped.")
        sys.exit(0)

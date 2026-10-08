# Kiwi Tool — Technical Report

**Project:** QA Screenshot Tool  
**Language:** Python 3.10+  
**Platform:** Windows 10 / 11  
**Date:** 2026-04-09

---

## 1. Purpose

Kiwi Tool is a lightweight QA documentation aid. Testers press a single key
(**Scroll Lock**) to capture the active window at any moment. Each capture is
automatically embedded — at full resolution, with a label and optional note —
into a structured Excel workbook. Multiple test sessions are separated into
individual sheet tabs; multiple test projects each get their own Excel file.

The goal is to remove friction from QA documentation: no switching apps,
no manual pasting, no file management during testing.

---

## 2. Project Structure

```
QA Tool/
├── qa_screenshot_tool.py   Main application (all logic)
├── run.bat                 Launcher (double-click to start)
├── install.bat             One-time dependency installer
├── requirements.txt        Python package list
├── kiwi_icon.png           Auto-generated tray icon (PNG)
├── kiwi_icon.ico           Auto-generated tray icon (ICO, multi-size)
├── screenshots/            Raw PNG captures (archive)
└── kiwi_<testname>.xlsx    Excel output files (one per test project)
```

---

## 3. Technologies & Libraries

### 3.1 Python Standard Library

| Module | Used for | Why not another? |
|--------|----------|-----------------|
| `threading` | Run screenshot capture in a background thread so the hotkey listener never blocks | Built-in; no overhead |
| `tempfile` | Temporary PNG file for openpyxl embedding | Needed because openpyxl reads from disk, not memory |
| `os` / `sys` | File paths, existence checks, exit | Built-in |
| `math` | Trigonometry for drawing kiwi seed positions | Built-in |
| `datetime` | Timestamps on labels and sheet names | Built-in |
| `io` | In-memory byte streams | Built-in |
| `glob` | Listing existing `.xlsx` files at startup | Built-in; simpler than `os.walk` for flat directory |
| `tkinter` | All GUI dialogs (project picker, note dialog) | Ships with Python on Windows; no extra install needed |
| `ctypes` | Detect Administrator privilege, set console title | Built-in Windows bridge |

### 3.2 Third-Party Libraries

#### `keyboard` (v0.13.5+)
**What it does:** Registers a global hotkey that fires even when the terminal
is not in focus.  
**Why keyboard:** It intercepts input at the OS driver level, so Scroll Lock is
captured regardless of which application is active — including elevated-privilege
windows when the tool itself runs as Administrator.  
**Why not alternatives:**
- `pynput` works similarly but has less reliable suppression on Windows.
- `win32api` keyboard hooks are lower-level and require more boilerplate.

#### `mss` (v9.0.1+)
**What it does:** Captures a rectangular region of the screen at native
resolution using the Windows GDI API.  
**Why mss:** It is the fastest Python screen-capture library available. It
captures directly from the GPU framebuffer in BGRX format, requires no
external executables, and works correctly on multi-monitor setups.  
**Why not alternatives:**
- `PIL.ImageGrab` is slower and doesn't handle all multi-monitor configurations.
- `pyautogui` has more dependencies and is slower for this use case.
- `win32gui` + `BitBlt` requires significant boilerplate and doesn't support
  hardware-accelerated windows.

#### `openpyxl` (v3.1.2+)
**What it does:** Creates, reads, and writes `.xlsx` Excel files; embeds images;
sets row heights, column widths, and cell fonts.  
**Why openpyxl:** It is the standard Python library for Excel files that don't
require Microsoft Office to be installed. It has full image-embedding support.  
**Why not alternatives:**
- `xlsxwriter` is write-only (cannot open existing files to append).
- `xlrd` is read-only.
- `pandas` + `openpyxl` adds unnecessary overhead for this use case.

#### `Pillow` / PIL (v10.0.0+)
**What it does:** Converts raw screen capture bytes into an `Image` object,
draws the kiwi icon programmatically, and saves PNG files.  
**Why Pillow:** It is the de-facto standard Python imaging library. No
alternative has comparable breadth for image manipulation, format support,
and drawing primitives.

#### `pywin32` — `win32gui` (v306+)
**What it does:** Gets the handle of the currently active (foreground) window
and retrieves its exact pixel coordinates on screen.  
**Why win32gui:** It is the direct Python binding to the Windows API
(`GetForegroundWindow`, `GetWindowRect`). No other library provides this
specific information without additional overhead.

---

## 4. Architecture & Flow

### 4.1 Startup Flow

```
run.bat
  └─ python qa_screenshot_tool.py
        ├─ main()
        │    ├─ SetConsoleTitleW("Kiwi Tool")       # terminal title
        │    ├─ _make_kiwi_icon()                    # generate PNG + ICO
        │    ├─ _ask_project()                       # GUI: pick / create Excel file
        │    │    └─ returns chosen filename
        │    ├─ EXCEL_FILE = chosen
        │    ├─ _init_excel_if_needed()              # create or open workbook
        │    │    └─ creates new sheet tab for this session
        │    ├─ keyboard.add_hotkey("scroll lock", _on_hotkey)
        │    └─ keyboard.wait()                      # blocks; listens forever
```

### 4.2 Hotkey → Screenshot Flow

```
User presses Scroll Lock
  └─ _on_hotkey()                          # called on keyboard listener thread
        └─ threading.Thread(_safe_run)     # offload to avoid blocking hotkey
              ├─ _capture()
              │    ├─ win32gui.GetForegroundWindow()   # get active window handle
              │    ├─ win32gui.GetWindowRect()         # get pixel bounds
              │    └─ mss.grab(region) → PIL Image     # capture that region
              │
              ├─ _ask_note()                           # GUI dialog for note text
              │    ├─ User types note → returns string
              │    └─ User presses Escape  → returns None (discard)
              │
              └─ _insert(img, note)                    # embed into Excel
                    ├─ _screenshot_num += 1
                    ├─ img.save(screenshots/kiwi_NNNN_TIMESTAMP.png)
                    ├─ img.save(temp/kiwi_tmp_N.png)  # temp for openpyxl
                    ├─ openpyxl.load_workbook(EXCEL_FILE)   # always fresh from disk
                    ├─ ws = wb[_session_sheet]
                    ├─ ws.cell(label row) = "#N | timestamp — note"
                    ├─ ws.row_dimensions[...].height = image height
                    ├─ ws.add_image(XLImage anchored to column B)
                    ├─ _current_row advances past image
                    ├─ wb.save(EXCEL_FILE + ".tmp")   # atomic: write to temp first
                    ├─ os.replace(tmp → EXCEL_FILE)   # then rename (atomic on Windows)
                    └─ os.remove(temp PNG)
```

### 4.3 Project Picker Flow (`_ask_project`)

```
Dialog opens
  ├─ Scans working directory for *.xlsx files
  ├─ Shows list of found files
  │
  ├─ User double-clicks existing file   → returns filename
  ├─ User selects + clicks Open         → returns filename
  ├─ User types name + clicks Create New→ returns "kiwi_<safe_name>.xlsx"
  └─ User closes / Escape               → returns None → sys.exit(0)
```

### 4.4 Session Sheet Naming (`_init_excel_if_needed`)

```
base_name = "YYYY-MM-DD HH.MM"   (no colons — Excel sheet name restriction)

If EXCEL_FILE does not exist:
    Create new workbook, rename default sheet to base_name

If EXCEL_FILE exists:
    Load workbook
    If base_name already in sheet names → append " (2)", " (3)" etc.
    Create new sheet with unique name

Save workbook (sheet is now registered but empty)
```

---

## 5. Key Design Decisions

### 5.1 Always reload workbook from disk before saving

```python
wb = openpyxl.load_workbook(EXCEL_FILE)   # inside _insert(), every time
```

openpyxl holds an in-memory model of the workbook. If the file is modified
externally (e.g. the user edits it in Excel) while the tool is running,
keeping a stale in-memory copy would overwrite those changes on the next save.
Reloading from disk each time guarantees consistency.

### 5.2 Atomic save (write-then-rename)

```python
tmp_xlsx = EXCEL_FILE + ".tmp"
wb.save(tmp_xlsx)
os.replace(tmp_xlsx, EXCEL_FILE)
```

If the process is interrupted mid-save, the original file remains intact.
`os.replace` is atomic on Windows (it uses `MoveFileEx` internally), so there
is never a moment where the file is partially written.

### 5.3 Thread safety with `_lock`

The hotkey callback fires on a `keyboard` library thread. Screenshot capture,
the note dialog, and Excel insertion all happen on a separate `threading.Thread`
(daemon). The `threading.Lock` around `_insert` ensures that if the user
somehow triggers two captures simultaneously, they are serialised — no two
threads write to the file at the same time.

### 5.4 Temporary PNG for embedding

openpyxl's `Image` class reads from a file path, not from a bytes buffer.
A temporary file in `tempfile.gettempdir()` is used and deleted immediately
after the workbook is saved, keeping the working directory clean.

### 5.5 Per-session sheet tabs

Each run of the tool creates a new sheet tab named by the session timestamp
(e.g. `2026-04-09 14.30`). This means:
- All sessions for one test project live in one file (easy to compare).
- Sessions are visually separated and independently scrollable.
- No screenshots are ever overwritten or mixed between sessions.

### 5.6 Kiwi icon drawn in code, not bundled

`_make_kiwi_icon` + `_draw_kiwi` generate the application icon at runtime
using Pillow's `ImageDraw`. This avoids shipping a binary asset with the
project. The icon is a kiwi fruit cross-section rendered as concentric ellipses
with seeds placed at trigonometric intervals.

---

## 6. Configuration Constants

Located at the top of `qa_screenshot_tool.py`:

| Constant | Default | Description |
|----------|---------|-------------|
| `HOTKEY` | `"scroll lock"` | Key that triggers a capture |
| `EXCEL_FILE` | `"kiwi_screenshots.xlsx"` | Default filename (overridden by project picker) |
| `IMAGES_FOLDER` | `"screenshots"` | Folder where raw PNGs are saved |
| `CELL_SPACING` | `2` | Blank rows between screenshots in the sheet |
| `_START_COL` | `"B"` | Excel column where images are anchored |
| `_LABEL_COL` | `1` (A) | Column where the text label is written |
| `_EXCEL_PT_PER_PX` | `0.75` | Conversion factor: screen pixels → Excel points |

---

## 7. File Outputs

### Excel workbook (`kiwi_<testname>.xlsx`)

- One file per test project.
- One sheet tab per session, named `YYYY-MM-DD HH.MM`.
- Each screenshot occupies:
  - **Row N**: bold label — `#1  |  2026-04-09 14:32:05  —  user note`
  - **Rows N+1 … N+K**: image embedded at native pixel dimensions, anchored to column B.
  - Row height is calculated so the image fits exactly without distortion.

### PNG archive (`screenshots/`)

Every capture is also saved as a standalone PNG named:
`kiwi_NNNN_YYYYMMDD_HHMMSS.png`

This provides a backup and allows the images to be used outside Excel.

---

## 8. Complete Code Walkthrough

This section explains every function and block of code in `qa_screenshot_tool.py`
in the order they appear in the file.

---

### 8.1 Module-level imports

```python
import io, math, os, sys, tempfile, threading
from datetime import datetime
import tkinter as tk
import win32gui, keyboard, mss, openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image, ImageDraw, ImageTk
```

All libraries are imported at the top. `XLImage` is aliased to avoid a name
clash with Pillow's own `Image` class (both are called `Image` in their
respective packages).

---

### 8.2 User configuration block

```python
HOTKEY        = "scroll lock"
EXCEL_FILE    = "kiwi_screenshots.xlsx"
IMAGES_FOLDER = "screenshots"
CELL_SPACING  = 2
```

These four constants are the only values a user ever needs to change:
- `HOTKEY` — change to any key name recognised by the `keyboard` library (e.g. `"f12"`).
- `EXCEL_FILE` — the default filename. At runtime this is overridden by the project picker dialog.
- `IMAGES_FOLDER` — subfolder where raw PNG backups are stored.
- `CELL_SPACING` — how many blank rows to leave between screenshots in the sheet.

---

### 8.3 Internal constants

```python
_START_COL       = "B"
_LABEL_COL       = 1
_EXCEL_PT_PER_PX = 0.75
_PX_PER_ROW      = 15.0 / _EXCEL_PT_PER_PX
_ICON_PATH       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kiwi_icon.png")
_ICO_PATH        = os.path.join(os.path.dirname(os.path.abspath(__file__)), "kiwi_icon.ico")
```

- `_START_COL = "B"` — images are placed in column B, leaving column A for text labels.
- `_LABEL_COL = 1` — column A in numeric form (openpyxl uses 1-based column numbers for cells).
- `_EXCEL_PT_PER_PX = 0.75` — Excel measures row height in points; one screen pixel equals 0.75 points. This ratio is used to size rows correctly so images fit without distortion.
- `_PX_PER_ROW` — how many pixels tall one default Excel row is. Used to calculate how many rows an image needs.
- `_ICON_PATH` / `_ICO_PATH` — absolute paths to the icon files, computed relative to the script's own location so the tool works regardless of where it is launched from.

---

### 8.4 Global state variables

```python
_lock           = threading.Lock()
_current_row    = 1
_screenshot_num = 0
_session_sheet  = None
```

- `_lock` — a mutex. Prevents two simultaneous Scroll Lock presses from writing to the Excel file at the same time.
- `_current_row` — tracks which Excel row the next screenshot should go to. Starts at 1 (top of the sheet) and advances after each insert.
- `_screenshot_num` — a counter for screenshots in the current session. Used in filenames and labels (`#1`, `#2`, …).
- `_session_sheet` — the name of the sheet tab created when the tool starts. Set once in `_init_excel_if_needed` and used by `_insert` to find the right sheet every time.

---

### 8.5 `_make_kiwi_icon(size=64)`

```python
def _make_kiwi_icon(size=64):
    img = _draw_kiwi(size)
    img.save(_ICON_PATH, format="PNG")

    sizes  = [16, 32, 48, 64]
    frames = [_draw_kiwi(s) for s in sizes]
    frames[0].save(_ICO_PATH, format="ICO", append_images=frames[1:],
                   sizes=[(s, s) for s in sizes])
    return img
```

Called once at startup. It:
1. Draws the kiwi at 64×64 and saves it as a PNG (used by tkinter dialogs).
2. Draws the same kiwi at 16, 32, 48, and 64 pixels and saves all four into one
   `.ico` file. Windows uses different sizes in different contexts (taskbar,
   Alt+Tab, window title bar), so a multi-size ICO looks sharp everywhere.

---

### 8.6 `_draw_kiwi(size)`

```python
def _draw_kiwi(size):
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    cx = cy = size // 2
    r  = size // 2 - 1
    draw.ellipse([cx-r, cy-r, cx+r, cy+r], fill="#7B5230")   # brown skin
    r2 = r - int(r * 0.12)
    draw.ellipse([cx-r2, cy-r2, cx+r2, cy+r2], fill="#8DB640") # dark green flesh
    r3 = int(r2 * 0.72)
    draw.ellipse([cx-r3, cy-r3, cx+r3, cy+r3], fill="#A8CC5A") # light green
    r4 = int(r2 * 0.22)
    draw.ellipse([cx-r4, cy-r4, cx+r4, cy+r4], fill="#F5F0DC") # cream centre
    num_seeds  = 10
    seed_dist  = r2 * 0.52
    sw, sh_seed = max(2, size // 20), max(3, size // 13)
    for i in range(num_seeds):
        angle = (2 * math.pi * i / num_seeds) - math.pi / 2
        sx = cx + seed_dist * math.cos(angle)
        sy = cy + seed_dist * math.sin(angle)
        draw.ellipse([sx-sw//2, sy-sh_seed//2, sx+sw//2, sy+sh_seed//2], fill="#3D2B1F")
    return img
```

Draws a kiwi fruit cross-section as a series of concentric ellipses:
1. Transparent background (`RGBA` with alpha=0).
2. Outer brown ring (skin).
3. Dark green layer (flesh near skin).
4. Lighter green layer (inner flesh).
5. Cream-coloured centre.
6. 10 seeds placed in a circle using trigonometry (`cos`/`sin` at evenly
   spaced angles). All sizes scale proportionally so the icon looks correct
   at every pixel size.

---

### 8.7 `_px_to_pt(px)` and `_rows_for_height(h)`

```python
def _px_to_pt(px):
    return px * _EXCEL_PT_PER_PX

def _rows_for_height(h):
    return max(1, round(h / _PX_PER_ROW))
```

Two small conversion utilities:
- `_px_to_pt` — converts a pixel measurement to Excel points for setting row heights.
- `_rows_for_height` — given an image height in pixels, returns how many Excel
  rows are needed to display it at 1:1 scale. The `max(1, …)` ensures at least
  one row is always allocated.

---

### 8.8 `_init_excel_if_needed()`

```python
def _init_excel_if_needed():
    global _session_sheet

    base_name = datetime.now().strftime("%Y-%m-%d %H.%M")
    _session_sheet = base_name

    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = _session_sheet
    else:
        wb = openpyxl.load_workbook(EXCEL_FILE)
        existing = wb.sheetnames
        suffix = 2
        while _session_sheet in existing:
            _session_sheet = f"{base_name} ({suffix})"
            suffix += 1
        ws = wb.create_sheet(title=_session_sheet)

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions[_START_COL].width = 170
    wb.save(EXCEL_FILE)
    print(f"  Session sheet : {_session_sheet}")
```

Called once at startup after the project is chosen. It:
1. Builds the session sheet name from the current minute (`%H.%M` uses a dot
   instead of a colon because Excel forbids `:` in sheet names).
2. **New file:** creates a fresh workbook and renames the default sheet.
3. **Existing file:** loads the workbook and creates a new sheet. If a sheet
   with the same name already exists (tool opened twice in the same minute),
   it appends ` (2)`, ` (3)` etc. until the name is unique.
4. Sets column widths: column A (labels) to 28 characters wide; column B
   (images) to 170 characters wide (wide enough for a typical screenshot).
5. Saves the workbook so the empty sheet is persisted before any captures.

---

### 8.9 `_capture()`

```python
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
```

Captures exactly the active window:
1. `GetForegroundWindow()` returns the handle of whatever window the user is
   looking at — not the terminal, not the Kiwi Tool itself.
2. `GetWindowRect()` returns the window's pixel coordinates on the screen
   (including its title bar and borders).
3. `mss.grab(region)` captures only that rectangle. The raw data from `mss`
   is in BGRX byte order (Blue, Green, Red, padding), so `Image.frombytes`
   decodes it correctly with the `"BGRX"` decoder to produce a standard RGB image.
4. A guard check ensures the window is not minimised or zero-sized before capture.

---

### 8.10 `_ask_note()`

```python
def _ask_note():
    result = [None]

    def on_confirm(event=None):
        result[0] = entry.get().strip()
        root.destroy()

    def on_skip(event=None):
        result[0] = None
        root.destroy()

    root = tk.Tk()
    ...
    root.mainloop()
    return result[0]
```

Shows a small dialog asking the tester to add a note to the screenshot.

- **`result = [None]`** — a list used as a mutable container so the inner
  functions (`on_confirm`, `on_skip`) can write a value that `_ask_note` can
  then return. (Python closures can read outer variables but not reassign them
  unless they are mutable objects like a list.)
- **`on_confirm`** — reads the text from the entry box, strips whitespace, and
  stores it. An empty string is valid (no note text, but the screenshot is kept).
- **`on_skip`** — stores `None`. The caller (`_safe_run`) treats `None` as
  "discard this screenshot entirely."
- **`root.grab_set()`** — makes the dialog modal: the user cannot click
  anything else until they respond.
- **`root.after(100, …)`** — waits 100 ms before forcing focus to the entry
  box, giving the OS time to fully render the window first.

---

### 8.11 `_insert(img, note)`

The core function — embeds a screenshot into Excel. Runs inside `_lock` so
only one thread can execute it at a time.

```python
_screenshot_num += 1
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
```
Increments the session counter and records the exact time of capture.

```python
img.save(img_path, format="PNG")
```
Saves a permanent PNG backup in the `screenshots/` folder named
`kiwi_NNNN_YYYYMMDD_HHMMSS.png`.

```python
tmp_png = os.path.join(tempfile.gettempdir(), f"kiwi_tmp_{_screenshot_num}.png")
img.save(tmp_png, format="PNG")
```
Saves a second copy to the system temp folder. openpyxl needs a file path
(not a bytes buffer) to embed an image, and we don't want a temporary file
sitting in the project folder.

```python
wb = openpyxl.load_workbook(EXCEL_FILE)
ws = wb[_session_sheet]
```
Reloads the workbook fresh from disk every time. This is the critical pattern
that prevents data loss if the file was edited externally between captures.

```python
_current_row += 2
label = f"#{_screenshot_num}  |  {timestamp}"
if note:
    label += f"  —  {note}"
cell = ws.cell(row=_current_row, column=_LABEL_COL, value=label)
cell.font = openpyxl.styles.Font(bold=True, size=11)
ws.row_dimensions[_current_row].height = 18
```
Skips 2 rows for visual spacing, then writes a bold label in column A showing
the screenshot number, timestamp, and note.

```python
actual_w, actual_h = img.size
rows = _rows_for_height(actual_h)
img_row = _current_row + 1
for r in range(img_row, img_row + rows):
    ws.row_dimensions[r].height = _px_to_pt(actual_h / rows)
```
Calculates how many Excel rows the image needs and sets each of those rows to
an equal fraction of the image height so the image fills the space exactly.

```python
xl_img = XLImage(tmp_png)
xl_img.width  = actual_w
xl_img.height = actual_h
xl_img.anchor = f"{_START_COL}{img_row}"
ws.add_image(xl_img)
```
Creates an openpyxl image object at native pixel dimensions and anchors its
top-left corner to cell B{img_row}.

```python
_current_row = img_row + rows + CELL_SPACING

tmp_xlsx = EXCEL_FILE + ".tmp"
wb.save(tmp_xlsx)
os.replace(tmp_xlsx, EXCEL_FILE)

os.remove(tmp_png)
```
Advances `_current_row` past the image plus the spacing gap.
Saves atomically (temp file → rename) and cleans up the temp PNG.

---

### 8.12 `_on_hotkey()` and `_safe_run()`

```python
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
```

`_on_hotkey` is the function registered with the `keyboard` library. It is
called on the keyboard listener's internal thread. Doing heavy work there would
block further key detection, so it immediately spawns a daemon thread to run
`_safe_run` instead.

`_safe_run` is the top-level capture sequence:
1. Capture the screen.
2. Ask for a note (blocks until the dialog is closed).
3. If the user skipped (`None`), log and return without saving.
4. Otherwise embed into Excel.
5. Any exception is caught and printed rather than crashing the tool.

`daemon=True` means this thread is automatically killed if the main program
exits, preventing it from keeping the process alive unexpectedly.

---

### 8.13 `_ask_project()`

```python
def _ask_project():
    import glob as _glob
    result = [None]

    script_dir = os.path.dirname(os.path.abspath(__file__))
    existing = sorted(
        os.path.basename(f)
        for f in _glob.glob(os.path.join(script_dir, "*.xlsx"))
    )
```

Scans the tool's own directory for all `.xlsx` files and sorts them
alphabetically. Uses `os.path.abspath(__file__)` so the scan always looks in
the right folder regardless of where the user launched the tool from.

```python
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
```

`on_open` — reads which item is selected in the listbox and returns that
filename.

`on_create` — takes the text the user typed, strips characters that Windows
does not allow in filenames (`\ / : * ? " < > |`), then wraps it as
`kiwi_<name>.xlsx`. The `kiwi_` prefix keeps all project files visually grouped.

```python
    def on_return(event=None):
        if entry.get().strip():
            on_create()
        else:
            on_open()
```

Smart Enter key: if there is text in the entry box, create a new file;
otherwise open the selected one. This lets power users work entirely from
the keyboard.

```python
    if not existing:
        listbox.insert("end", "(no existing test files found)")
        listbox.config(state="disabled")
```

If no `.xlsx` files exist yet, the listbox shows a placeholder message and
is disabled (not selectable), guiding the user to type a new name.

```python
    listbox.bind("<Double-Button-1>", lambda e: on_open())
```

Double-clicking a file in the list immediately opens it without needing to
click the button.

---

### 8.14 `main()`

```python
def main():
    try:
        import ctypes
        ctypes.windll.kernel32.SetConsoleTitleW("Kiwi Tool")
    except Exception:
        pass
```
Sets the terminal window title to "Kiwi Tool" using the Windows API.
Wrapped in try/except so it fails silently on non-Windows systems.

```python
    _make_kiwi_icon(size=64)
```
Generates the icon files. Done before the project picker so the dialog can
display the kiwi icon.

```python
    global EXCEL_FILE
    chosen = _ask_project()
    if chosen is None:
        print("  No test selected. Exiting.")
        sys.exit(0)
    EXCEL_FILE = chosen
```
Shows the project picker. `global EXCEL_FILE` is required because Python needs
an explicit declaration to reassign a module-level variable inside a function.
If the user closed the dialog without choosing, exit cleanly.

```python
    try:
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("  WARNING: Not running as Administrator.")
    except Exception:
        pass
```
Checks whether the process has Administrator rights. If not, warns the user
that the hotkey may not work in some elevated windows.

```python
    _init_excel_if_needed()
    keyboard.add_hotkey(HOTKEY, _on_hotkey, suppress=False)
    keyboard.wait()
```
Sets up the Excel file/sheet, registers the hotkey, then calls
`keyboard.wait()` which blocks the main thread forever (until Ctrl+C).
`suppress=False` means the Scroll Lock key still performs its normal function
in addition to triggering the capture.

---

### 8.15 Entry point guard

```python
if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n  Kiwi Tool stopped.")
        sys.exit(0)
```

Standard Python pattern: `main()` only runs when the file is executed directly
(not imported as a module). `KeyboardInterrupt` is raised when the user presses
Ctrl+C in the terminal, which is the normal way to stop the tool. The message
is printed and the process exits with code 0 (success).

---

## 9. Known Limitations

| Limitation | Detail |
|------------|--------|
| Windows only | Uses `win32gui` (Windows API), `mss` Windows backend, and `ctypes.windll` |
| Administrator recommended | The `keyboard` library captures global hotkeys at the driver level; without admin rights it may miss keystrokes in elevated-privilege windows |
| Excel must not have the file locked exclusively | If Excel has the `.xlsx` open and has locked it (rare), the atomic save will fail |
| Sheet name max 31 chars | Excel limit; the timestamp format `YYYY-MM-DD HH.MM` is 16 chars, well within limit |

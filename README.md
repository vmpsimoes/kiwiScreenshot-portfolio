# Kiwi Tool — README

A QA screenshot tool that captures the active window with a hotkey and embeds
each screenshot directly into an Excel file, organised by test session.

---

## Requirements

- Windows 10 or 11
- Python 3.10 or newer  
  Download: https://www.python.org/downloads/  
  During install, tick **"Add Python to PATH"**

---

## First Time Setup

1. Open the `QA Tool` folder.
2. Double-click **`install.bat`**.  
   This installs all required Python libraries automatically.
3. Wait until you see `Done. You can now run the tool with run.bat`.
4. Press any key to close the installer window.

That's it. You only need to do this once.

---

## Running the Tool (Every Time)

1. Double-click **`run.bat`**.
2. A dialog appears — **Select Test**:

   | Action | How |
   |--------|-----|
   | Continue an existing test | Click the file name in the list, then click **Open Selected** (or double-click the file) |
   | Start a brand new test | Type a name in the **Name** field, then click **Create New** |
   | Cancel / exit | Close the dialog or press **Escape** |

3. The tool starts and a terminal window opens showing the status.
4. Switch to the window you want to capture.
5. Press **Scroll Lock** to take a screenshot.
6. A small dialog asks you to add a note (e.g. `"Login page — correct behaviour"`).
   - Press **Enter** or click **Save** to confirm.
   - Press **Escape** or click **Skip** to discard the screenshot.
7. The screenshot is saved to the Excel file automatically.
8. Repeat steps 4–7 as many times as needed.
9. When finished, press **Ctrl+C** in the terminal window, or just close it.

---

## Where Are My Files?

| Item | Location |
|------|----------|
| Excel reports | `QA Tool\kiwi_<testname>.xlsx` |
| Raw PNG images | `QA Tool\screenshots\` |

Each Excel file contains one **sheet tab per session** (named by date and time,
e.g. `2026-04-09 14.30`). All sessions for the same test are kept in one file.

---

## Tips

- Run as **Administrator** for the hotkey to work in all applications
  (right-click `run.bat` → *Run as administrator*).
- You can have the Excel file open in Excel while capturing —
  the tool saves safely without corrupting it.
- If you accidentally press Scroll Lock and don't want to save,
  just click **Skip** in the note dialog.

---

## Stopping the Tool

Press **Ctrl+C** in the terminal window, or simply close it.

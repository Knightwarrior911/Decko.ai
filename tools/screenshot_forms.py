"""Open carrier, show each UserForm modeless, screenshot it.

Usage: python tools/screenshot_forms.py
Saves: tools/_screens/frmExport.png, frmExecute.png, frmImportSlides.png
"""
import sys
import time
import ctypes
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DIR = REPO_ROOT / "tools" / "_screens"
OUT_DIR.mkdir(parents=True, exist_ok=True)

import win32com.client
import win32gui
import win32ui
import win32con
from PIL import ImageGrab


def find_window(title_substr: str):
    """Find a top-level window whose title contains the substring."""
    matches = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if title_substr.lower() in t.lower():
                matches.append((hwnd, t))
        return True
    win32gui.EnumWindows(cb, None)
    return matches


def grab_window(hwnd, out_path: Path):
    """Capture a window: raise to front, then ImageGrab from screen rect."""
    # Make process DPI-aware so screen coords match physical pixels
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass
    try:
        ctypes.windll.user32.ShowWindow(hwnd, 9)  # SW_RESTORE
    except Exception:
        pass
    # Use AttachThreadInput trick to allow SetForegroundWindow to work
    try:
        cur_thread = ctypes.windll.kernel32.GetCurrentThreadId()
        target_thread = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, 0)
        ctypes.windll.user32.AttachThreadInput(target_thread, cur_thread, True)
        ctypes.windll.user32.BringWindowToTop(hwnd)
        ctypes.windll.user32.SetForegroundWindow(hwnd)
        ctypes.windll.user32.AttachThreadInput(target_thread, cur_thread, False)
    except Exception as e:
        print(f"  WARN raise: {e}")
    time.sleep(0.8)

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    img = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    img.save(str(out_path))
    print(f"  saved {out_path} ({img.size}) rect=({left},{top},{right},{bottom})")


def show_and_grab(app, macro_name: str, form_title_hint: str, out_name: str):
    print(f"[run] {macro_name}")
    qualified = f"'{CARRIER.name}'!modUI.{macro_name}"
    app.Run(qualified)
    time.sleep(2.0)
    matches = find_window(form_title_hint)
    if not matches:
        print(f"  WARN: no window with '{form_title_hint}' in title")
        # Fallback: enumerate all PowerPoint-related windows
        all_matches = []
        def cb(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                t = win32gui.GetWindowText(hwnd)
                if t and ("PPT" in t or "PowerPoint" in t or "Editor" in t):
                    all_matches.append((hwnd, t))
            return True
        win32gui.EnumWindows(cb, None)
        for h, t in all_matches:
            print(f"    candidate hwnd={h} title={t!r}")
        if all_matches:
            hwnd = all_matches[0][0]
        else:
            return
    else:
        hwnd = matches[0][0]
        print(f"  hwnd={hwnd} title={matches[0][1]!r}")
    grab_window(hwnd, OUT_DIR / out_name)
    # Close form by sending Escape
    try:
        ctypes.windll.user32.PostMessageW(hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception as e:
        print(f"  WARN close: {e}")
    time.sleep(1.0)


def main():
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    pres = app.Presentations.Open(str(CARRIER), WithWindow=True)
    time.sleep(2.0)
    try:
        show_and_grab(app, "ExportSnapshot",      "Export Snapshot",      "frmExport.png")
        show_and_grab(app, "ExecuteInstructions", "Execute Instructions", "frmExecute.png")
        show_and_grab(app, "ImportSlides",        "Import Slides",        "frmImportSlides.png")
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        app.Quit()


if __name__ == "__main__":
    sys.exit(main())

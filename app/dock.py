"""Decko Desktop — dock-to-PowerPoint snap engine (SP6).

Windows-only. No COM. Pure ctypes + win32gui. Importable on any platform
(callables degrade to no-ops when win32 isn't available), so the smoke
gate can validate the surface in CI.

Public surface:
    find_ppt_window() -> Optional[int]
    compute_dock_rect(ppt_hwnd, width=380, min_height=600)
        -> tuple[int, int, int, int]
    start_dock_loop(decko_hwnd, on_dock_event=None) -> DockLoop
    stop_dock_loop(loop) -> None

Events emitted via on_dock_event(name: str, payload: dict):
    "move_resize"  payload={"rect": (x, y, w, h)}
    "minimize"     payload={}
    "restore"      payload={"rect": (x, y, w, h)}
    "slideshow_enter"  payload={}
    "slideshow_exit"   payload={"rect": (x, y, w, h)}
    "ppt_gone"     payload={"rect": (x, y, w, h)}  # centered default
"""
from __future__ import annotations

import ctypes
import sys
import threading
import time
from ctypes import wintypes
from dataclasses import dataclass, field
from typing import Callable, Optional

# Soft-import: on non-Windows we still provide callables that return
# safe defaults so smoke gates and unit imports work everywhere.
_IS_WIN = sys.platform.startswith("win")

if _IS_WIN:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
else:  # pragma: no cover — Decko Desktop is Windows-only at runtime.
    user32 = None
    kernel32 = None

# Win32 constants we use.
EVENT_OBJECT_LOCATIONCHANGE = 0x800B
EVENT_SYSTEM_FOREGROUND = 0x0003
EVENT_SYSTEM_MINIMIZESTART = 0x0016
EVENT_SYSTEM_MINIMIZEEND = 0x0017
WINEVENT_OUTOFCONTEXT = 0x0000
WINEVENT_SKIPOWNPROCESS = 0x0002

MONITOR_DEFAULTTONEAREST = 0x00000002
MONITOR_DEFAULTTOPRIMARY = 0x00000001

WM_QUIT = 0x0012

# Hook callback signature.
if _IS_WIN:
    WINEVENTPROC = ctypes.WINFUNCTYPE(
        None,
        wintypes.HANDLE, wintypes.DWORD, wintypes.HWND,
        wintypes.LONG, wintypes.LONG, wintypes.DWORD, wintypes.DWORD,
    )

PPT_CLASS_NAMES = ("PPTFrameClass",)        # main PPT window
SLIDESHOW_CLASS_NAMES = ("screenClass",)    # F5 slideshow


# ----------------------------------------------------------------------
# Win32 wrappers (return safe defaults off-Windows).

def _get_window_text(hwnd: int) -> str:
    if not _IS_WIN or not hwnd:
        return ""
    n = user32.GetWindowTextLengthW(hwnd)
    buf = ctypes.create_unicode_buffer(n + 1)
    user32.GetWindowTextW(hwnd, buf, n + 1)
    return buf.value


def _get_class_name(hwnd: int) -> str:
    if not _IS_WIN or not hwnd:
        return ""
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _is_window_visible(hwnd: int) -> bool:
    return bool(_IS_WIN and hwnd and user32.IsWindowVisible(hwnd))


def _is_window(hwnd: int) -> bool:
    return bool(_IS_WIN and hwnd and user32.IsWindow(hwnd))


def _is_iconic(hwnd: int) -> bool:
    return bool(_IS_WIN and hwnd and user32.IsIconic(hwnd))


def _get_window_rect(hwnd: int) -> Optional[tuple[int, int, int, int]]:
    if not _IS_WIN or not hwnd:
        return None
    r = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(r)):
        return None
    return (r.left, r.top, r.right, r.bottom)


def _get_foreground_window() -> int:
    if not _IS_WIN:
        return 0
    return int(user32.GetForegroundWindow() or 0)


def _enum_top_windows() -> list[int]:
    hits: list[int] = []
    if not _IS_WIN:
        return hits
    EnumWindowsProc = ctypes.WINFUNCTYPE(
        wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    def _cb(hwnd, _lparam):
        if user32.IsWindowVisible(hwnd):
            hits.append(int(hwnd))
        return True

    user32.EnumWindows(EnumWindowsProc(_cb), 0)
    return hits


class _MONITORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("rcMonitor", wintypes.RECT),
        ("rcWork", wintypes.RECT),
        ("dwFlags", wintypes.DWORD),
    ]


def _monitor_rect_for_hwnd(hwnd: int) -> tuple[int, int, int, int]:
    """Returns the monitor *work area* (excludes taskbar) for hwnd, or
    the primary monitor's work area when hwnd is invalid."""
    if not _IS_WIN:
        return (0, 0, 1920, 1080)
    flag = MONITOR_DEFAULTTONEAREST if hwnd else MONITOR_DEFAULTTOPRIMARY
    h_mon = user32.MonitorFromWindow(hwnd, flag)
    info = _MONITORINFO()
    info.cbSize = ctypes.sizeof(_MONITORINFO)
    if not user32.GetMonitorInfoW(h_mon, ctypes.byref(info)):
        return (0, 0, 1920, 1080)
    r = info.rcWork
    return (r.left, r.top, r.right, r.bottom)


# ----------------------------------------------------------------------
# Public callables.

# Last-foreground PPT hwnd, updated by the hook callback when running.
_last_ppt_hwnd: int = 0
# Last-foreground slideshow hwnd.
_last_slideshow_hwnd: int = 0


def _is_ppt_main(hwnd: int) -> bool:
    if not _is_window_visible(hwnd):
        return False
    cls = _get_class_name(hwnd)
    if cls in PPT_CLASS_NAMES:
        return True
    title = _get_window_text(hwnd)
    return title.endswith(" - PowerPoint") or title.endswith("- PowerPoint")


def _is_slideshow(hwnd: int) -> bool:
    if not _is_window_visible(hwnd):
        return False
    return _get_class_name(hwnd) in SLIDESHOW_CLASS_NAMES


def find_ppt_window() -> Optional[int]:
    """Pick the PowerPoint window Decko should dock against.

    Selection order:
      1. The last-known foreground PPT hwnd, if still a live PPT window.
      2. Any visible top-level window matching PPT class or " - PowerPoint"
         title.
    Slideshow windows are excluded.
    """
    global _last_ppt_hwnd
    if _last_ppt_hwnd and _is_window(_last_ppt_hwnd) and _is_ppt_main(_last_ppt_hwnd):
        return _last_ppt_hwnd
    for h in _enum_top_windows():
        if _is_ppt_main(h):
            _last_ppt_hwnd = h
            return h
    _last_ppt_hwnd = 0
    return None


def compute_dock_rect(ppt_hwnd: Optional[int],
                      width: int = 380,
                      min_height: int = 600) -> tuple[int, int, int, int]:
    """Returns (x, y, w, h) screen pixels for Decko's window.

    When ppt_hwnd is falsy or invalid, returns a centered default rect on
    the primary monitor work area.
    """
    width = int(width)
    min_height = int(min_height)
    if not ppt_hwnd or not _is_window(int(ppt_hwnd)):
        mleft, mtop, mright, mbottom = _monitor_rect_for_hwnd(0)
        h = max(min_height, (mbottom - mtop) - 80)
        x = mleft + ((mright - mleft) - width) // 2
        y = mtop + ((mbottom - mtop) - h) // 2
        return (x, y, width, h)
    p = _get_window_rect(int(ppt_hwnd))
    if p is None:
        return compute_dock_rect(0, width, min_height)
    p_left, p_top, p_right, p_bottom = p
    m_left, m_top, m_right, m_bottom = _monitor_rect_for_hwnd(int(ppt_hwnd))
    h = max(min_height, p_bottom - p_top)
    y = max(p_top, m_top)
    # Prefer docking flush to the right edge; if no room, overlap PPT.
    if p_right + width <= m_right:
        x = p_right
    else:
        x = max(m_left, p_right - width)
    return (int(x), int(y), int(width), int(h))


# ----------------------------------------------------------------------
# Dock loop.

@dataclass
class DockLoop:
    decko_hwnd: int
    on_dock_event: Optional[Callable[[str, dict], None]] = None
    _hooks: list[int] = field(default_factory=list)
    _msg_thread: Optional[threading.Thread] = None
    _poll_thread: Optional[threading.Thread] = None
    _stop: threading.Event = field(default_factory=threading.Event)
    _state: str = "normal"  # "normal" | "minimized" | "slideshow"
    _last_emitted_rect: Optional[tuple[int, int, int, int]] = None
    # Hold a strong reference to the WINEVENTPROC so it isn't GC'd while
    # Windows holds the pointer.
    _proc_ref: object = None


def _emit(loop: DockLoop, name: str, payload: dict) -> None:
    if loop.on_dock_event is None:
        return
    try:
        loop.on_dock_event(name, payload)
    except Exception:  # noqa: BLE001
        # Hook callbacks must never raise — they're called from Win32.
        pass


def _recompute_and_emit(loop: DockLoop, reason: str) -> None:
    """Top-level event router. Called from hook callback AND polling
    fallback. Decides between move_resize / minimize / restore /
    slideshow_enter / slideshow_exit / ppt_gone."""
    global _last_slideshow_hwnd
    fg = _get_foreground_window()
    if fg and _is_slideshow(fg):
        _last_slideshow_hwnd = fg
        if loop._state != "slideshow":
            loop._state = "slideshow"
            _emit(loop, "slideshow_enter", {})
        return
    # Foreground isn't a slideshow window. If we were in slideshow, exit.
    if loop._state == "slideshow":
        loop._state = "normal"
        rect = compute_dock_rect(find_ppt_window())
        loop._last_emitted_rect = rect
        _emit(loop, "slideshow_exit", {"rect": rect})
        return

    ppt = find_ppt_window()
    if ppt is None:
        # No live PPT — recenter and tell the app it's gone (debounced).
        if loop._state != "ppt_gone":
            loop._state = "ppt_gone"
            rect = compute_dock_rect(None)
            loop._last_emitted_rect = rect
            _emit(loop, "ppt_gone", {"rect": rect})
        return
    if _is_iconic(ppt):
        if loop._state != "minimized":
            loop._state = "minimized"
            _emit(loop, "minimize", {})
        return
    # Restore from minimize/gone if needed.
    rect = compute_dock_rect(ppt)
    if loop._state in ("minimized", "ppt_gone"):
        loop._state = "normal"
        loop._last_emitted_rect = rect
        _emit(loop, "restore", {"rect": rect})
        return
    # Move/resize: only emit when the rect actually changed.
    if rect != loop._last_emitted_rect:
        loop._last_emitted_rect = rect
        loop._state = "normal"
        _emit(loop, "move_resize", {"rect": rect})


def _hook_callback_factory(loop: DockLoop):
    """Build the WINEVENTPROC closure tied to this loop instance."""
    global _last_ppt_hwnd

    def _cb(hWinEventHook, event, hwnd, idObject, idChild,
            dwEventThread, dwmsEventTime):
        # Track foreground PPT for find_ppt_window's last-foreground cache.
        try:
            if event == EVENT_SYSTEM_FOREGROUND and _is_ppt_main(int(hwnd)):
                globals()["_last_ppt_hwnd"] = int(hwnd)
            _recompute_and_emit(loop, "hook")
        except Exception:  # noqa: BLE001
            pass

    return _cb


def _msg_pump(loop: DockLoop) -> None:
    """Run a GetMessage loop on this thread so SetWinEventHook delivers
    callbacks. Exits when WM_QUIT is posted."""
    if not _IS_WIN:
        return
    msg = wintypes.MSG()
    while not loop._stop.is_set():
        # PeekMessage non-blocking so we can poll _stop quickly.
        if user32.PeekMessageW(ctypes.byref(msg), 0, 0, 0, 1):  # PM_REMOVE
            if msg.message == WM_QUIT:
                break
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        else:
            time.sleep(0.01)


def _poll_fallback(loop: DockLoop) -> None:
    """Used if SetWinEventHook fails to register."""
    while not loop._stop.is_set():
        try:
            _recompute_and_emit(loop, "poll")
        except Exception:  # noqa: BLE001
            pass
        time.sleep(0.25)


def start_dock_loop(decko_hwnd: int,
                    on_dock_event: Optional[Callable[[str, dict], None]] = None
                    ) -> DockLoop:
    """Start watching PPT and snapping Decko to it.

    Returns a DockLoop handle. Off-Windows, returns a no-op loop with
    the same fields populated so callers can still call stop_dock_loop.
    """
    loop = DockLoop(decko_hwnd=int(decko_hwnd or 0),
                    on_dock_event=on_dock_event)

    if not _IS_WIN:
        return loop

    proc = WINEVENTPROC(_hook_callback_factory(loop))
    loop._proc_ref = proc  # keep strong ref so GC doesn't free the cb

    # Try to install hooks for the events we care about.
    handles: list[int] = []
    pairs = [
        (EVENT_OBJECT_LOCATIONCHANGE, EVENT_OBJECT_LOCATIONCHANGE),
        (EVENT_SYSTEM_MINIMIZESTART, EVENT_SYSTEM_MINIMIZEEND),
        (EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND),
    ]
    for emin, emax in pairs:
        h = user32.SetWinEventHook(
            emin, emax, 0, proc, 0, 0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS)
        if h:
            handles.append(int(h))
    loop._hooks = handles

    if handles:
        t = threading.Thread(target=_msg_pump, args=(loop,),
                             name="decko-dock-msgpump", daemon=True)
        t.start()
        loop._msg_thread = t
    else:
        # Hooks failed; fall back to polling.
        t = threading.Thread(target=_poll_fallback, args=(loop,),
                             name="decko-dock-poll", daemon=True)
        t.start()
        loop._poll_thread = t

    # Prime the initial snap.
    try:
        _recompute_and_emit(loop, "init")
    except Exception:  # noqa: BLE001
        pass

    return loop


def stop_dock_loop(loop: Optional[DockLoop]) -> None:
    if loop is None:
        return
    loop._stop.set()
    if _IS_WIN:
        for h in loop._hooks:
            try:
                user32.UnhookWinEvent(h)
            except Exception:  # noqa: BLE001
                pass
        if loop._msg_thread is not None:
            try:
                user32.PostThreadMessageW(
                    ctypes.windll.kernel32.GetCurrentThreadId(), WM_QUIT, 0, 0)
            except Exception:  # noqa: BLE001
                pass
    loop._hooks = []
    # Worker threads are daemon=True; they'll exit when _stop is set.

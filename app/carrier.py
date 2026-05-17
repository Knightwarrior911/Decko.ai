"""Locate the bundled carrier and copy it to %APPDATA%\\Decko\\engine
on first run only. Never run update_macros on the user's machine
(spec §4)."""
import shutil
import sys
from pathlib import Path

from app.config import CARRIER_NAME, INSTALLED_CARRIER, ensure_app_dirs


def _bundled_carrier() -> Path:
    # PyInstaller unpacks data to sys._MEIPASS; dev fallback = repo root.
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / CARRIER_NAME


def ensure_carrier() -> Path:
    ensure_app_dirs()
    INSTALLED_CARRIER.parent.mkdir(parents=True, exist_ok=True)
    if not INSTALLED_CARRIER.exists():
        shutil.copy2(_bundled_carrier(), INSTALLED_CARRIER)
    return INSTALLED_CARRIER

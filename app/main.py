"""Entry point (placeholder — full pywebview UI is built in a later
SP1 task). `--selfcheck` proves the PyInstaller-bundled carrier
resolves and copies, so packaging_smoke can verify the SHIPPED exe."""
import sys

from app.carrier import _bundled_carrier, ensure_carrier
from app.config import ensure_app_dirs


def _selfcheck() -> int:
    ensure_app_dirs()
    src = _bundled_carrier()
    if not src.exists():
        print(f"SELFCHECK FAIL: bundled carrier missing at {src}")
        return 1
    dest = ensure_carrier()
    if not dest.exists():
        print(f"SELFCHECK FAIL: carrier not installed at {dest}")
        return 1
    print(f"SELFCHECK OK: bundled={src} installed={dest}")
    return 0


def main() -> int:
    if "--selfcheck" in sys.argv[1:]:
        return _selfcheck()
    print("Decko Desktop — UI not yet built (placeholder entry).")
    return 0


if __name__ == "__main__":
    sys.exit(main())

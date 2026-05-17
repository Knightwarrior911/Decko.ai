"""Entry point (placeholder — full pywebview UI is built in a later
SP1 task). Importing the core must not fail under PyInstaller."""
import sys

from app.carrier import ensure_carrier  # noqa: F401
from app.config import ensure_app_dirs  # noqa: F401


def main() -> int:
    print("Decko Desktop — UI not yet built (placeholder entry).")
    return 0


if __name__ == "__main__":
    sys.exit(main())

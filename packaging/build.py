"""Build pipeline. Bakes current src/ into the carrier via the existing
update_macros.py (spec §4 carrier provenance), then runs PyInstaller."""
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def main() -> int:
    print("[build] baking carrier via update_macros.py")
    r = subprocess.run([sys.executable, "update_macros.py"], cwd=REPO)
    if r.returncode != 0:
        print("[build] update_macros failed")
        return 1
    print("[build] PyInstaller")
    r = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm",
         str(REPO / "packaging" / "decko.spec")], cwd=REPO)
    return r.returncode


if __name__ == "__main__":
    sys.exit(main())

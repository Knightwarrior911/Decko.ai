"""packaging_smoke: PyInstaller build succeeds, the SHIPPED exe's
bundled carrier resolves (Decko.exe --selfcheck), AND that carrier
loads via COM. Exit 0 only on PASS."""
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def main() -> int:
    r = subprocess.run([sys.executable, "packaging/build.py"], cwd=REPO)
    if r.returncode != 0:
        print("FAIL: packaging build failed")
        return 1
    exe = REPO / "dist" / "Decko.exe"
    if not exe.exists():
        print(f"FAIL: {exe} missing")
        return 1

    # Prove the SHIPPED exe's bundled carrier resolves + installs.
    sc = subprocess.run([str(exe), "--selfcheck"], cwd=REPO,
                        capture_output=True, text=True, timeout=180)
    print(sc.stdout.strip())
    if sc.returncode != 0 or "SELFCHECK OK" not in sc.stdout:
        print(f"FAIL: exe --selfcheck failed (rc={sc.returncode})\n"
              f"{sc.stdout}\n{sc.stderr}")
        return 1

    # The installed (from-bundle) carrier must load via COM.
    from app.config import INSTALLED_CARRIER
    os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
    time.sleep(2.0)
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    try:
        p = app.Presentations.Open(str(INSTALLED_CARRIER), WithWindow=True)
        n = len(app.Run("PPT_AI_Editor!GetAllActionTypes"))
        p.Saved = True
        p.Close()
    finally:
        app.Quit()
        time.sleep(1.0)
    if n < 1000:
        print(f"FAIL: GetAllActionTypes too short ({n})")
        return 1
    print(f"  packaging_smoke: Decko.exe built; bundled carrier "
          f"self-checked + loads via COM (len={n})")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())

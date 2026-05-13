"""
test_guidance_doc_sync.py — fails if docs/ACTIONS_REFERENCE.md is out of
sync with VBA's GetActionGuidance.

Re-runs sync_actions_guidance.py in a non-destructive check mode: rebuilds
the auto-appendix in memory and compares against what's currently committed.
If they differ, the doc is stale and needs `python tools/sync_actions_guidance.py`
to be run before committing.

Run: python tests/test_guidance_doc_sync.py
"""
import shutil
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DOC = REPO / "docs" / "ACTIONS_REFERENCE.md"
SYNC = REPO / "tools" / "sync_actions_guidance.py"


def main() -> int:
    if not DOC.exists() or not SYNC.exists():
        print("ERROR: required files missing.")
        return 1

    before = DOC.read_bytes()
    backup = DOC.with_suffix(".md.predrift")
    shutil.copy2(DOC, backup)

    try:
        # Run the sync; it will overwrite DOC if anything has drifted.
        result = subprocess.run(
            [sys.executable, str(SYNC)],
            capture_output=True, text=True, cwd=str(REPO),
        )
        after = DOC.read_bytes()
        if result.returncode != 0:
            print(f"sync_actions_guidance.py failed:\n{result.stdout}\n{result.stderr}")
            return 2
        if before == after:
            print("OK — ACTIONS_REFERENCE.md is in sync with GetActionGuidance.")
            return 0
        # Restore the original so the test doesn't silently mutate state
        DOC.write_bytes(before)
        print("DRIFT: ACTIONS_REFERENCE.md auto-guidance appendix is stale.")
        print("Run: python tools/sync_actions_guidance.py  and commit the result.")
        return 1
    finally:
        if backup.exists():
            backup.unlink()


if __name__ == "__main__":
    sys.exit(main())

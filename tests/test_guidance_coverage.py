"""
test_guidance_coverage.py — ensures GetActionGuidance has a Case entry for
EVERY action type the validator/dispatcher knows about.

If this test reports any uncovered actions, add a Case for each one to
modExecuteInstructions.GetActionGuidance so Fix Errors covers them.
"""

import re
import sys
from pathlib import Path
import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
DISPATCH_BAS = REPO_ROOT / "src" / "modExecuteInstructions.bas"

FALLBACK_MARKER = "No canonical guidance entry"


def extract_all_action_types() -> set[str]:
    """Parse modExecuteInstructions.bas for every Case '...' label in the
    DispatchAction routine — these are the canonical action types."""
    text = DISPATCH_BAS.read_text(encoding="utf-8", errors="replace")
    # Match: Case "name" or Case "a", "b", "c" — capture every quoted string
    types: set[str] = set()
    for line in text.splitlines():
        s = line.strip()
        if not s.startswith("Case "):
            continue
        if s.startswith("Case Else"):
            continue
        for m in re.finditer(r'"([^"]+)"', s):
            types.add(m.group(1))
    return types


def main() -> int:
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break
    carrier = app.Presentations.Open(str(CARRIER))

    types = extract_all_action_types()
    print(f"Found {len(types)} distinct action-type strings in dispatcher.\n")

    missing = []
    for t in sorted(types):
        guide = app.Run(
            "PPT_AI_Editor.pptm!modExecuteInstructions.GetActionGuidance", t
        )
        if FALLBACK_MARKER in guide:
            missing.append(t)

    carrier.Close()

    if missing:
        print(f"COVERAGE GAP: {len(missing)} action(s) hit the fallback "
              f"guidance — add Case entries for these:")
        for t in missing:
            print(f"  - {t}")
        return 1

    print(f"OK — every one of {len(types)} action types has canonical guidance.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

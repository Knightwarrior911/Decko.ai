"""
stress_scan/harness.py — stress-test the scan_palette action.

Opens JAZZ-Pitch-Book.pptx via COM, runs scan_palette (deck scope),
reads output from %TEMP%\decko_palette.json, parses JSON, asserts correctness.

Assertions:
  - At least 5 colors returned
  - Each entry has hex, count, roles fields
  - hex is "#RRGGBB" format (7 chars, starts with #)
  - count is a positive integer
  - roles is a non-empty list containing only "fill", "font", "border"
  - Output is sorted by count descending

Exit 0 on pass, 1 on any failure.
Run: python tests/stress_scan/harness.py
"""

import sys, json, time, os
from pathlib import Path

import win32com.client

ROOT      = Path(__file__).resolve().parents[2]
JAZZ      = Path(r"C:\Users\vinit\Downloads\JAZZ-Pitch-Book.pptx")
CARRIER   = ROOT / "PPT_AI_Editor.pptm"
TEMP_JSON = Path(os.environ.get("TEMP", "/tmp")) / "decko_palette.json"

VALID_ROLES = {"fill", "font", "border"}


def read_palette_output() -> str:
    """Read scan_palette output from temp file written by VBA."""
    if not TEMP_JSON.exists():
        raise FileNotFoundError(f"Temp file not found: {TEMP_JSON}")
    return TEMP_JSON.read_text(encoding="utf-8")


def run_batch(app, actions: list) -> str:
    payload = json.dumps({"actions": actions, "verify_after": False}, ensure_ascii=True)
    return str(app.Run("PPT_AI_Editor.pptm!modExecuteInstructions.ExecuteFromString", payload))


def main() -> int:
    if not JAZZ.exists():
        print(f"ERROR: {JAZZ} not found")
        return 1
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found")
        return 1

    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True

    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    carrier = app.Presentations.Open(str(CARRIER.resolve()), WithWindow=False)
    time.sleep(1)

    deck = app.Presentations.Open(str(JAZZ.resolve()), WithWindow=True)
    deck.Windows(1).Activate()
    time.sleep(1)

    ok = True
    try:
        # ── deck-scope scan ───────────────────────────────────────────────────
        print("Running scan_palette (scope: deck)...")
        result = run_batch(app, [{"type": "scan_palette", "scope": "deck"}])
        print(f"  Batch result: {result.strip()[:120]}")

        time.sleep(0.5)
        clip = read_palette_output()
        print(f"  Temp file ({len(clip)} chars): {clip[:300]}")

        # ── parse ─────────────────────────────────────────────────────────────
        try:
            colors = json.loads(clip)
        except json.JSONDecodeError as e:
            print(f"FAIL: output is not valid JSON: {e}\n{clip[:400]}")
            return 1

        # ── assert: >= 5 colors ───────────────────────────────────────────────
        if len(colors) < 5:
            print(f"FAIL: expected >= 5 colors, got {len(colors)}")
            ok = False
        else:
            print(f"  OK: {len(colors)} colors returned")

        # ── assert: field presence + types ────────────────────────────────────
        for idx, c in enumerate(colors):
            if not isinstance(c, dict):
                print(f"FAIL: entry {idx} is not a dict: {c}")
                ok = False
                continue
            missing = {"hex", "count", "roles"} - c.keys()
            if missing:
                print(f"FAIL: entry {idx} missing fields {missing}: {c}")
                ok = False
                continue
            h = c["hex"]
            if not (isinstance(h, str) and len(h) == 7 and h.startswith("#")):
                print(f"FAIL: entry {idx} bad hex: {h!r}")
                ok = False
            cnt = c["count"]
            if not (isinstance(cnt, int) and cnt > 0):
                print(f"FAIL: entry {idx} bad count: {cnt!r}")
                ok = False
            roles = c["roles"]
            if not (isinstance(roles, list) and len(roles) > 0):
                print(f"FAIL: entry {idx} roles must be non-empty list: {roles!r}")
                ok = False
            elif not set(roles).issubset(VALID_ROLES):
                print(f"FAIL: entry {idx} unknown roles: {set(roles) - VALID_ROLES}")
                ok = False

        # ── assert: sorted by count desc ──────────────────────────────────────
        counts = [c["count"] for c in colors]
        if counts != sorted(counts, reverse=True):
            print(f"FAIL: not sorted by count desc: {counts[:10]}")
            ok = False
        else:
            print(f"  OK: sorted by count desc")

        # ── single-slide scan (slide 1) ───────────────────────────────────────
        print("\nRunning scan_palette (scope: slide:1)...")
        result2 = run_batch(app, [{"type": "scan_palette", "scope": "slide:1"}])
        print(f"  Batch result: {result2.strip()[:120]}")
        time.sleep(0.5)
        clip2 = read_palette_output()
        try:
            colors2 = json.loads(clip2)
            print(f"  OK: slide:1 scan returned {len(colors2)} colors")
        except json.JSONDecodeError:
            print(f"FAIL: slide:1 scan output not valid JSON: {clip2[:200]}")
            ok = False

        # ── print top-10 for reference ────────────────────────────────────────
        print("\nTop-10 colors (deck scan):")
        for c in colors[:10]:
            print(f"  {c['hex']}  count={c['count']}  roles={c['roles']}")

    finally:
        try:
            deck.Close()
        except Exception:
            pass
        try:
            carrier.Saved = True
            carrier.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

    if ok:
        print("\nAll assertions passed.")
        return 0
    print("\nFAIL: one or more assertions failed.")
    return 1


if __name__ == "__main__":
    sys.exit(main())

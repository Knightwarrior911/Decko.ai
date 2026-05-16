"""Executor partial-failure contract smoke test, via PowerPoint COM.

Run: python tests/run_smoke_failcontract.py

For a planted failing action at first / middle / last position in a
batch, asserts:
  (a) Reporting  — ExecuteFromString's returned summary names the exact
      batch index + type + reason of the failure (no silent swallow).
  (b) Applied    — every non-failing action's effect is visible in
      BuildSnapshotJson (loop continues past the failure).
  (c) No collateral — a shape NOT targeted by the batch is unchanged.

Each case runs on a freshly rebuilt build_problem_deck (ExecuteFromString
mutates irreversibly: no undo, no backup). Non-run actions only, so the
executed index == submitted position (ReorderForRunIndexSafety only
reorders run actions). verify_after:false keeps the summary clean.

Exit non-zero unless all three checks are 100%. No AI/API; deterministic.
"""
import json
import os
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"

sys.path.insert(0, str(REPO_ROOT / "tests"))
from test_verify_loop import build_problem_deck  # noqa: E402

# build_problem_deck slide 1: shape_id 2 ("Off-slide right"), 3 ("Dup A"),
# 4 ("Dup B"). Use 2 & 3 as valid targets, 4 as the untouched control
# (its original text "Dup B" must survive).
GOOD_A_ID, GOOD_B_ID, CONTROL_ID = 2, 3, 4
CONTROL_ORIG_TEXT = "Dup B"
BAD_ID = 999  # absent -> deterministic "shape_not_found: id=999"


def fail_action(tag):
    return {"type": "set_text", "slide": 1, "shape_id": BAD_ID, "value": tag + "_X"}


def good(idx_id, tag, n):
    return {"type": "set_text", "slide": 1, "shape_id": idx_id, "value": f"{tag}_{n}"}


def cases():
    # (name, actions, fail_index_1based, tag)
    return [
        ("fail_first", "FCF",
         lambda tag: [fail_action(tag), good(GOOD_A_ID, tag, "A"), good(GOOD_B_ID, tag, "B")], 1),
        ("fail_middle", "FCM",
         lambda tag: [good(GOOD_A_ID, tag, "A"), fail_action(tag), good(GOOD_B_ID, tag, "B")], 2),
        ("fail_last", "FCL",
         lambda tag: [good(GOOD_A_ID, tag, "A"), good(GOOD_B_ID, tag, "B"), fail_action(tag)], 3),
    ]


def norm(s):
    return (s or "").replace("\r\n", "\n").replace("\r", "\n")


def shape_text(snap, slide_idx0, shape_id):
    for sh in snap["slides"][slide_idx0]["shapes"]:
        if sh.get("shape_id") == shape_id:
            return (sh.get("text") or "").strip()
    return None


def open_app():
    last = None
    for _ in range(15):
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = True
            return app
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2.0)
    raise RuntimeError(f"PowerPoint COM bring-up failed: {last!r}")


def run_session():
    pythoncom.CoInitialize()
    app = open_app()
    carrier = None
    try:
        while app.Presentations.Count > 0:
            try:
                app.Presentations(1).Close()
            except Exception:
                break
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)

        fails = []
        for name, tagbase, build, fail_idx in [(c[0], c[1], c[2], c[3]) for c in cases()]:
            test_pres = build_problem_deck(app)
            test_pres.Windows(1).Activate()
            time.sleep(1.0)
            try:
                tag = tagbase
                actions = build(tag)
                summary = norm(app.Run(
                    "PPT_AI_Editor!ExecuteFromString",
                    json.dumps({"actions": actions, "verify_after": False})))

                # (a) reporting
                want = f"action #{fail_idx} set_text: shape_not_found: id={BAD_ID}"
                if want not in summary or "FAILURES (" not in summary:
                    fails.append((name, "(a) reporting",
                                  f"want {want!r} + 'FAILURES (' in: {summary!r}"))

                snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))

                # (b) non-failing actions applied
                ta = shape_text(snap, 0, GOOD_A_ID)
                tb = shape_text(snap, 0, GOOD_B_ID)
                if ta != f"{tag}_A":
                    fails.append((name, "(b) applied",
                                  f"shape {GOOD_A_ID} text {ta!r} != {tag+'_A'!r}"))
                if tb != f"{tag}_B":
                    fails.append((name, "(b) applied",
                                  f"shape {GOOD_B_ID} text {tb!r} != {tag+'_B'!r}"))

                # (c) no collateral mutation
                tc = shape_text(snap, 0, CONTROL_ID)
                if tc != CONTROL_ORIG_TEXT:
                    fails.append((name, "(c) collateral",
                                  f"control shape {CONTROL_ID} text {tc!r} "
                                  f"!= original {CONTROL_ORIG_TEXT!r}"))
            finally:
                try:
                    test_pres.Saved = True
                    test_pres.Close()
                except Exception:
                    pass
        return len(cases()), fails
    finally:
        try:
            if carrier is not None:
                carrier.Saved = True
                carrier.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(2.0)


def main() -> int:
    transient = (pywintypes.com_error, AttributeError, RuntimeError)
    result = None
    last = None
    for attempt in range(1, 4):
        try:
            result = run_session()
            break
        except transient as e:  # noqa: PERF203
            last = e
            print(f"  retry transient COM error (attempt {attempt}): {e!r}")
            os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
            time.sleep(5.0)
    if result is None:
        print(f"FAIL: failcontract run failed after retries: {last!r}")
        return 1

    n_cases, fails = result
    print(f"cases: {n_cases} (fail at first/middle/last)")
    print(f"checks: (a) reporting  (b) preceding+subsequent applied  (c) no collateral")
    if fails:
        for name, check, why in fails:
            print(f"  FAIL [{name}] {check}: {why}")
    else:
        print("  all checks passed for all cases")

    ok = not fails
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

"""modVerify precision/recall smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_verify.py

Builds the deliberate-defect deck (reusing build_problem_deck from
tests/test_verify_loop.py), runs modVerify through the carrier, and scores
the produced warnings against a FROZEN ground-truth contract.

Why a frozen file: the live verify run writes its output to
tests/test_verify_deck.pptx.warnings.json -- the very path the project
treats as the labelled contract. Scoring live-vs-itself is meaningless, so
on first run we snapshot that sidecar into tests/verify_ground_truth.json
(the immutable contract) and from then on always compare the fresh run
against the frozen file, restoring the live sidecar afterwards so running
the test never mutates the contract.

To intentionally change the contract (a detector change that is genuinely
correct), edit tests/verify_ground_truth.json deliberately with a one-line
justification -- never to make a failing run pass.

Match key per the goal: the tuple (kind, slide, shape_id, severity).
message and suggestion are advisory free text and are excluded.
Comparison is multiset-aware (the contract contains intentional
duplicates). Exit non-zero unless precision == 1.0 AND recall == 1.0.

No AI/API -- fully deterministic.
"""
import json
import os
import shutil
import sys
import time
from collections import Counter
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
TEST_DECK = REPO_ROOT / "tests" / "test_verify_deck.pptx"
LIVE_SIDECAR = Path(str(TEST_DECK) + ".warnings.json")
GROUND_TRUTH = REPO_ROOT / "tests" / "verify_ground_truth.json"

sys.path.insert(0, str(REPO_ROOT / "tests"))
from test_verify_loop import build_problem_deck  # noqa: E402

VERIFY_INSTRUCTIONS = json.dumps(
    {"actions": [], "verify_after": True, "verify_scope": "deck"}
)


def key(w: dict) -> tuple:
    return (w["kind"], w["slide"], w["shape_id"], w["severity"])


def load_warnings(path: Path) -> list:
    with path.open(encoding="utf-8") as f:
        return json.load(f).get("warnings", [])


def open_app():
    """COM-resilient PowerPoint bring-up (mirrors run_smoke.py)."""
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


def run_verify_once() -> list:
    """Build the defect deck, run verify, return the predicted warnings."""
    pythoncom.CoInitialize()
    app = open_app()
    carrier = None
    test_pres = None
    try:
        while app.Presentations.Count > 0:
            try:
                app.Presentations(1).Close()
            except Exception:
                break
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
        if LIVE_SIDECAR.exists():
            LIVE_SIDECAR.unlink()
        test_pres = build_problem_deck(app)
        test_pres.Windows(1).Activate()
        time.sleep(1.0)
        app.Run(
            "PPT_AI_Editor.pptm!modExecuteInstructions.ExecuteFromString",
            VERIFY_INSTRUCTIONS,
        )
        if not LIVE_SIDECAR.exists():
            raise RuntimeError("verify did not write the sidecar")
        return load_warnings(LIVE_SIDECAR)
    finally:
        for p in (test_pres, carrier):
            try:
                if p is not None:
                    p.Saved = True
                    p.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(2.0)


def main() -> int:
    if not GROUND_TRUTH.exists():
        if not LIVE_SIDECAR.exists():
            print(f"ERROR: no ground truth and no sidecar to seed from "
                  f"({GROUND_TRUTH} / {LIVE_SIDECAR})")
            return 1
        shutil.copy2(LIVE_SIDECAR, GROUND_TRUTH)
        print(f"[seed] froze current sidecar as contract -> {GROUND_TRUTH.name} "
              f"({len(load_warnings(GROUND_TRUTH))} warnings)")

    gt = load_warnings(GROUND_TRUTH)
    gt_backup = GROUND_TRUTH.read_bytes()

    # Preserve the live sidecar (it IS the contract path) across the run.
    sidecar_backup = LIVE_SIDECAR.read_bytes() if LIVE_SIDECAR.exists() else None

    transient = (pywintypes.com_error, AttributeError, RuntimeError)
    pred = None
    last = None
    for attempt in range(1, 4):
        try:
            pred = run_verify_once()
            break
        except transient as e:  # noqa: PERF203
            last = e
            print(f"  retry transient COM error (attempt {attempt}): {e!r}")
            os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
            time.sleep(5.0)
    try:
        if pred is None:
            print(f"FAIL: verify run failed after retries: {last!r}")
            return 1

        gt_c = Counter(key(w) for w in gt)
        pred_c = Counter(key(w) for w in pred)

        matched = sum((gt_c & pred_c).values())
        missed = gt_c - pred_c   # false negatives
        extra = pred_c - gt_c    # false positives

        n_pred = sum(pred_c.values())
        n_gt = sum(gt_c.values())
        precision = matched / n_pred if n_pred else 0.0
        recall = matched / n_gt if n_gt else 0.0
        f1 = (2 * precision * recall / (precision + recall)
              if (precision + recall) else 0.0)

        print(f"\nground truth: {n_gt} warnings   predicted: {n_pred} warnings")
        print(f"matched: {matched}")
        if missed:
            print(f"\nMISSED (false negatives) — in contract, not produced:")
            for k, c in sorted(missed.items()):
                print(f"  {c}x {k}")
        if extra:
            print(f"\nEXTRA (false positives) — produced, not in contract:")
            for k, c in sorted(extra.items()):
                print(f"  {c}x {k}")
        print(f"\nprecision={precision:.4f}  recall={recall:.4f}  f1={f1:.4f}")

        ok = (precision == 1.0 and recall == 1.0)
        print("RESULT:", "PASS" if ok else "FAIL")
        return 0 if ok else 1
    finally:
        # Never let running the test mutate the contract or lose the sidecar.
        GROUND_TRUTH.write_bytes(gt_backup)
        if sidecar_backup is not None:
            LIVE_SIDECAR.write_bytes(sidecar_backup)


if __name__ == "__main__":
    sys.exit(main())

"""apply_template smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_template.py   (exit 0 only at 100%)

Builds a fresh 960x540 deck, applies each template via ExecuteFromString
(appends a new slide), then via BuildSnapshotJson asserts the expected
slot text is present and the headline box geometry is within tolerance
of the deterministic layout. COM-resilient. No AI/network.
"""
import json
import os
import sys
import time
from unittest import mock  # noqa: F401  (kept for parity; unused)

import pythoncom
import pywintypes
import win32com.client

REPO_ROOT = __import__("pathlib").Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"

SWp, SHp, M, GAP = 960.0, 540.0, 36.0, 18.0
TOL = 3.0


def expect_box(frac_left, frac_top, frac_w, frac_h):
    return (M if frac_left is None else frac_left,
            SHp * frac_top, SWp - 2 * M if frac_w is None else frac_w,
            SHp * frac_h)


# template -> (content, [slot sentinels expected in some shape text],
#              headline sentinel, expected (left,top,width) of headline box)
CASES = {
    "title": (
        {"title": "TPL_TITLE", "subtitle": "TPL_SUB"},
        ["TPL_TITLE", "TPL_SUB"], "TPL_TITLE",
        (M, SHp * 0.34, SWp - 2 * M),
    ),
    "section": (
        {"section_number": "TPL_NUM", "section_title": "TPL_STITLE"},
        ["TPL_NUM", "TPL_STITLE"], "TPL_NUM",
        (M, SHp * 0.28, SWp - 2 * M),
    ),
    "bullets": (
        {"heading": "TPL_HEAD", "bullets": ["TPL_B1", "TPL_B2", "TPL_B3"]},
        ["TPL_HEAD", "TPL_B1", "TPL_B2", "TPL_B3"], "TPL_HEAD",
        (M, M, SWp - 2 * M),
    ),
    "two_col": (
        {"heading": "TPL_H", "left_body": "TPL_LB", "right_body": "TPL_RB"},
        ["TPL_H", "TPL_LB", "TPL_RB"], "TPL_H",
        (M, M, SWp - 2 * M),
    ),
    "comparison": (
        {"heading": "TPL_CH", "left_label": "TPL_LL", "left_body": "TPL_LBOD",
         "right_label": "TPL_RL", "right_body": "TPL_RBOD"},
        ["TPL_CH", "TPL_LL", "TPL_LBOD", "TPL_RL", "TPL_RBOD"], "TPL_CH",
        (M, M, SWp - 2 * M),
    ),
    "kpi_dashboard": (
        {"heading": "TPL_K", "tiles": [{"stat": "42%", "label": "TPL_GROW"},
                                       {"stat": "7", "label": "TPL_MKT"}]},
        ["TPL_K", "42%", "TPL_GROW", "7", "TPL_MKT"], "TPL_K",
        (M, M, SWp - 2 * M),
    ),
    "quote": (
        {"quote_text": "TPL_QUOTE", "attribution": "TPL_ATTR"},
        ["TPL_QUOTE", "TPL_ATTR"], "TPL_QUOTE",
        (M, SHp * 0.3, SWp - 2 * M),
    ),
}


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
    pres = None
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
        pres = app.Presentations.Add()
        pres.PageSetup.SlideWidth = SWp
        pres.PageSetup.SlideHeight = SHp
        pres.Windows(1).Activate()
        time.sleep(1.0)

        fails = []
        for tpl, (content, sentinels, head, (eL, eT, eW)) in CASES.items():
            instr = json.dumps({"actions": [
                {"type": "apply_template", "template": tpl, "content": content}
            ], "verify_after": False})
            summary = app.Run("PPT_AI_Editor!ExecuteFromString", instr)
            if "1 applied" not in summary:
                fails.append((tpl, "apply", f"summary={summary!r}"))
                continue
            snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
            shapes = snap["slides"][-1]["shapes"]
            texts = [(s.get("text") or "") for s in shapes]
            blob = " || ".join(texts)
            for sen in sentinels:
                if sen not in blob:
                    fails.append((tpl, "slot text", f"{sen!r} not in slide texts"))
            hs = next((s for s in shapes if head in (s.get("text") or "")), None)
            if hs is None:
                fails.append((tpl, "headline", f"no shape with {head!r}"))
            else:
                p = hs.get("pos") or {}
                for label, got, want in (("left", p.get("left"), eL),
                                         ("top", p.get("top"), eT),
                                         ("width", p.get("width"), eW)):
                    if got is None or abs(got - want) > TOL:
                        fails.append((tpl, f"geom {label}",
                                      f"got {got}, want ~{want:.1f}"))
        return list(CASES), fails
    finally:
        for p in (pres, carrier):
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
        print(f"FAIL: template run failed after retries: {last!r}")
        return 1

    tpls, fails = result
    print(f"templates: {len(tpls)} ({', '.join(tpls)})")
    if fails:
        for t, chk, why in fails:
            print(f"  FAIL [{t}] {chk}: {why}")
    else:
        print("  all templates: applied, slot text present, headline geometry in tolerance")
    ok = not fails
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

"""Smoke-test every chart type Decko claims to support.

Drives PowerPoint via COM, opens the carrier (PPT_AI_Editor.pptm) so its VBA
loads, creates a fresh test deck, and for each chart type adds a blank slide
then calls modExecuteInstructions.ExecuteFromString with a single add_chart
action. Reads the per-action result back from the deck's action log.

Bypasses the MSForms TextBox in frmExecute (which corrupts large pastes).

Usage: python tools/test_charts.py
"""
import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CARRIER = ROOT / "PPT_AI_Editor.pptm"
TESTDECK = Path.home() / "Downloads" / "decko_chart_test.pptx"
LOG = Path(str(TESTDECK) + ".action_log.jsonl")

TYPES = [
    "columnclustered", "columnstacked", "columnstackedpercent",
    "column3d", "columnclustered3d", "columnstacked3d",
    "barclustered", "barstacked", "barstackedpercent",
    "bar3d", "barclustered3d", "barstacked3d",
    "line", "linemarkers", "linestacked", "linestackedmarkers", "line3d",
    "area", "areastacked", "areapercent", "area3d", "areastacked3d",
    "pie", "pie3d", "pieexploded3d", "doughnut",
    "scatter", "radar", "radarmarkers", "radarfilled",
    "surface", "surfacewireframe",
    "waterfall", "pareto", "funnel", "histogram", "boxwhisker",
    "treemap", "sunburst",
]
SINGLE_SERIES = {"pie", "pie3d", "pieexploded3d", "doughnut", "funnel"}


def _log_lines():
    if not LOG.exists():
        return []
    return [l for l in LOG.read_text(encoding="utf-8", errors="replace").splitlines() if l.strip()]


def main() -> int:
    try:
        import win32com.client as win32
    except ImportError:
        print("ERROR: pywin32 not installed.  pip install pywin32")
        return 1

    if LOG.exists():
        LOG.unlink()

    app = win32.DispatchEx("PowerPoint.Application")
    app.Visible = True
    try:
        app.AutomationSecurity = 1  # msoAutomationSecurityLow -> run macros, no prompt
    except Exception as e:
        print(f"warn: AutomationSecurity not set: {e}")

    print(f"[open] {CARRIER}")
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)

    if TESTDECK.exists():
        try:
            TESTDECK.unlink()
        except Exception:
            pass
    deck = app.Presentations.Add()
    deck.SaveAs(str(TESTDECK))   # gives FullName -> known action-log path
    print(f"[deck] {TESTDECK}")

    # pick a blank-ish layout
    try:
        blank_layout = deck.SlideMaster.CustomLayouts(7)   # "Blank" in default theme
    except Exception:
        blank_layout = deck.SlideMaster.CustomLayouts(1)

    def run(payload: str) -> str:
        deck.Windows(1).Activate()
        try:
            return app.Run("ExecuteFromString", payload)
        except Exception:
            return app.Run("PPT_AI_Editor.pptm!ExecuteFromString", payload)

    results = []
    for i, t in enumerate(TYPES, 1):
        slide = deck.Slides.AddSlide(deck.Slides.Count + 1, blank_layout)
        sidx = slide.SlideIndex
        series = [{"name": "Alpha", "values": [10, 24, 17, 32]}]
        if t not in SINGLE_SERIES:
            series.append({"name": "Beta", "values": [14, 9, 21, 12]})
        cats = ["1", "2", "3", "4"] if t == "scatter" else ["Q1", "Q2", "Q3", "Q4"]
        action = {
            "type": "add_chart", "slide": sidx, "chart_type": t,
            "pos": {"left": 40, "top": 60, "width": 620, "height": 400},
            "categories": cats, "series": series,
            "title": f"{i:02d} {t}",
        }
        before = len(_log_lines())
        ret = ""
        try:
            ret = run(json.dumps({"actions": [action]}))
        except Exception as e:
            print(f"{i:02d} {t:22s} -> RUNFAIL {e}")
            results.append((t, "RUNFAIL", repr(e)))
            continue
        new = [json.loads(l) for l in _log_lines()[before:]]
        ch = [e for e in new if e.get("op") == "add_chart"]
        if ch:
            status, reason = ch[-1].get("status"), ch[-1].get("reason", "")
        else:
            status, reason = "NOLOG", str(ret)
        results.append((t, status, reason))
        print(f"{i:02d} {t:22s} -> {status:8s} {reason}")
        time.sleep(0.03)

    deck.Save()
    ok = [t for t, s, r in results if s == "ok"]
    bad = [(t, s, r) for t, s, r in results if s != "ok"]
    print("\n=== SUMMARY ===")
    print(f"ok {len(ok)}/{len(results)}")
    for t, s, r in bad:
        print(f"  FAIL {t}: {s} {r}")
    print(f"\ntest deck: {TESTDECK}\nlog: {LOG}")
    print("(PowerPoint left open for inspection)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

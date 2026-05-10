"""Validate P1 chart capability additions: waterfall, custom chart titles,
gradient fill on series, drop/hi-lo/up-down lines, and new chart types
(100%-stacked, radar, area).

Synthetic generic data — no real-world values.

Run: python tests/run_chart_p1_tests.py
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "p1_chart_tests.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "p1_chart_tests_pngs"

NAVY = "#15283C"
GRAY = "#B6BCC9"
RED = "#C73E1D"
GREEN = "#2D7A4D"
GOLD = "#C8A951"


def chart(slide, ref, x, y, w, h, chart_type, categories, series, **kw):
    return {
        "type": "add_chart", "slide": slide, "ref_name": ref,
        "chart_type": chart_type,
        "pos": {"left": x, "top": y, "width": w, "height": h},
        "categories": categories, "series": series,
        "show_legend": kw.get("show_legend", False),
        "show_values": kw.get("show_values", False),
        "clean_style": kw.get("clean_style", True),
        "value_format": kw.get("value_format", ""),
    }


def title(slide, x, y, txt):
    return {"type": "add_text_box", "slide": slide,
            "pos": {"left": x, "top": y, "width": 600, "height": 30},
            "text": txt, "font_color": NAVY, "font_size": 14, "font_bold": True}


def build_slide_1():
    """Waterfall chart - finance bridge analysis."""
    return [
        {"type": "clear_slide", "slide": 1},
        title(1, 30, 20, "Waterfall Chart (P1)"),
        chart(1, "ch_wf", 30, 70, 900, 350, "waterfall",
              ["Start", "Revenue", "Expenses", "Tax", "End"],
              [{"name": "Bridge", "values": [100, 50, -30, -10, 110]}]),
    ]


def build_slide_2():
    """Custom chart title font + position."""
    cats = ["Q1", "Q2", "Q3", "Q4"]
    return [
        {"type": "add_slide", "position": 2, "layout_index": 6},
        {"type": "clear_slide", "slide": 2},
        chart(2, "ch_title", 30, 60, 900, 380, "columnclustered", cats,
              [{"name": "Rev", "values": [100, 120, 140, 160], "color": NAVY}]),
        {"type": "set_chart_title", "slide": 2, "shape_name": "ch_title",
         "value": "Quarterly Revenue (USD M)", "enabled": True,
         "props": {
             "font_size": 16, "font_color": NAVY,
             "font_bold": True, "font_italic": True,
             "position": "above",
         }},
    ]


def build_slide_3():
    """Gradient fill on series."""
    cats = ["A", "B", "C", "D", "E"]
    return [
        {"type": "add_slide", "position": 3, "layout_index": 6},
        {"type": "clear_slide", "slide": 3},
        title(3, 30, 20, "Gradient Series Fill (P1)"),
        chart(3, "ch_grad_v", 30, 70, 440, 320, "columnclustered", cats,
              [{"name": "Vertical", "values": [30, 45, 60, 75, 90]}]),
        {"type": "set_chart_series", "slide": 3, "shape_name": "ch_grad_v",
         "series_index": 1, "props": {
             "gradient_fill": {"from": NAVY, "to": GOLD, "direction": "vertical"},
         }},
        title(3, 30, 400, "Vertical gradient: navy → gold"),
        chart(3, "ch_grad_h", 510, 70, 440, 320, "barclustered", cats,
              [{"name": "Horizontal", "values": [30, 45, 60, 75, 90]}]),
        {"type": "set_chart_series", "slide": 3, "shape_name": "ch_grad_h",
         "series_index": 1, "props": {
             "gradient_fill": {"from": GREEN, "to": GRAY, "direction": "horizontal"},
         }},
        title(3, 510, 400, "Horizontal gradient: green → gray"),
    ]


def build_slide_4():
    """Drop lines / hi-lo lines / up-down bars."""
    cats = ["1", "2", "3", "4", "5", "6", "7", "8"]
    return [
        {"type": "add_slide", "position": 4, "layout_index": 6},
        {"type": "clear_slide", "slide": 4},
        title(4, 30, 20, "Drop Lines / Hi-Lo / Up-Down Bars (P1)"),
        chart(4, "ch_drop", 30, 70, 290, 320, "linemarkers", cats,
              [{"name": "Series", "values": [22, 45, 18, 52, 30, 48, 35, 60],
                "color": NAVY}]),
        {"type": "set_chart_format", "slide": 4, "shape_name": "ch_drop",
         "props": {"drop_lines": True}},
        title(4, 30, 400, "Drop lines"),
        chart(4, "ch_hilo", 340, 70, 290, 320, "linemarkers", cats,
              [{"name": "High", "values": [50, 55, 48, 60, 52, 58, 62, 65],
                "color": GREEN},
               {"name": "Low", "values": [20, 22, 18, 25, 23, 28, 30, 32],
                "color": RED}]),
        {"type": "set_chart_format", "slide": 4, "shape_name": "ch_hilo",
         "props": {"hi_lo_lines": True}},
        title(4, 340, 400, "Hi-lo lines"),
        chart(4, "ch_ud", 650, 70, 290, 320, "linemarkers", cats,
              [{"name": "Open", "values": [30, 35, 28, 40, 32, 38, 42, 45],
                "color": NAVY},
               {"name": "Close", "values": [35, 32, 38, 35, 40, 42, 38, 50],
                "color": GRAY}]),
        {"type": "set_chart_format", "slide": 4, "shape_name": "ch_ud",
         "props": {"up_down_bars": True}},
        title(4, 650, 400, "Up-down bars"),
    ]


def build_slide_5():
    """New chart types — 100%-stacked, radar, area."""
    cats = ["Q1", "Q2", "Q3", "Q4"]
    series2 = [
        {"name": "A", "values": [30, 35, 40, 38], "color": NAVY},
        {"name": "B", "values": [25, 30, 28, 32], "color": GRAY},
        {"name": "C", "values": [15, 20, 18, 22], "color": GOLD},
    ]
    return [
        {"type": "add_slide", "position": 5, "layout_index": 6},
        {"type": "clear_slide", "slide": 5},
        title(5, 30, 20, "New Chart Types (P1)"),
        chart(5, "ch_pct", 30, 70, 290, 320, "column_100pct", cats, series2,
              show_legend=True),
        title(5, 30, 400, "100% stacked column"),
        chart(5, "ch_area", 340, 70, 290, 320, "area_stacked", cats, series2,
              show_legend=True),
        title(5, 340, 400, "Stacked area"),
        chart(5, "ch_radar", 650, 70, 290, 320, "radar_markers",
              ["Speed", "Power", "Range", "Stealth", "Cost"],
              [{"name": "Model A", "values": [80, 65, 90, 45, 70], "color": NAVY}]),
        title(5, 650, 400, "Radar with markers"),
    ]


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_p1_"))
    try:
        try: app.Visible = True
        except Exception: pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "p1.pptx"))
        deck.Windows(1).Activate()

        all_actions = (build_slide_1() + build_slide_2() + build_slide_3() +
                       build_slide_4() + build_slide_5())
        result = app.Run("PPT_AI_Editor!ExecuteFromString",
                         json.dumps({"actions": all_actions}))
        print(f"Build: {result}")

        if "skipped" in result and " 0 skipped" not in result:
            log = deck.FullName + ".action_log.jsonl"
            try:
                with open(log) as f:
                    lines = f.readlines()
                for ln in lines[-len(all_actions):]:
                    e = json.loads(ln)
                    if e.get("status") in ("skipped", "error"):
                        print(f"  !! {e['status']}: {e.get('op')} -> {e.get('reason','')}")
            except Exception as ex:
                print(f"  log read failed: {ex}")

        OUT_DECK.parent.mkdir(parents=True, exist_ok=True)
        deck.SaveAs(str(OUT_DECK))
        PNG_DIR.mkdir(parents=True, exist_ok=True)
        for sl in deck.Slides:
            sl.Export(str(PNG_DIR / f"slide_{sl.SlideNumber:02d}.png"), "PNG", 1280, 720)
        print(f"Saved deck: {OUT_DECK}")

    finally:
        try:
            for p in list(app.Presentations):
                try: p.Saved = True; p.Close()
                except Exception: pass
        except Exception: pass
        try: app.Quit()
        except Exception: pass
        time.sleep(0.5)
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main() or 0)

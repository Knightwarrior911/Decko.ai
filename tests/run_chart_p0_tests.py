"""Validate P0 chart capability additions: gap_width/overlap, trendlines,
error bars, per-point label positions, reverse plot order, log scale,
doughnut hole size, bar shapes.

Synthetic generic data — no real-world values.

Run: python tests/run_chart_p0_tests.py
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "p0_chart_tests.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "p0_chart_tests_pngs"

NAVY = "#15283C"
GRAY = "#B6BCC9"
RED = "#C73E1D"
GREEN = "#2D7A4D"


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
    """gap_width / overlap test — banker-tight bars vs default."""
    cats = ["Q1", "Q2", "Q3", "Q4"]
    a = [
        {"type": "clear_slide", "slide": 1},
        title(1, 30, 20, "Gap Width and Overlap (P0)"),
        # Default chart (left)
        chart(1, "ch_default", 30, 70, 440, 240, "columnclustered", cats,
              [{"name": "A", "values": [10, 12, 14, 16], "color": NAVY},
               {"name": "B", "values": [8, 9, 11, 13], "color": GRAY}],
              show_legend=True),
        # Same chart with gap_width=30 + overlap=-10 (bars overlap slightly, banker style)
        chart(1, "ch_tight", 510, 70, 440, 240, "columnclustered", cats,
              [{"name": "A", "values": [10, 12, 14, 16], "color": NAVY},
               {"name": "B", "values": [8, 9, 11, 13], "color": GRAY}],
              show_legend=True),
        {"type": "set_chart_format", "slide": 1, "shape_name": "ch_tight",
         "props": {"gap_width": 30, "overlap": -10}},
        title(1, 30, 320, "Default (gap=150, overlap=0)"),
        title(1, 510, 320, "Tight (gap=30, overlap=-10) — banker style"),
    ]
    return a


def build_slide_2():
    """Trendline test — linear, polynomial, moving avg."""
    a = [
        {"type": "add_slide", "position": 2, "layout_index": 6},
        {"type": "clear_slide", "slide": 2},
        title(2, 30, 20, "Trendlines (P0)"),
        chart(2, "ch_lin", 30, 70, 290, 240, "columnclustered",
              ["2020", "2021", "2022", "2023", "2024", "2025"],
              [{"name": "Revenue",
                "values": [100, 115, 128, 140, 158, 175], "color": NAVY}]),
        {"type": "add_chart_trendline", "slide": 2, "shape_name": "ch_lin",
         "series_index": 1, "props": {
             "type": "linear", "forward": 2,
             "color": RED, "weight": 2, "dash": "dash",
             "display_equation": True, "display_r_squared": True,
         }},
        title(2, 30, 320, "Linear + 2-period forecast"),
        chart(2, "ch_poly", 340, 70, 290, 240, "columnclustered",
              ["Y1", "Y2", "Y3", "Y4", "Y5", "Y6"],
              [{"name": "Metric",
                "values": [50, 80, 95, 90, 70, 40], "color": NAVY}]),
        {"type": "add_chart_trendline", "slide": 2, "shape_name": "ch_poly",
         "series_index": 1, "props": {
             "type": "polynomial", "order": 3,
             "color": RED, "weight": 2,
         }},
        title(2, 340, 320, "Polynomial order=3"),
        chart(2, "ch_ma", 650, 70, 290, 240, "linemarkers",
              ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],
              [{"name": "Daily",
                "values": [22, 35, 18, 42, 28, 50, 33, 45, 38, 52], "color": GRAY}]),
        {"type": "add_chart_trendline", "slide": 2, "shape_name": "ch_ma",
         "series_index": 1, "props": {
             "type": "moving_avg", "period": 3,
             "color": NAVY, "weight": 2,
         }},
        title(2, 650, 320, "Moving avg period=3"),
    ]
    return a


def build_slide_3():
    """Error bars test."""
    a = [
        {"type": "add_slide", "position": 3, "layout_index": 6},
        {"type": "clear_slide", "slide": 3},
        title(3, 30, 20, "Error Bars (P0)"),
        chart(3, "ch_err", 30, 70, 440, 280, "columnclustered",
              ["A", "B", "C", "D"],
              [{"name": "Mean", "values": [50, 65, 58, 72], "color": NAVY}]),
        {"type": "set_chart_error_bars", "slide": 3, "shape_name": "ch_err",
         "series_index": 1, "props": {
             "direction": "y", "include": "both", "type": "fixed", "amount": 5,
             "end_style": "cap", "color": RED, "weight": 1.5,
         }},
        title(3, 30, 360, "Fixed value ±5"),
        chart(3, "ch_err_pct", 510, 70, 440, 280, "columnclustered",
              ["A", "B", "C", "D"],
              [{"name": "Mean", "values": [50, 65, 58, 72], "color": NAVY}]),
        {"type": "set_chart_error_bars", "slide": 3, "shape_name": "ch_err_pct",
         "series_index": 1, "props": {
             "direction": "y", "include": "plus", "type": "percent", "amount": 10,
             "end_style": "cap", "color": GREEN, "weight": 1.5,
         }},
        title(3, 510, 360, "Percent +10% (plus only)"),
    ]
    return a


def build_slide_4():
    """Per-point label position — pos bars above, neg below."""
    a = [
        {"type": "add_slide", "position": 4, "layout_index": 6},
        {"type": "clear_slide", "slide": 4},
        title(4, 30, 20, "Per-Point Label Positions (P0)"),
        chart(4, "ch_pp", 30, 70, 900, 280, "columnclustered",
              ["A", "B", "C", "D", "E", "F"],
              [{"name": "Δ%", "values": [12, -8, 15, -5, 22, -18], "color": NAVY}]),
        {"type": "set_chart_series", "slide": 4, "shape_name": "ch_pp",
         "series_index": 1, "props": {
             "show_labels": True,
             "custom_labels": ["+12%", "-8%", "+15%", "-5%", "+22%", "-18%"],
             "label_size": 11, "label_bold": True,
             # Positive above, negative below — per-point override
             "point_label_positions": ["above", "below", "above", "below", "above", "below"],
             "point_fills": [NAVY, RED, NAVY, RED, NAVY, RED],
             "point_label_colors": [NAVY, RED, NAVY, RED, NAVY, RED],
         }},
        {"type": "set_chart_axis", "slide": 4, "shape_name": "ch_pp",
         "axis": "x", "props": {"tick_label_position": "low"}},
    ]
    return a


def build_slide_5():
    """Reverse plot order + log scale + doughnut hole size."""
    a = [
        {"type": "add_slide", "position": 5, "layout_index": 6},
        {"type": "clear_slide", "slide": 5},
        title(5, 30, 20, "Reverse Order, Log Scale, Doughnut Hole (P0)"),
        # Reversed categories — most recent on left
        chart(5, "ch_rev", 30, 70, 290, 240, "columnclustered",
              ["2020", "2021", "2022", "2023", "2024", "2025"],
              [{"name": "Rev", "values": [100, 115, 128, 140, 158, 175], "color": NAVY}]),
        {"type": "set_chart_format", "slide": 5, "shape_name": "ch_rev",
         "props": {"reverse_categories": True}},
        title(5, 30, 320, "Reverse categories"),
        # Log scale axis
        chart(5, "ch_log", 340, 70, 290, 240, "columnclustered",
              ["A", "B", "C", "D"],
              [{"name": "Range", "values": [10, 100, 1000, 10000], "color": NAVY}]),
        {"type": "set_chart_format", "slide": 5, "shape_name": "ch_log",
         "props": {"scale_type": "logarithmic"}},
        title(5, 340, 320, "Log scale"),
        # Doughnut with smaller hole
        chart(5, "ch_dn", 650, 70, 290, 240, "doughnut",
              ["X", "Y", "Z", "W"],
              [{"name": "Share", "values": [30, 25, 20, 25],
                "color": NAVY}]),
        {"type": "set_chart_format", "slide": 5, "shape_name": "ch_dn",
         "props": {"doughnut_hole_size": 35}},
        title(5, 650, 320, "Doughnut hole size=35%"),
    ]
    return a


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_p0_"))
    try:
        try: app.Visible = True
        except Exception: pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "p0.pptx"))
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

"""Validate P2 chart additions: axis label rotation, plot/chart area styling,
axis title font, pie leader lines, pattern fill, series border control.

Synthetic generic data — no real-world values.

Run: python tests/run_chart_p2_tests.py
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "p2_chart_tests.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "p2_chart_tests_pngs"

NAVY = "#15283C"
GRAY = "#B6BCC9"
RED = "#C73E1D"
GREEN = "#2D7A4D"
GOLD = "#C8A951"
PALE = "#F1F2F4"


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
    """Axis label rotation — 45 deg + 90 deg + tick mark styles."""
    cats = ["Quarter 1 2024", "Quarter 2 2024", "Quarter 3 2024", "Quarter 4 2024"]
    return [
        {"type": "clear_slide", "slide": 1},
        title(1, 30, 20, "Axis Label Rotation + Tick Marks (P2)"),
        chart(1, "ch_45", 30, 70, 290, 320, "columnclustered", cats,
              [{"name": "Rev", "values": [100, 120, 140, 160], "color": NAVY}],
              clean_style=False),
        {"type": "set_chart_axis", "slide": 1, "shape_name": "ch_45",
         "axis": "x", "props": {"label_rotation": -45, "major_tick_mark": "outside"}},
        title(1, 30, 400, "Tilt -45° + outside ticks"),
        chart(1, "ch_90", 340, 70, 290, 320, "columnclustered", cats,
              [{"name": "Rev", "values": [100, 120, 140, 160], "color": GREEN}],
              clean_style=False),
        {"type": "set_chart_axis", "slide": 1, "shape_name": "ch_90",
         "axis": "x", "props": {"label_rotation": -90, "major_tick_mark": "cross"}},
        title(1, 340, 400, "Tilt -90° + cross ticks"),
        chart(1, "ch_no", 650, 70, 290, 320, "columnclustered", cats,
              [{"name": "Rev", "values": [100, 120, 140, 160], "color": GOLD}],
              clean_style=False),
        {"type": "set_chart_axis", "slide": 1, "shape_name": "ch_no",
         "axis": "x", "props": {"label_rotation": 0, "major_tick_mark": "none"}},
        title(1, 650, 400, "No rotation, no ticks"),
    ]


def build_slide_2():
    """Plot/chart area styling — background fills + borders."""
    cats = ["A", "B", "C", "D"]
    return [
        {"type": "add_slide", "position": 2, "layout_index": 6},
        {"type": "clear_slide", "slide": 2},
        title(2, 30, 20, "Chart/Plot Area Fill + Border (P2)"),
        chart(2, "ch_pa", 30, 70, 440, 320, "columnclustered", cats,
              [{"name": "X", "values": [50, 65, 58, 72], "color": NAVY}],
              clean_style=False),
        {"type": "set_chart_format", "slide": 2, "shape_name": "ch_pa",
         "props": {
             "chart_area_fill": "#F8F8F8",
             "chart_area_border": NAVY,
             "plot_area_fill": "#FFFFFF",
             "plot_area_border": GRAY,
         }},
        title(2, 30, 400, "Light gray chart area + white plot area"),
        chart(2, "ch_dark", 510, 70, 440, 320, "columnclustered", cats,
              [{"name": "X", "values": [50, 65, 58, 72], "color": GOLD}],
              clean_style=False),
        {"type": "set_chart_format", "slide": 2, "shape_name": "ch_dark",
         "props": {
             "chart_area_fill": NAVY,
             "plot_area_fill": "#1F3349",
         }},
        {"type": "set_chart_axis", "slide": 2, "shape_name": "ch_dark",
         "axis": "x", "props": {"label_color": "#FFFFFF"}},
        {"type": "set_chart_axis", "slide": 2, "shape_name": "ch_dark",
         "axis": "y", "props": {"label_color": "#FFFFFF"}},
        title(2, 510, 400, "Dark mode (navy chart area)"),
    ]


def build_slide_3():
    """Pie + doughnut leader lines + custom title."""
    cats = ["Engineering", "Sales", "Marketing", "Operations", "Finance"]
    vals = [{"name": "Headcount %", "values": [40, 25, 12, 18, 5]}]
    return [
        {"type": "add_slide", "position": 3, "layout_index": 6},
        {"type": "clear_slide", "slide": 3},
        title(3, 30, 20, "Pie Leader Lines + Custom Title (P2)"),
        chart(3, "ch_pie", 30, 70, 440, 380, "pie", cats, vals,
              show_legend=True),
        {"type": "set_chart_title", "slide": 3, "shape_name": "ch_pie",
         "value": "Department Headcount Mix", "enabled": True,
         "props": {"font_size": 14, "font_color": NAVY, "font_bold": True}},
        {"type": "set_chart_series", "slide": 3, "shape_name": "ch_pie",
         "series_index": 1, "props": {
             "show_labels": True, "label_position": "outside_end",
             "show_leader_lines": True, "leader_line_color": GRAY,
             "label_size": 10,
         }},
        chart(3, "ch_dn", 510, 70, 440, 380, "doughnut", cats, vals,
              show_legend=True),
        {"type": "set_chart_format", "slide": 3, "shape_name": "ch_dn",
         "props": {"doughnut_hole_size": 50}},
        {"type": "set_chart_series", "slide": 3, "shape_name": "ch_dn",
         "series_index": 1, "props": {
             "show_labels": True, "label_position": "outside_end",
             "show_leader_lines": True, "leader_line_color": NAVY,
             "label_size": 10,
         }},
    ]


def build_slide_4():
    """Pattern fill on series + axis title with font."""
    cats = ["A", "B", "C", "D"]
    return [
        {"type": "add_slide", "position": 4, "layout_index": 6},
        {"type": "clear_slide", "slide": 4},
        title(4, 30, 20, "Pattern Fill + Axis Title Font (P2)"),
        chart(4, "ch_pat", 30, 70, 440, 320, "columnclustered", cats,
              [{"name": "Region 1", "values": [50, 65, 58, 72]},
               {"name": "Region 2", "values": [30, 45, 38, 52]}],
              show_legend=True, clean_style=False),
        {"type": "set_chart_series", "slide": 4, "shape_name": "ch_pat",
         "series_index": 1, "props": {
             "pattern_fill": {"type": "diagonal_brick", "fore": NAVY, "back": "#FFFFFF"},
         }},
        {"type": "set_chart_series", "slide": 4, "shape_name": "ch_pat",
         "series_index": 2, "props": {
             "pattern_fill": {"type": "small_checker", "fore": GREEN, "back": "#FFFFFF"},
         }},
        {"type": "set_chart_axis", "slide": 4, "shape_name": "ch_pat",
         "axis": "y", "props": {
             "title": "Revenue (USD M)",
             "title_size": 11, "title_color": NAVY, "title_bold": True,
         }},
        title(4, 30, 400, "Two patterns: diagonal brick + small checker"),
        chart(4, "ch_axtitle", 510, 70, 440, 320, "columnclustered", cats,
              [{"name": "Y", "values": [10, 25, 18, 30], "color": GOLD}],
              clean_style=False),
        {"type": "set_chart_axis", "slide": 4, "shape_name": "ch_axtitle",
         "axis": "x", "props": {
             "title": "Quarters",
             "title_size": 12, "title_color": GREEN, "title_italic": True,
         }},
        {"type": "set_chart_axis", "slide": 4, "shape_name": "ch_axtitle",
         "axis": "y", "props": {
             "title": "Index", "title_size": 12, "title_color": RED,
         }},
        title(4, 510, 400, "Both axis titles styled"),
    ]


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_p2_"))
    try:
        try: app.Visible = True
        except Exception: pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "p2.pptx"))
        deck.Windows(1).Activate()

        all_actions = (build_slide_1() + build_slide_2() +
                       build_slide_3() + build_slide_4())
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

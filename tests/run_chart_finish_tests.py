"""Finish remaining items: stacked-bar+line combo, chart background image, callouts."""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "p_finish_chart_tests.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "p_finish_chart_tests_pngs"

NAVY = "#15283C"
GRAY = "#B6BCC9"
GOLD = "#C8A951"
RED = "#C73E1D"


def chart(slide, ref, x, y, w, h, chart_type, categories, series, **kw):
    return {
        "type": "add_chart", "slide": slide, "ref_name": ref,
        "chart_type": chart_type,
        "pos": {"left": x, "top": y, "width": w, "height": h},
        "categories": categories, "series": series,
        "show_legend": kw.get("show_legend", True),
        "show_values": kw.get("show_values", False),
        "clean_style": kw.get("clean_style", False),
        "value_format": kw.get("value_format", ""),
    }


def title(slide, x, y, txt):
    return {"type": "add_text_box", "slide": slide,
            "pos": {"left": x, "top": y, "width": 600, "height": 30},
            "text": txt, "font_color": NAVY, "font_size": 14, "font_bold": True}


def build_slide_1():
    """Stacked column + line combo (item #24)."""
    cats = ["2022", "2023", "2024", "2025"]
    return [
        {"type": "clear_slide", "slide": 1},
        title(1, 30, 20, "Stacked Column + Line Combo (item 24)"),
        chart(1, "ch_combo", 30, 70, 900, 380, "columnstacked", cats,
              [{"name": "Product A", "values": [10, 12, 14, 16], "color": NAVY},
               {"name": "Product B", "values": [6, 8, 10, 12], "color": GRAY},
               {"name": "Margin %", "values": [25, 28, 32, 35], "color": GOLD}],
              show_legend=True, clean_style=False),
        # Move series 3 to line on secondary axis
        {"type": "set_chart_series", "slide": 1, "shape_name": "ch_combo",
         "series_index": 3, "props": {
             "chart_type": "linemarkers",
             "axis_group": "secondary",
             "line_color": GOLD,
             "line_weight": 2.5,
             "marker_style": "circle",
             "marker_size": 7,
             "marker_fill": GOLD,
             "marker_line": GOLD,
             "show_labels": True,
             "label_position": "above",
             "custom_labels": ["25%", "28%", "32%", "35%"],
             "label_color": GOLD,
             "label_bold": True,
             "label_size": 10,
         }},
    ]


def build_slide_2():
    """Outside-chart callouts with connectors (item #15)."""
    cats = ["Q1", "Q2", "Q3", "Q4"]
    return [
        {"type": "add_slide", "position": 2, "layout_index": 6},
        {"type": "clear_slide", "slide": 2},
        title(2, 30, 20, "Outside-Chart Callouts with Connectors (item 15)"),
        chart(2, "ch_call", 30, 70, 600, 380, "columnclustered", cats,
              [{"name": "Rev", "values": [50, 65, 58, 72], "color": NAVY}],
              clean_style=True),
        # Outside callout box
        {"type": "add_shape", "slide": 2, "kind": "rrect", "ref_name": "callout1",
         "pos": {"left": 660, "top": 80, "width": 280, "height": 70},
         "fill": "#FFFFFF", "stroke": NAVY, "stroke_weight_pt": 1.2,
         "text": "Q1 dip due to seasonal\nmaintenance window",
         "font_color": NAVY, "font_size": 11},
        # Connector from chart bar 1 to callout — manual line shape
        {"type": "add_line", "slide": 2,
         "x1": 145, "y1": 280, "x2": 660, "y2": 115,
         "color": NAVY, "weight_pt": 1.0, "dash_style": "dash",
         "arrow_end": "filled"},
        {"type": "add_shape", "slide": 2, "kind": "rrect", "ref_name": "callout2",
         "pos": {"left": 660, "top": 250, "width": 280, "height": 70},
         "fill": "#FFFFFF", "stroke": GOLD, "stroke_weight_pt": 1.2,
         "text": "Q4 peak — record high\non year-end push",
         "font_color": GOLD, "font_size": 11, "font_bold": True},
        {"type": "add_line", "slide": 2,
         "x1": 535, "y1": 200, "x2": 660, "y2": 285,
         "color": GOLD, "weight_pt": 1.0, "dash_style": "dash",
         "arrow_end": "filled"},
    ]


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_finish_"))
    try:
        try: app.Visible = True
        except Exception: pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "finish.pptx"))
        deck.Windows(1).Activate()

        actions = build_slide_1() + build_slide_2()
        result = app.Run("PPT_AI_Editor!ExecuteFromString",
                         json.dumps({"actions": actions}))
        print(f"Build: {result}")

        if "skipped" in result and " 0 skipped" not in result:
            log = deck.FullName + ".action_log.jsonl"
            try:
                with open(log) as f:
                    for ln in f.readlines()[-len(actions):]:
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

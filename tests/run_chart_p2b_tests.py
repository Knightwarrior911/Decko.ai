"""P2 follow-up: 3D chart types + rotation/elevation/perspective.

Synthetic generic data — no real-world values.

Run: python tests/run_chart_p2b_tests.py
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "p2b_chart_tests.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "p2b_chart_tests_pngs"

NAVY = "#15283C"
GRAY = "#B6BCC9"
GOLD = "#C8A951"


def chart(slide, ref, x, y, w, h, chart_type, categories, series, **kw):
    return {
        "type": "add_chart", "slide": slide, "ref_name": ref,
        "chart_type": chart_type,
        "pos": {"left": x, "top": y, "width": w, "height": h},
        "categories": categories, "series": series,
        "show_legend": kw.get("show_legend", False),
        "show_values": kw.get("show_values", False),
        "clean_style": kw.get("clean_style", False),
        "value_format": kw.get("value_format", ""),
    }


def title(slide, x, y, txt):
    return {"type": "add_text_box", "slide": slide,
            "pos": {"left": x, "top": y, "width": 600, "height": 30},
            "text": txt, "font_color": NAVY, "font_size": 14, "font_bold": True}


def build_slide():
    cats = ["Q1", "Q2", "Q3", "Q4"]
    series = [
        {"name": "A", "values": [50, 65, 58, 72], "color": NAVY},
        {"name": "B", "values": [40, 50, 55, 60], "color": GOLD},
    ]
    return [
        {"type": "clear_slide", "slide": 1},
        title(1, 30, 20, "3D Charts + Rotation/Elevation (P2b)"),
        chart(1, "ch_3d_def", 30, 70, 290, 320, "column_clustered_3d",
              cats, series, show_legend=True),
        title(1, 30, 400, "Default 3D"),
        chart(1, "ch_3d_rot", 340, 70, 290, 320, "column_clustered_3d",
              cats, series, show_legend=True),
        {"type": "set_chart_format", "slide": 1, "shape_name": "ch_3d_rot",
         "props": {"rotation": 30, "elevation": 25, "perspective": 30,
                   "height_percent": 80, "gap_depth": 50}},
        title(1, 340, 400, "Rotation 30°, Elevation 25°"),
        chart(1, "ch_pie3d", 650, 70, 290, 320, "pie3d",
              cats, [{"name": "Mix", "values": [30, 25, 20, 25]}],
              show_legend=True),
        {"type": "set_chart_format", "slide": 1, "shape_name": "ch_pie3d",
         "props": {"elevation": 35, "rotation": 60}},
        title(1, 650, 400, "3D Pie tilted"),
    ]


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_p2b_"))
    try:
        try: app.Visible = True
        except Exception: pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "p2b.pptx"))
        deck.Windows(1).Activate()

        actions = build_slide()
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
        deck.Slides(1).Export(str(PNG_DIR / "slide_01.png"), "PNG", 1280, 720)
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

"""Quick test: skip fetch, use existing image folder for picker + grid table.

Verifies multi-slide picker (>=24 imgs spans slides) and contain-fit images.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "picker_only_test.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "picker_only_test_pngs"

# Existing folder from a prior fetch run.
SRC_FOLDER = sys.argv[1] if len(sys.argv) > 1 else \
    r"C:\Users\vinit\AppData\Local\Temp\pptai_web_4frr7e30\assets\page_httpswww.apple.comiphone_20260510_214927"


def main() -> int:
    import win32com.client

    if not Path(SRC_FOLDER).exists():
        print(f"[FAIL] folder missing: {SRC_FOLDER}")
        return 1

    images = sorted(p for p in Path(SRC_FOLDER).iterdir()
                    if p.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif",
                                              ".bmp", ".webp"})
    print(f"[setup] {len(images)} images in folder")

    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_picker_"))
    try:
        try:
            app.Visible = True
        except Exception:
            pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "p.pptx"))
        deck.Windows(1).Activate()

        # Phase 1: multi-slide picker
        result1 = app.Run(
            "PPT_AI_Editor!ExecuteFromString",
            json.dumps({
                "actions": [
                    {"type": "build_image_picker_slide",
                     "folder": SRC_FOLDER,
                     "cols": 4,
                     "max_per_slide": 16},
                ]
            }),
        )
        print(f"[picker] {result1}")

        # Phase 2: grid table with sample 4 rows (contain default)
        sample = images[:4]
        rows = [{
            "name": p.stem,
            "image_path": str(p),
            "bullets": [f"Bullet 1 for {p.stem}",
                        f"Bullet 2 for {p.stem}",
                        "Lorem ipsum"],
        } for p in sample]

        target = deck.Slides.Count + 1
        result2 = app.Run(
            "PPT_AI_Editor!ExecuteFromString",
            json.dumps({
                "actions": [
                    {"type": "add_slide", "position": target, "layout_index": 6},
                    {"type": "build_image_grid_table",
                     "slide": target,
                     "ref_name": "apps_tbl",
                     "pos": {"left": 30, "top": 60, "width": 900, "height": 480},
                     "image_col": 1, "desc_col": 2,
                     "name_position": "bottom",
                     "name_strip_pt": 30, "image_pad_pt": 6,
                     "col1_width_pt": 280, "col2_width_pt": 620,
                     "name_font": {"size": 12, "bold": True, "color": "#15283C"},
                     "desc_font": {"size": 10, "color": "#333333"},
                     "rows": rows},
                ]
            }),
        )
        print(f"[grid] {result2}")

        OUT_DECK.parent.mkdir(parents=True, exist_ok=True)
        deck.SaveAs(str(OUT_DECK))
        PNG_DIR.mkdir(parents=True, exist_ok=True)
        for sl in deck.Slides:
            sl.Export(
                str(PNG_DIR / f"slide_{sl.SlideNumber:02d}.png"),
                "PNG", 1280, 720,
            )
        print(f"Saved deck: {OUT_DECK}")

    finally:
        try:
            for p in list(app.Presentations):
                try:
                    p.Saved = True
                    p.Close()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(0.5)
        shutil.rmtree(tmpdir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)

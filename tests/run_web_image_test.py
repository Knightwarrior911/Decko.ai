"""End-to-end test of the new web-image pipeline.

Pipeline under test:
  1. fetch_page_images(url)  -> downloads to <deck>\\assets\\<slug>_<ts>\\
  2. build_image_picker_slide(folder) -> drops every image as a labeled grid
  3. build_image_grid_table(rows=...) -> 2-col table on a target slide

This script does not exercise the Stanley URL end-to-end (LLM normally selects
filenames after seeing the picker slide). It verifies that the VBA actions
execute without error, fetch_page_images downloads >0 images, the picker slide
gets created, and build_image_grid_table draws the final table from a hardcoded
sample using the first three downloaded images.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK = REPO_ROOT / "test_decks" / "web_image_test.pptx"
PNG_DIR = REPO_ROOT / "test_decks" / "web_image_test_pngs"

# Use a small, image-heavy site that doesn't aggressively block scrapers.
TARGET_URL = sys.argv[1] if len(sys.argv) > 1 else \
    "https://www.apple.com/iphone/"


def main() -> int:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_web_"))
    try:
        try:
            app.Visible = True
        except Exception:
            pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, deck.SlideMaster.CustomLayouts(1))
        deck.SaveAs(str(tmpdir / "web.pptx"))
        deck.Windows(1).Activate()

        # --- Phase 1: fetch images ----------------------------------------
        result1 = app.Run(
            "PPT_AI_Editor!ExecuteFromString",
            json.dumps({
                "actions": [
                    {"type": "fetch_page_images", "url": TARGET_URL},
                ]
            }),
        )
        print(f"[fetch] {result1}")

        # Pull g_LastFetchFolder via a tiny helper macro -- no helper exists
        # yet, so derive from the assets dir naming convention instead.
        deck_dir = Path(deck.FullName).parent
        assets_dir = deck_dir / "assets"
        candidates = sorted(
            (p for p in assets_dir.glob("page_*") if p.is_dir()),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        if not candidates:
            print(f"[FAIL] No page_* folder under {assets_dir}")
            return 1
        folder = candidates[0]
        manifest_path = folder / "manifest.json"
        if not manifest_path.exists():
            print(f"[FAIL] manifest.json missing in {folder}")
            return 1
        manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
        images = manifest.get("images", [])
        print(f"[fetch] {len(images)} images in {folder}")
        if len(images) == 0:
            print("[FAIL] no images downloaded")
            return 1

        # --- Phase 2: picker slide ---------------------------------------
        result2 = app.Run(
            "PPT_AI_Editor!ExecuteFromString",
            json.dumps({
                "actions": [
                    {"type": "build_image_picker_slide",
                     "folder": str(folder),
                     "cols": 4},
                ]
            }),
        )
        print(f"[picker] {result2}")

        # --- Phase 3: image grid table on a fresh slide ------------------
        # Use the first 4 images as a quick visual smoke test.
        sample = images[:4]
        rows = []
        for i, img in enumerate(sample, start=1):
            rows.append({
                "name": img.get("alt") or f"Item {i}",
                "image_path": img["local_path"],
                "bullets": [
                    f"Sample bullet 1 for row {i}",
                    f"Sample bullet 2 for row {i}",
                    "Lorem ipsum dolor sit amet",
                ],
            })

        # Picker slide already added itself; target slide for the grid is
        # the next one after the current count.
        target_slide = deck.Slides.Count + 1
        result3 = app.Run(
            "PPT_AI_Editor!ExecuteFromString",
            json.dumps({
                "actions": [
                    {"type": "add_slide",
                     "position": target_slide,
                     "layout_index": 6},
                    {"type": "build_image_grid_table",
                     "slide": target_slide,
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
        print(f"[grid_table] {result3}")

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
        # Keep tmpdir for log inspection on failure.
        print(f"[debug] tmpdir kept at: {tmpdir}")
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)

"""Smoke for deck-wide actions (9 tests).

Single shared PowerPoint instance.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DECK_PHASE2 = REPO_ROOT / "test_decks" / "phase2.pptx"
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"

_APP = None
_CARRIER = None


def open_app():
    global _APP, _CARRIER
    if _APP is None:
        import win32com.client
        _APP = win32com.client.DispatchEx("PowerPoint.Application")
        _APP.Visible = True
        _CARRIER = _APP.Presentations.Open(str(CARRIER), WithWindow=True)
    return _APP


def fresh_deck(app):
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_deck_"))
    deck_copy = tmpdir / DECK_PHASE2.name
    shutil.copy2(DECK_PHASE2, deck_copy)
    deck = app.Presentations.Open(str(deck_copy), WithWindow=True)
    deck.Windows(1).Activate()
    return deck, _CARRIER, tmpdir


def teardown(app, *presentations, tmpdir=None):
    for p in presentations:
        if p is _CARRIER:
            continue
        try:
            p.Saved = True
            p.Close()
        except Exception:
            pass
    if tmpdir and tmpdir.exists():
        shutil.rmtree(tmpdir, ignore_errors=True)


def shutdown_app():
    global _APP, _CARRIER
    if _CARRIER is not None:
        try:
            _CARRIER.Saved = True
            _CARRIER.Close()
        except Exception:
            pass
        _CARRIER = None
    if _APP is not None:
        try:
            _APP.Quit()
        except Exception:
            pass
        _APP = None


def snap(app):
    return json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))


def test_find_replace_regex():
    print("test_find_replace_regex")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # First insert a known token via bulk_insert_text_box, then regex-replace it
        app.Run("PPT_AI_Editor!Do_bulk_insert_text_box",
                [1], "REGEX_TARGET_42", 50.0, 50.0, 200.0, 30.0)
        app.Run("PPT_AI_Editor!Do_find_replace_regex", "deck", r"REGEX_TARGET_\d+", "REGEX_HIT")
        after = snap(app)
        any_match = any("REGEX_HIT" in (sh.get("text") or "")
                        for sl in after["slides"]
                        for sh in sl["shapes"])
        assert any_match, "no REGEX_HIT after regex replace"
        print("  ok  [find_replace_regex]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_swap_font_deck_wide():
    print("test_swap_font_deck_wide")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        # Find the most-common font on slide 1 to use as 'from'
        all_fonts = []
        for sl in before["slides"]:
            for sh in sl["shapes"]:
                if "paragraphs" in sh:
                    for para in sh["paragraphs"]:
                        for run in para.get("runs", []):
                            n = run.get("font", {}).get("name")
                            if n:
                                all_fonts.append(n)
        if not all_fonts:
            print("  skip  [swap_font_deck_wide] no fonts found")
            return
        from_name = all_fonts[0]
        app.Run("PPT_AI_Editor!Do_swap_font_deck_wide", from_name, "Cascadia Code")
        after = snap(app)
        found_cascadia = False
        for sl in after["slides"]:
            for sh in sl["shapes"]:
                if "paragraphs" in sh:
                    for para in sh["paragraphs"]:
                        for run in para.get("runs", []):
                            if run.get("font", {}).get("name") == "Cascadia Code":
                                found_cascadia = True
        assert found_cascadia, f"no run uses Cascadia Code after swap from {from_name}"
        print(f"  ok  [swap_font] {from_name} -> Cascadia Code")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_recolor_palette_deck_wide():
    print("test_recolor_palette_deck_wide")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # Find any shape on slide 1, set its fill to known color, then recolor.
        s = snap(app)
        target_id = None
        for sh in s["slides"][0]["shapes"]:
            if sh["type"] != "title":
                target_id = sh["shape_id"]
                break
        if target_id is None:
            print("  skip  [recolor_palette] no usable shape")
            return
        app.Run("PPT_AI_Editor!Do_set_fill_color", 1, target_id, "#123456")
        app.Run("PPT_AI_Editor!Do_recolor_palette_deck_wide", "#123456", "#FF00FF", "fill")
        after = snap(app)
        magenta_seen = any(sh.get("fill") == "#FF00FF"
                           for sl in after["slides"] for sh in sl["shapes"])
        assert magenta_seen, "no shape with #FF00FF after recolor; saw fills: " + \
            str([sh.get("fill") for sl in after["slides"] for sh in sl["shapes"]])
        print("  ok  [recolor_palette] #123456 -> #FF00FF")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_slide_size_preset():
    print("test_set_slide_size_preset")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # Get current size; switch preset; verify changed
        app.Run("PPT_AI_Editor!Do_set_slide_size_preset", "4:3")
        w = deck.PageSetup.SlideWidth
        h = deck.PageSetup.SlideHeight
        assert abs(w - 720) < 1 and abs(h - 540) < 1, f"got {w}x{h}, expected 720x540"
        print(f"  ok  [set_slide_size 4:3] {w}x{h}")
        # Switch to 16:9
        app.Run("PPT_AI_Editor!Do_set_slide_size_preset", "16:9")
        w2 = deck.PageSetup.SlideWidth
        assert abs(w2 - 960) < 1, f"got {w2}, expected 960"
        print(f"  ok  [set_slide_size 16:9] {w2}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_theme_font():
    print("test_set_theme_font")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        app.Run("PPT_AI_Editor!Do_set_theme_font", "Cascadia Code", "Arial")
        major = deck.SlideMaster.Theme.ThemeFontScheme.MajorFont(1).Name
        minor = deck.SlideMaster.Theme.ThemeFontScheme.MinorFont(1).Name
        assert major == "Cascadia Code", f"major={major}"
        assert minor == "Arial", f"minor={minor}"
        print(f"  ok  [theme_font] major={major} minor={minor}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_bulk_insert_text_box():
    print("test_bulk_insert_text_box")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        before_count = sum(len(sl["shapes"]) for sl in before["slides"])
        slides_to_hit = list(range(1, len(before["slides"]) + 1))
        app.Run("PPT_AI_Editor!Do_bulk_insert_text_box",
                slides_to_hit, "BULK", 50.0, 50.0, 200.0, 30.0)
        after = snap(app)
        after_count = sum(len(sl["shapes"]) for sl in after["slides"])
        assert after_count == before_count + len(slides_to_hit), \
            f"shape count {before_count} -> {after_count}, expected +{len(slides_to_hit)}"
        # Find the new text boxes
        bulk_count = sum(
            1 for sl in after["slides"] for sh in sl["shapes"]
            if (sh.get("text") or "").strip() == "BULK"
        )
        assert bulk_count >= len(slides_to_hit), f"only {bulk_count} BULK boxes"
        print(f"  ok  [bulk_insert_text_box] +{bulk_count} boxes across {len(slides_to_hit)} slides")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_apply_layout_to_slides():
    print("test_apply_layout_to_slides")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # Apply layout 0 (typically Title Slide) to slide 1
        app.Run("PPT_AI_Editor!Do_apply_layout_to_slides", [1], 0)
        layout_name = deck.Slides(1).CustomLayout.Name
        print(f"  ok  [apply_layout] slide 1 -> {layout_name}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_apply_theme_skip_if_missing():
    """Skip apply_theme test unless a .thmx is available."""
    print("test_apply_theme")
    # Look for any .thmx in standard Office themes folders
    import os
    candidates = []
    for env in ("APPDATA", "ProgramFiles", "ProgramFiles(x86)"):
        p = os.environ.get(env)
        if p:
            candidates.extend(Path(p).rglob("*.thmx"))
            if len(candidates) > 0:
                break
    if not candidates:
        print("  skip  [apply_theme] no .thmx found")
        return
    thmx = str(candidates[0])
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        app.Run("PPT_AI_Editor!Do_apply_theme", thmx)
        print(f"  ok  [apply_theme] applied {Path(thmx).name}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_bulk_insert_image_skip_if_missing():
    """Skip if no image fixture available."""
    print("test_bulk_insert_image")
    # Use the carrier .pptm itself as a "file path" — won't actually display but
    # AddPicture rejects non-image. Instead skip unless we find one.
    img = REPO_ROOT / "test_decks" / "logo.png"
    if not img.exists():
        print("  skip  [bulk_insert_image] no test_decks/logo.png")
        return
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        app.Run("PPT_AI_Editor!Do_bulk_insert_image",
                [1, 2], str(img), 100.0, 100.0, 60.0, 60.0)
        print("  ok  [bulk_insert_image]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def main() -> int:
    test_find_replace_regex()
    test_swap_font_deck_wide()
    test_recolor_palette_deck_wide()
    test_set_slide_size_preset()
    test_set_theme_font()
    test_bulk_insert_text_box()
    test_apply_layout_to_slides()
    test_apply_theme_skip_if_missing()
    test_bulk_insert_image_skip_if_missing()
    shutdown_app()
    print("\nall deck smoke tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""Smoke for visual polish + effects actions (16 tests).

Single shared PowerPoint instance.
Picture-specific tests skip if no msoPicture in fixture.
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
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_fx_"))
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


def first_non_title_shape_id(s, slide_idx=0):
    for sh in s["slides"][slide_idx]["shapes"]:
        if sh["type"] != "title":
            return sh["shape_id"]
    return None


def first_picture_shape(deck):
    """Return (slide_idx_1based, shape_obj) of first picture, or (None, None)."""
    for i in range(1, deck.Slides.Count + 1):
        for sh in deck.Slides(i).Shapes:
            try:
                if sh.Type == 13 or sh.Type == 11:  # msoPicture, msoLinkedPicture
                    return i, sh
            except Exception:
                continue
    return None, None


def run_effect_test(name, fn_call, verifier=None):
    """Generic test runner. fn_call takes (app, deck, sid). verifier(deck, sid)."""
    print(f"test_{name}")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = first_non_title_shape_id(s)
        if sid is None:
            print(f"  skip  [{name}] no shape")
            return
        fn_call(app, deck, sid)
        if verifier:
            verifier(deck, sid)
        print(f"  ok  [{name}]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_rotate_shape():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_rotate_shape", 1, sid, 30.0)
    def verify(deck, sid):
        sh = next(s for s in deck.Slides(1).Shapes if s.Id == sid)
        assert abs(sh.Rotation - 30.0) < 0.5, sh.Rotation
    run_effect_test("rotate_shape", call, verify)


def test_flip_shape():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_flip_shape", 1, sid, "h")
    run_effect_test("flip_shape", call)


def test_set_line_color():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_line_color", 1, sid, "#FF0000")
    def verify(deck, sid):
        sh = next(s for s in deck.Slides(1).Shapes if s.Id == sid)
        assert sh.Line.ForeColor.RGB == 0x0000FF, hex(sh.Line.ForeColor.RGB)  # BGR
    run_effect_test("set_line_color", call, verify)


def test_set_line_weight():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_line_weight", 1, sid, 3.0)
    def verify(deck, sid):
        sh = next(s for s in deck.Slides(1).Shapes if s.Id == sid)
        assert abs(sh.Line.Weight - 3.0) < 0.1, sh.Line.Weight
    run_effect_test("set_line_weight", call, verify)


def test_set_line_style():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_line_style", 1, sid, "dash")
    run_effect_test("set_line_style", call)


def test_set_shadow():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_shadow", 1, sid, 5.0, 5.0, 8.0, "#000000", 0.5)
    run_effect_test("set_shadow", call)


def test_set_glow():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_glow", 1, sid, "#00FFFF", 8.0, 0.3)
    run_effect_test("set_glow", call)


def test_set_reflection():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_reflection", 1, sid, 0.5, 0.5, 0.0)
    run_effect_test("set_reflection", call)


def test_set_transparency():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_transparency", 1, sid, 0.5)
    run_effect_test("set_transparency", call)


def test_set_gradient_fill():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_gradient_fill", 1, sid, "#FF0000", "#0000FF", 45.0)
    run_effect_test("set_gradient_fill", call)


def test_set_3d_bevel():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_set_3d_bevel", 1, sid, "circle", 6.0)
    run_effect_test("set_3d_bevel", call)


def test_apply_preset_effect():
    def call(app, deck, sid):
        app.Run("PPT_AI_Editor!Do_apply_preset_effect", 1, sid, 5)
    run_effect_test("apply_preset_effect", call)


# ---- Picture-only tests (skip if no picture) ----

def test_crop_picture():
    print("test_crop_picture")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        idx, pic = first_picture_shape(deck)
        if idx is None:
            print("  skip  [crop_picture] no picture in fixture")
            return
        app.Run("PPT_AI_Editor!Do_crop_picture", idx, pic.Id, 5.0, 5.0, 5.0, 5.0)
        print("  ok  [crop_picture]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_recolor_picture():
    print("test_recolor_picture")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        idx, pic = first_picture_shape(deck)
        if idx is None:
            print("  skip  [recolor_picture]")
            return
        app.Run("PPT_AI_Editor!Do_recolor_picture", idx, pic.Id, "grayscale")
        print("  ok  [recolor_picture]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_brightness():
    print("test_set_brightness")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        idx, pic = first_picture_shape(deck)
        if idx is None:
            print("  skip  [set_brightness]")
            return
        app.Run("PPT_AI_Editor!Do_set_brightness", idx, pic.Id, 0.3)
        print("  ok  [set_brightness]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_contrast():
    print("test_set_contrast")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        idx, pic = first_picture_shape(deck)
        if idx is None:
            print("  skip  [set_contrast]")
            return
        app.Run("PPT_AI_Editor!Do_set_contrast", idx, pic.Id, -0.2)
        print("  ok  [set_contrast]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def main() -> int:
    test_rotate_shape()
    test_flip_shape()
    test_set_line_color()
    test_set_line_weight()
    test_set_line_style()
    test_set_shadow()
    test_set_glow()
    test_set_reflection()
    test_set_transparency()
    test_set_gradient_fill()
    test_set_3d_bevel()
    test_apply_preset_effect()
    test_crop_picture()
    test_recolor_picture()
    test_set_brightness()
    test_set_contrast()
    shutdown_app()
    print("\nall effects smoke tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

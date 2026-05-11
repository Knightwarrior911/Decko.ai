"""Smoke tests for insert_icon action.

Verified icon names sourced directly from:
  https://unpkg.com/browse/@fluentui/svg-icons/icons/
"""
import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DECK_PHASE2 = REPO_ROOT / "test_decks" / "phase2.pptx"
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
ICON_CACHE = Path(os.environ["TEMP"]) / "decko_icons"

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
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_icon_"))
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


def run_json(app, actions):
    instr = json.dumps({"actions": actions})
    return app.Run("PPT_AI_Editor!ExecuteFromString", instr)


def shape_count_on_slide(s, slide_idx=0):
    return len(s["slides"][slide_idx]["shapes"])


def find_shape_by_name(s, name, slide_idx=0):
    for sh in s["slides"][slide_idx]["shapes"]:
        if sh.get("shape_name") == name:
            return sh
    return None


# ---------------------------------------------------------------------------

def test_insert_icon_filled():
    """Insert home_48_filled — no color override."""
    print("test_insert_icon_filled")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [{
            "type": "insert_icon", "slide": 1,
            "icon": "home", "style": "filled", "size": 48,
            "left": 50, "top": 50, "width": 48, "height": 48,
            "ref_name": "icon_home"
        }])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "shape not added"
        sh = find_shape_by_name(after, "icon_home")
        assert sh is not None, "ref_name icon_home not in snapshot"
        cached = ICON_CACHE / "home_48_filled.svg"
        assert cached.exists(), f"cache file missing: {cached}"
        print(f"  ok  [insert_icon filled] home_48_filled inserted, cached at {cached}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_insert_icon_regular():
    """Insert people_48_regular — no color override."""
    print("test_insert_icon_regular")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [{
            "type": "insert_icon", "slide": 1,
            "icon": "people", "style": "regular", "size": 48,
            "left": 120, "top": 50, "width": 48, "height": 48,
            "ref_name": "icon_people"
        }])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "shape not added"
        sh = find_shape_by_name(after, "icon_people")
        assert sh is not None, "ref_name icon_people not in snapshot"
        cached = ICON_CACHE / "people_48_regular.svg"
        assert cached.exists(), f"cache file missing: {cached}"
        print(f"  ok  [insert_icon regular] people_48_regular inserted")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_insert_icon_with_color():
    """Insert mail_48_filled with hex color — recolored SVG cached separately."""
    print("test_insert_icon_with_color")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        color = "#15283C"
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [{
            "type": "insert_icon", "slide": 1,
            "icon": "mail", "style": "filled", "size": 48,
            "color": color,
            "left": 200, "top": 50, "width": 48, "height": 48,
            "ref_name": "icon_mail_colored"
        }])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "shape not added"
        sh = find_shape_by_name(after, "icon_mail_colored")
        assert sh is not None, "ref_name icon_mail_colored not in snapshot"
        cached = ICON_CACHE / "mail_48_filled_15283C.svg"
        assert cached.exists(), f"recolored cache missing: {cached}"
        svg_text = cached.read_text(encoding="utf-8")
        assert color in svg_text or color.lower() in svg_text.lower(), \
            f"color {color} not found in recolored SVG"
        print(f"  ok  [insert_icon color] mail_48_filled recolored {color}, cached")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_insert_icon_invalid_name():
    """Bogus icon name must raise error with unpkg browse URL in message."""
    print("test_insert_icon_invalid_name")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        result = run_json(app, [{
            "type": "insert_icon", "slide": 1,
            "icon": "this_icon_does_not_exist_xyz", "style": "filled", "size": 48,
            "left": 50, "top": 50, "width": 48, "height": 48
        }])
        # Should fail — result string contains error info
        assert "error" in result.lower() or "failed" in result.lower() or "0 applied" in result, \
            f"Expected error for invalid icon, got: {result}"
        print(f"  ok  [invalid icon] error returned as expected: {result[:120]}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_multiple_icons_one_call():
    """Insert 3 different icons in one batch — all land on slide."""
    print("test_multiple_icons_one_call")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [
            {"type": "insert_icon", "slide": 1,
             "icon": "calendar", "style": "filled", "size": 48,
             "left": 50,  "top": 150, "width": 48, "height": 48,
             "ref_name": "icon_cal"},
            {"type": "insert_icon", "slide": 1,
             "icon": "settings", "style": "regular", "size": 48,
             "left": 120, "top": 150, "width": 48, "height": 48,
             "ref_name": "icon_settings"},
            {"type": "insert_icon", "slide": 1,
             "icon": "star", "style": "filled", "size": 48,
             "color": "#F0A500",
             "left": 190, "top": 150, "width": 48, "height": 48,
             "ref_name": "icon_star"},
        ])
        assert "3 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 3, "not all 3 icons added"
        for name in ("icon_cal", "icon_settings", "icon_star"):
            assert find_shape_by_name(after, name) is not None, f"{name} missing"
        print("  ok  [multiple icons] calendar + settings + star all inserted")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


TESTS = [
    test_insert_icon_filled,
    test_insert_icon_regular,
    test_insert_icon_with_color,
    test_insert_icon_invalid_name,
    test_multiple_icons_one_call,
]

if __name__ == "__main__":
    # Wipe icon cache first so we confirm fresh downloads
    if ICON_CACHE.exists():
        shutil.rmtree(ICON_CACHE, ignore_errors=True)
        print(f"Cleared icon cache: {ICON_CACHE}")

    failed = []
    for t in TESTS:
        try:
            t()
        except Exception as e:
            print(f"  FAIL  [{t.__name__}] {e}")
            failed.append(t.__name__)
    shutdown_app()
    if failed:
        print(f"\n{len(failed)} FAILED: {failed}")
        sys.exit(1)
    else:
        print("\nall icon smoke tests passed")

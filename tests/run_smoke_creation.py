"""Smoke tests for Phase 1 creation actions:
  add_text_box, add_shape (with inline text + ref_name),
  z_order, duplicate_shape, shape_name lookup, expanded kinds, connector upgrades.
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
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_create_"))
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
    result = app.Run("PPT_AI_Editor!ExecuteFromString", instr)
    return result


def shape_count_on_slide(s, slide_idx=0):
    return len(s["slides"][slide_idx]["shapes"])


def shape_names_on_slide(s, slide_idx=0):
    return {sh.get("name", ""): sh for sh in s["slides"][slide_idx]["shapes"]}


def find_shape_by_name(s, name, slide_idx=0):
    for sh in s["slides"][slide_idx]["shapes"]:
        if sh.get("shape_name") == name:
            return sh
    return None


# ---------------------------------------------------------------------------

def test_add_text_box():
    print("test_add_text_box")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [
            {"type": "add_text_box", "slide": 1, "text": "Hello IB", "ref_name": "tb_hello",
             "pos": {"left": 50, "top": 50, "width": 200, "height": 40},
             "font_color": "#FF0000", "font_size": 18, "font_bold": True}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "shape not added"
        sh = find_shape_by_name(after, "tb_hello")
        assert sh is not None, "ref_name not found in snapshot"
        assert sh["text"].strip() == "Hello IB", sh["text"]
        print(f"  ok  [add_text_box] shape added, text='{sh['text'].strip()}', name=tb_hello")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_add_shape_with_text_and_ref():
    print("test_add_shape_with_text_and_ref")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "circle", "ref_name": "badge_A",
             "pos": {"left": 30, "top": 100, "width": 32, "height": 32},
             "fill": "#1F4E79", "stroke": None,
             "text": "A", "font_color": "#FFFFFF", "font_size": 11, "font_bold": True}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "shape not added"
        sh = find_shape_by_name(after, "badge_A")
        assert sh is not None, "ref_name badge_A not found"
        assert sh["text"].strip() == "A", sh["text"]
        print(f"  ok  [add_shape+text+ref] text='{sh['text'].strip()}', name=badge_A")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_shape_name_lookup():
    print("test_shape_name_lookup")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # Create shape with ref_name, then edit it using shape_name instead of shape_id
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "named_rect",
             "pos": {"left": 100, "top": 200, "width": 150, "height": 50},
             "fill": "#CCCCCC", "stroke": None}
        ])
        result = run_json(app, [
            {"type": "set_text", "slide": 1, "shape_name": "named_rect", "value": "via name"}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        sh = find_shape_by_name(after, "named_rect")
        assert sh is not None, "named_rect not found"
        assert "via name" in sh["text"], sh["text"]
        print(f"  ok  [shape_name_lookup] text='{sh['text'].strip()}'")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_z_order():
    print("test_z_order")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        ids = [sh["shape_id"] for sh in s["slides"][0]["shapes"] if sh.get("type") != "title"]
        if len(ids) < 1:
            print("  skip  [no shapes]")
            return
        sid = ids[0]
        result = run_json(app, [
            {"type": "z_order", "slide": 1, "shape_id": sid, "order": "front"}
        ])
        assert "1 applied" in result, result
        result2 = run_json(app, [
            {"type": "z_order", "slide": 1, "shape_id": sid, "order": "back"}
        ])
        assert "1 applied" in result2, result2
        print(f"  ok  [z_order] front+back on shape {sid}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_duplicate_shape():
    print("test_duplicate_shape")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        ids = [sh["shape_id"] for sh in s["slides"][0]["shapes"] if sh.get("type") != "title"]
        if len(ids) < 1:
            print("  skip  [no shapes]")
            return
        sid = ids[0]
        count_before = shape_count_on_slide(s)
        result = run_json(app, [
            {"type": "duplicate_shape", "slide": 1, "shape_id": sid,
             "left": 400, "top": 300, "ref_name": "dup_shape"}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "dup not added"
        sh = find_shape_by_name(after, "dup_shape")
        assert sh is not None, "dup ref_name not found"
        assert abs(sh["pos"]["left"] - 400) < 2, sh["pos"]
        assert abs(sh["pos"]["top"] - 300) < 2, sh["pos"]
        print(f"  ok  [duplicate_shape] at {sh['pos']['left']},{sh['pos']['top']}, name=dup_shape")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_expanded_shape_kinds():
    print("test_expanded_shape_kinds")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        kinds = ["chevron", "hexagon", "left_arrow", "up_arrow", "parallelogram", "cross"]
        actions = []
        for i, k in enumerate(kinds):
            actions.append({
                "type": "add_shape", "slide": 1, "kind": k,
                "pos": {"left": 50 + i * 110, "top": 350, "width": 90, "height": 40},
                "fill": "#2E75B6", "stroke": None
            })
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, actions)
        assert f"{len(kinds)} applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + len(kinds), "not all shapes added"
        print(f"  ok  [expanded_kinds] {kinds}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_numeric_shape_kind():
    print("test_numeric_shape_kind")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        # mso_52 = chevron, mso_9 = oval
        result = run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "mso_52",
             "pos": {"left": 50, "top": 400, "width": 100, "height": 40},
             "fill": "#70AD47", "stroke": None},
            {"type": "add_shape", "slide": 1, "kind": "9",
             "pos": {"left": 160, "top": 400, "width": 60, "height": 60},
             "fill": "#ED7D31", "stroke": None}
        ])
        assert "2 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 2
        print("  ok  [numeric_kind] mso_52 and bare '9' both created")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_connector_with_points():
    print("test_connector_with_points")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # Create two rects then connect them with a specific from/to point and start arrow
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "box_top",
             "pos": {"left": 100, "top": 80, "width": 120, "height": 50},
             "fill": "#1F4E79", "stroke": None},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "box_bot",
             "pos": {"left": 100, "top": 200, "width": 120, "height": 50},
             "fill": "#1F4E79", "stroke": None},
        ])
        s = snap(app)
        top_id = find_shape_by_name(s, "box_top")["shape_id"]
        bot_id = find_shape_by_name(s, "box_bot")["shape_id"]
        count_before = shape_count_on_slide(s)
        result = run_json(app, [
            {"type": "add_connector", "slide": 1,
             "from_shape_id": top_id, "to_shape_id": bot_id,
             "kind": "straight", "arrow_end": "filled", "arrow_start": "none",
             "from_point": "bottom", "to_point": "top",
             "color": "#FF0000", "weight_pt": 2.0, "dash_style": "solid"}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 1, "connector not added"
        print("  ok  [connector_with_points] bottom-to-top, red, 2pt")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_badge_pattern():
    """Full badge pattern: rect + circle overlay using ref_name and z_order."""
    print("test_badge_pattern")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        count_before = shape_count_on_slide(before)
        result = run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "badge_rect",
             "pos": {"left": 50, "top": 140, "width": 130, "height": 65},
             "fill": "#2E75B6", "stroke": None,
             "text": "Precedent\nTransactions", "font_color": "#FFFFFF",
             "font_size": 11, "font_bold": True},
            {"type": "add_shape", "slide": 1, "kind": "circle", "ref_name": "badge_circle",
             "pos": {"left": 36, "top": 128, "width": 30, "height": 30},
             "fill": "#1F4E79", "stroke": None,
             "text": "B", "font_color": "#FFFFFF", "font_size": 11, "font_bold": True},
            {"type": "z_order", "slide": 1, "shape_name": "badge_circle", "order": "front"}
        ])
        assert "3 applied" in result, result
        after = snap(app)
        assert shape_count_on_slide(after) == count_before + 2
        rect_sh = find_shape_by_name(after, "badge_rect")
        circ_sh = find_shape_by_name(after, "badge_circle")
        assert rect_sh is not None and circ_sh is not None
        assert "Precedent" in rect_sh["text"], rect_sh["text"]
        assert circ_sh["text"].strip() == "B", circ_sh["text"]
        print("  ok  [badge_pattern] rect+circle+z_order in 3 actions")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


TESTS = [
    test_add_text_box,
    test_add_shape_with_text_and_ref,
    test_shape_name_lookup,
    test_z_order,
    test_duplicate_shape,
    test_expanded_shape_kinds,
    test_numeric_shape_kind,
    test_connector_with_points,
    test_badge_pattern,
]

if __name__ == "__main__":
    failed = []
    for t in TESTS:
        try:
            t()
        except Exception as e:
            print(f"  FAIL  {e}")
            failed.append(t.__name__)
    shutdown_app()
    if failed:
        print(f"\n{len(failed)} FAILED: {failed}")
        sys.exit(1)
    else:
        print(f"\nall creation smoke tests passed")

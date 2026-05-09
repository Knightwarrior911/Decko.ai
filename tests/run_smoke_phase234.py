"""Smoke tests for Phase 2/3/4: tables, slide bg, copy formatting, strikethrough."""
import json
import shutil
import sys
import tempfile
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
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_p234_"))
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


def find_shape_by_name(s, name, slide_idx=0):
    for sh in s["slides"][slide_idx]["shapes"]:
        if sh.get("shape_name") == name:
            return sh
    return None


def shape_count(s, slide_idx=0):
    return len(s["slides"][slide_idx]["shapes"])


# ---------------------------------------------------------------------------

def test_add_table():
    print("test_add_table")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        cb = shape_count(before)
        result = run_json(app, [
            {"type": "add_table", "slide": 1, "rows": 4, "cols": 3, "ref_name": "comp_table",
             "pos": {"left": 50, "top": 100, "width": 600, "height": 200}}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count(after) == cb + 1
        sh = find_shape_by_name(after, "comp_table")
        assert sh is not None, "table ref_name not found"
        assert sh["type"] == "table", sh["type"]
        print(f"  ok  [add_table] 4x3 table created, name=comp_table")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_table_sizing_and_borders():
    print("test_table_sizing_and_borders")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_table", "slide": 1, "rows": 3, "cols": 3, "ref_name": "tbl",
             "pos": {"left": 50, "top": 100, "width": 600, "height": 150}}
        ])
        s = snap(app)
        sid = find_shape_by_name(s, "tbl")["shape_id"]
        result = run_json(app, [
            {"type": "set_table_col_width", "slide": 1, "shape_id": sid, "col": 1, "width_pt": 300},
            {"type": "set_table_row_height", "slide": 1, "shape_id": sid, "row": 1, "height_pt": 60},
            {"type": "set_cell_border", "slide": 1, "shape_id": sid, "row": 1, "col": 1,
             "side": "all", "color": "#FF0000", "weight_pt": 2.0, "visible": True},
            {"type": "set_cell_text_align", "slide": 1, "shape_id": sid, "row": 1, "col": 1,
             "h_align": "center", "v_align": "middle"},
            {"type": "set_cell_fill", "slide": 1, "shape_id": sid, "row": 1, "col": 1, "color": "#1F4E79"}
        ])
        assert "5 applied" in result, result
        print(f"  ok  [table_sizing+borders+fill+align] all 5 actions applied")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_apply_table_style():
    print("test_apply_table_style")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_table", "slide": 1, "rows": 3, "cols": 3, "ref_name": "tbl_styled",
             "pos": {"left": 50, "top": 100, "width": 600, "height": 150}}
        ])
        s = snap(app)
        sid = find_shape_by_name(s, "tbl_styled")["shape_id"]
        result = run_json(app, [
            {"type": "apply_table_style", "slide": 1, "shape_id": sid,
             "style_id": "medium_style_2_accent1"}
        ])
        assert "1 applied" in result, result
        print("  ok  [apply_table_style] medium_style_2_accent1 applied")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_slide_background():
    print("test_set_slide_background")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        result = run_json(app, [
            {"type": "set_slide_background_color", "slide": 1, "color": "#1F4E79"}
        ])
        assert "1 applied" in result, result
        print("  ok  [set_slide_background] navy applied")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_copy_formatting():
    print("test_copy_formatting")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "src_box",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50},
             "fill": "#1F4E79", "stroke": "#FF0000", "stroke_weight_pt": 3.0,
             "text": "Source", "font_color": "#FFFFFF", "font_size": 14, "font_bold": True},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "tgt_box",
             "pos": {"left": 200, "top": 50, "width": 100, "height": 50},
             "fill": "#CCCCCC", "stroke": None, "text": "Target"}
        ])
        s = snap(app)
        src_id = find_shape_by_name(s, "src_box")["shape_id"]
        tgt_id = find_shape_by_name(s, "tgt_box")["shape_id"]
        result = run_json(app, [
            {"type": "copy_formatting", "slide": 1,
             "source_shape_id": src_id, "target_shape_id": tgt_id}
        ])
        assert "1 applied" in result, result
        print("  ok  [copy_formatting] source style copied to target")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_insert_slide_number():
    print("test_insert_slide_number")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        cb = shape_count(before)
        result = run_json(app, [
            {"type": "insert_slide_number", "slide": 1, "ref_name": "page_num",
             "pos": {"left": 600, "top": 500, "width": 80, "height": 25},
             "font_color": "#888888", "font_size": 10}
        ])
        assert "1 applied" in result, result
        after = snap(app)
        assert shape_count(after) == cb + 1
        print("  ok  [insert_slide_number] page number text box created")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_strikethrough():
    print("test_strikethrough")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_text_box", "slide": 1, "text": "Old number", "ref_name": "tb_old",
             "pos": {"left": 50, "top": 50, "width": 200, "height": 30},
             "font_size": 14}
        ])
        s = snap(app)
        sid = find_shape_by_name(s, "tb_old")["shape_id"]
        result = run_json(app, [
            {"type": "set_run_strikethrough", "slide": 1, "shape_id": sid,
             "paragraph_index": 0, "run_index": 0, "value": True}
        ])
        assert "1 applied" in result, result
        print("  ok  [strikethrough] applied to run 0")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


TESTS = [
    test_add_table,
    test_table_sizing_and_borders,
    test_apply_table_style,
    test_set_slide_background,
    test_copy_formatting,
    test_insert_slide_number,
    test_strikethrough,
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
        print(f"\nall phase 2/3/4 smoke tests passed")

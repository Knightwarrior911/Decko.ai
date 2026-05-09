"""Bug audit smoke: creation features edge cases + enum correctness."""
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
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_audit_"))
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


def find_shape(s, name=None, sid=None, slide_idx=0):
    for sh in s["slides"][slide_idx]["shapes"]:
        if name and sh.get("shape_name") == name:
            return sh
        if sid and sh.get("shape_id") == sid:
            return sh
    return None


# ---------------------------------------------------------------------------
# BUG 1: connector arrow_end="filled" should produce triangle (msoArrowheadTriangle=2)

def test_connector_arrow_filled_is_triangle():
    print("test_connector_arrow_filled_is_triangle")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "from_box",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50}},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "to_box",
             "pos": {"left": 300, "top": 50, "width": 100, "height": 50}}
        ])
        s = snap(app)
        from_id = find_shape(s, name="from_box")["shape_id"]
        to_id = find_shape(s, name="to_box")["shape_id"]
        run_json(app, [
            {"type": "add_connector", "slide": 1,
             "from_shape_id": from_id, "to_shape_id": to_id,
             "kind": "straight", "arrow_end": "filled"}
        ])
        # Read back the arrow head style of the new connector via COM
        slide1 = deck.Slides(1)
        conn = None
        for i in range(1, slide1.Shapes.Count + 1):
            sh = slide1.Shapes(i)
            try:
                if sh.Connector:
                    conn = sh
            except Exception:
                pass
        assert conn is not None, "connector not found"
        end_style = conn.Line.EndArrowheadStyle
        # msoArrowheadTriangle = 2 ; msoArrowheadDiamond = 5
        assert end_style == 2, f"arrow_end='filled' should yield triangle (2), got {end_style}"
        print(f"  ok  [arrow_filled] EndArrowheadStyle={end_style} (msoArrowheadTriangle)")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_connector_arrow_open_is_open():
    print("test_connector_arrow_open_is_open")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50}},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "b",
             "pos": {"left": 300, "top": 50, "width": 100, "height": 50}}
        ])
        s = snap(app)
        run_json(app, [
            {"type": "add_connector", "slide": 1,
             "from_shape_id": find_shape(s, name="a")["shape_id"],
             "to_shape_id": find_shape(s, name="b")["shape_id"],
             "kind": "straight", "arrow_end": "open"}
        ])
        slide1 = deck.Slides(1)
        conn = None
        for i in range(1, slide1.Shapes.Count + 1):
            sh = slide1.Shapes(i)
            try:
                if sh.Connector:
                    conn = sh
            except Exception:
                pass
        end_style = conn.Line.EndArrowheadStyle
        # msoArrowheadOpen = 3
        assert end_style == 3, f"arrow_end='open' should yield open (3), got {end_style}"
        print(f"  ok  [arrow_open] EndArrowheadStyle={end_style} (msoArrowheadOpen)")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_connector_arrow_diamond_is_diamond():
    print("test_connector_arrow_diamond_is_diamond")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50}},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "b",
             "pos": {"left": 300, "top": 50, "width": 100, "height": 50}}
        ])
        s = snap(app)
        run_json(app, [
            {"type": "add_connector", "slide": 1,
             "from_shape_id": find_shape(s, name="a")["shape_id"],
             "to_shape_id": find_shape(s, name="b")["shape_id"],
             "kind": "straight", "arrow_end": "diamond"}
        ])
        slide1 = deck.Slides(1)
        conn = None
        for i in range(1, slide1.Shapes.Count + 1):
            sh = slide1.Shapes(i)
            try:
                if sh.Connector:
                    conn = sh
            except Exception:
                pass
        end_style = conn.Line.EndArrowheadStyle
        # msoArrowheadDiamond = 5
        assert end_style == 5, f"arrow_end='diamond' should yield diamond (5), got {end_style}"
        print(f"  ok  [arrow_diamond] EndArrowheadStyle={end_style} (msoArrowheadDiamond)")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 2: star4 and star16 collision

def test_star4_distinct_from_star16():
    print("test_star4_distinct_from_star16")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "star4", "ref_name": "s4",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 100}},
            {"type": "add_shape", "slide": 1, "kind": "star16", "ref_name": "s16",
             "pos": {"left": 200, "top": 50, "width": 100, "height": 100}}
        ])
        slide1 = deck.Slides(1)
        s4_kind = None
        s16_kind = None
        for i in range(1, slide1.Shapes.Count + 1):
            sh = slide1.Shapes(i)
            if sh.Name == "s4":
                s4_kind = sh.AutoShapeType
            elif sh.Name == "s16":
                s16_kind = sh.AutoShapeType
        assert s4_kind is not None and s16_kind is not None, "shapes not found"
        assert s4_kind != s16_kind, f"star4 and star16 share AutoShapeType={s4_kind} — should differ"
        print(f"  ok  [star_distinct] star4={s4_kind}, star16={s16_kind}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 3: add_shape v_align param (currently hardcoded middle)

def test_add_shape_v_align_top():
    print("test_add_shape_v_align_top")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        result = run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "tb",
             "pos": {"left": 50, "top": 50, "width": 200, "height": 100},
             "text": "Top aligned", "v_align": "top"}
        ])
        assert "1 applied" in result, result
        slide1 = deck.Slides(1)
        for i in range(1, slide1.Shapes.Count + 1):
            sh = slide1.Shapes(i)
            if sh.Name == "tb":
                # msoAnchorTop = 1, msoAnchorMiddle = 3, msoAnchorBottom = 4
                anchor = sh.TextFrame.VerticalAnchor
                assert anchor == 1, f"v_align=top should yield msoAnchorTop (1), got {anchor}"
                print(f"  ok  [v_align_top] VerticalAnchor={anchor}")
                return
        raise AssertionError("shape not found")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 4: add_text_box should respect width/height (no auto-grow)

def test_add_text_box_respects_dimensions():
    print("test_add_text_box_respects_dimensions")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_text_box", "slide": 1, "text": "Hi",
             "ref_name": "tb_size",
             "pos": {"left": 50, "top": 50, "width": 200, "height": 80}}
        ])
        slide1 = deck.Slides(1)
        for i in range(1, slide1.Shapes.Count + 1):
            sh = slide1.Shapes(i)
            if sh.Name == "tb_size":
                w, h = sh.Width, sh.Height
                # Allow ~2pt tolerance
                assert abs(w - 200) < 5, f"width should be ~200, got {w}"
                assert abs(h - 80) < 5, f"height should be ~80, got {h}"
                print(f"  ok  [textbox_dims] w={w:.1f}, h={h:.1f}")
                return
        raise AssertionError("shape not found")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 5: ref_name uniqueness — what happens if two shapes get same ref_name?

def test_duplicate_ref_name():
    print("test_duplicate_ref_name")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        before = snap(app)
        cb = len(before["slides"][0]["shapes"])
        result = run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "dup",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50}},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "dup",
             "pos": {"left": 200, "top": 50, "width": 100, "height": 50}}
        ])
        assert "2 applied" in result, result
        after = snap(app)
        assert len(after["slides"][0]["shapes"]) == cb + 2, "both shapes should be created"
        # FindShapeByName returns first match — that's the documented behavior
        slide1 = deck.Slides(1)
        dup_count = sum(1 for i in range(1, slide1.Shapes.Count + 1)
                        if slide1.Shapes(i).Name == "dup")
        print(f"  ok  [duplicate_ref_name] {dup_count} shapes named 'dup' (PPT allows)")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 6: shape_name lookup for z_order and duplicate_shape

def test_shape_name_works_for_z_order():
    print("test_shape_name_works_for_z_order")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "back_rect",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 100},
             "fill": "#FF0000"},
            {"type": "add_shape", "slide": 1, "kind": "circle", "ref_name": "front_circle",
             "pos": {"left": 75, "top": 75, "width": 50, "height": 50},
             "fill": "#00FF00"}
        ])
        result = run_json(app, [
            {"type": "z_order", "slide": 1, "shape_name": "back_rect", "order": "back"}
        ])
        assert "1 applied" in result, result
        print("  ok  [z_order_via_name] resolved shape_name -> shape_id")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 7: invalid kind returns clear error (not silent fail)

def test_invalid_shape_kind_errors():
    print("test_invalid_shape_kind_errors")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        result = run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "nonsense_kind_xyz",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50}}
        ])
        # Should be skipped with error message about unknown kind
        assert "0 applied" in result and "1 skipped" in result, \
            f"invalid kind should be skipped, got: {result}"
        print(f"  ok  [invalid_kind_skipped] {result.split('Log')[0].strip()}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 8: rows=0 should fail clearly

def test_add_table_zero_rows_errors():
    print("test_add_table_zero_rows_errors")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        result = run_json(app, [
            {"type": "add_table", "slide": 1, "rows": 0, "cols": 3,
             "pos": {"left": 50, "top": 50, "width": 400, "height": 100}}
        ])
        assert "0 applied" in result and "1 skipped" in result, \
            f"rows=0 should be skipped, got: {result}"
        print(f"  ok  [rows_zero_skipped]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 9: connector with from_point but to_point=auto

def test_connector_partial_auto_route():
    print("test_connector_partial_auto_route")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a",
             "pos": {"left": 50, "top": 50, "width": 100, "height": 50}},
            {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "b",
             "pos": {"left": 50, "top": 200, "width": 100, "height": 50}}
        ])
        s = snap(app)
        result = run_json(app, [
            {"type": "add_connector", "slide": 1,
             "from_shape_id": find_shape(s, name="a")["shape_id"],
             "to_shape_id": find_shape(s, name="b")["shape_id"],
             "kind": "elbow", "from_point": "bottom", "to_point": "auto",
             "arrow_end": "filled"}
        ])
        assert "1 applied" in result, result
        print(f"  ok  [partial_auto] no crash w/ mixed explicit+auto endpoints")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------------------------------------------------------------------------
# BUG 10: set_run_strikethrough idempotency (set true then false)

def test_strikethrough_idempotent():
    print("test_strikethrough_idempotent")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        run_json(app, [
            {"type": "add_text_box", "slide": 1, "text": "Test",
             "ref_name": "tb_strike",
             "pos": {"left": 50, "top": 50, "width": 200, "height": 30},
             "font_size": 14}
        ])
        s = snap(app)
        sid = find_shape(s, name="tb_strike")["shape_id"]
        result = run_json(app, [
            {"type": "set_run_strikethrough", "slide": 1, "shape_id": sid,
             "paragraph_index": 0, "run_index": 0, "value": True},
            {"type": "set_run_strikethrough", "slide": 1, "shape_id": sid,
             "paragraph_index": 0, "run_index": 0, "value": False}
        ])
        assert "2 applied" in result, result
        print("  ok  [strike_idempotent] set true then false")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


TESTS = [
    test_connector_arrow_filled_is_triangle,
    test_connector_arrow_open_is_open,
    test_connector_arrow_diamond_is_diamond,
    test_star4_distinct_from_star16,
    test_add_shape_v_align_top,
    test_add_text_box_respects_dimensions,
    test_duplicate_ref_name,
    test_shape_name_works_for_z_order,
    test_invalid_shape_kind_errors,
    test_add_table_zero_rows_errors,
    test_connector_partial_auto_route,
    test_strikethrough_idempotent,
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
        print(f"\nall audit smoke tests passed")

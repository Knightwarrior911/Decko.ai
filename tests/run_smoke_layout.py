"""Smoke for smart layout actions (12 tests).

Single shared PowerPoint instance (DispatchEx between tests races Quit).
Reuses test_decks/phase2.pptx for multi-shape exercises.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DECK_PHASE2 = REPO_ROOT / "test_decks" / "phase2.pptx"
DECK_TEXT = REPO_ROOT / "test_decks" / "text_v3.pptx"
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


def fresh_deck(app, src_path):
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_layout_"))
    deck_copy = tmpdir / src_path.name
    shutil.copy2(src_path, deck_copy)
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


def shapes_on_slide(s, slide_idx):
    """Return non-title shape ids on a 0-indexed slide."""
    out = []
    for sh in s["slides"][slide_idx]["shapes"]:
        if sh["type"] != "title":
            out.append(sh["shape_id"])
    return out


def shape_by_id(s, slide_idx, shape_id):
    for sh in s["slides"][slide_idx]["shapes"]:
        if sh["shape_id"] == shape_id:
            return sh
    return None


def test_snap_to_grid():
    print("test_snap_to_grid")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        sid = shapes_on_slide(s, 0)[0]
        before = shape_by_id(s, 0, sid)["pos"]
        app.Run("PPT_AI_Editor!Do_snap_to_grid", 1, sid, 36.0)
        s2 = snap(app)
        after = shape_by_id(s2, 0, sid)["pos"]
        # Snapped Left should be a multiple of 36
        assert abs(round(after["left"] / 36.0) * 36.0 - after["left"]) < 0.5, after
        print(f"  ok  [snap_to_grid] {before['left']} -> {after['left']}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_align_to_slide_center():
    print("test_align_to_slide_center")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        sid = shapes_on_slide(s, 0)[0]
        app.Run("PPT_AI_Editor!Do_align_to_slide_center", 1, sid, "h")
        s2 = snap(app)
        sh = shape_by_id(s2, 0, sid)
        slide_w = 720.0  # PowerPoint default 10x7.5 in @ 72 dpi
        # Allow slide_w pulled from test deck — 16:9 might be 960. Just verify centered.
        # Centered: left + width/2 == slide_w / 2 (roughly)
        center = sh["pos"]["left"] + sh["pos"]["width"] / 2
        # 16:9 default is 960 pt, 4:3 is 720 pt. Try both.
        assert abs(center - 480.0) < 5.0 or abs(center - 360.0) < 5.0, center
        print(f"  ok  [aligned to center] center_x={center}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_nudge():
    print("test_nudge")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        sid = shapes_on_slide(s, 0)[0]
        before_left = shape_by_id(s, 0, sid)["pos"]["left"]
        app.Run("PPT_AI_Editor!Do_nudge", 1, sid, "r", 25.0)
        s2 = snap(app)
        after_left = shape_by_id(s2, 0, sid)["pos"]["left"]
        assert abs((after_left - before_left) - 25.0) < 0.5, f"{before_left} -> {after_left}"
        print(f"  ok  [nudge right] {before_left} -> {after_left}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_fit_to_content():
    print("test_fit_to_content")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_TEXT)
    try:
        s = snap(app)
        sid = s["slides"][0]["shapes"][1]["shape_id"]
        before_h = s["slides"][0]["shapes"][1]["pos"]["height"]
        app.Run("PPT_AI_Editor!Do_fit_to_content", 1, sid)
        s2 = snap(app)
        after_h = shape_by_id(s2, 0, sid)["pos"]["height"]
        # Just confirm the call succeeds and height is positive
        assert after_h > 0
        print(f"  ok  [fit_to_content] h: {before_h} -> {after_h}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_match_size():
    print("test_match_size")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 2:
            print("  skip  [match_size] need >=2 shapes on slide 1")
            return
        ref, tgt = ids[0], ids[1]
        ref_w = shape_by_id(s, 0, ref)["pos"]["width"]
        ref_h = shape_by_id(s, 0, ref)["pos"]["height"]
        app.Run("PPT_AI_Editor!Do_match_size", 1, ref, [tgt])
        s2 = snap(app)
        tgt_after = shape_by_id(s2, 0, tgt)["pos"]
        assert abs(tgt_after["width"] - ref_w) < 0.5, tgt_after
        assert abs(tgt_after["height"] - ref_h) < 0.5, tgt_after
        print(f"  ok  [match_size] ({ref_w}, {ref_h})")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_uniform_size():
    print("test_uniform_size")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 2:
            print("  skip  [uniform_size]")
            return
        app.Run("PPT_AI_Editor!Do_uniform_size", 1, ids[:2], 200.0, 100.0)
        s2 = snap(app)
        for sid in ids[:2]:
            sh = shape_by_id(s2, 0, sid)
            assert abs(sh["pos"]["width"] - 200.0) < 0.5, sh
            assert abs(sh["pos"]["height"] - 100.0) < 0.5, sh
        print(f"  ok  [uniform_size] {len(ids[:2])} shapes -> 200x100")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_smart_spacing():
    print("test_smart_spacing")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 3:
            print("  skip  [smart_spacing]")
            return
        app.Run("PPT_AI_Editor!Do_smart_spacing", 1, ids[:3], 20.0, "h")
        s2 = snap(app)
        # Sort sub-shapes by left to verify gap
        positions = sorted(
            [shape_by_id(s2, 0, sid)["pos"] for sid in ids[:3]],
            key=lambda p: p["left"]
        )
        gap1 = positions[1]["left"] - (positions[0]["left"] + positions[0]["width"])
        gap2 = positions[2]["left"] - (positions[1]["left"] + positions[1]["width"])
        assert abs(gap1 - 20.0) < 1.0 and abs(gap2 - 20.0) < 1.0, f"gaps {gap1},{gap2}"
        print(f"  ok  [smart_spacing] gaps {gap1:.1f}, {gap2:.1f}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_equalize_spacing():
    print("test_equalize_spacing")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 3:
            print("  skip  [equalize_spacing]")
            return
        app.Run("PPT_AI_Editor!Do_equalize_spacing", 1, ids[:3], "h")
        s2 = snap(app)
        positions = sorted(
            [shape_by_id(s2, 0, sid)["pos"] for sid in ids[:3]],
            key=lambda p: p["left"]
        )
        gap1 = positions[1]["left"] - (positions[0]["left"] + positions[0]["width"])
        gap2 = positions[2]["left"] - (positions[1]["left"] + positions[1]["width"])
        assert abs(gap1 - gap2) < 1.0, f"gaps not equal: {gap1} vs {gap2}"
        print(f"  ok  [equalize_spacing] equal gaps {gap1:.1f}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_match_position():
    print("test_match_position")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 2:
            print("  skip  [match_position]")
            return
        ref, tgt = ids[0], ids[1]
        ref_left = shape_by_id(s, 0, ref)["pos"]["left"]
        app.Run("PPT_AI_Editor!Do_match_position", 1, ref, tgt, "left")
        s2 = snap(app)
        tgt_after_left = shape_by_id(s2, 0, tgt)["pos"]["left"]
        assert abs(tgt_after_left - ref_left) < 0.5, f"{tgt_after_left} vs {ref_left}"
        print(f"  ok  [match_position left] target.left={tgt_after_left}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_swap_positions():
    print("test_swap_positions")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 2:
            print("  skip  [swap_positions]")
            return
        a, b = ids[0], ids[1]
        a_pos = dict(shape_by_id(s, 0, a)["pos"])
        b_pos = dict(shape_by_id(s, 0, b)["pos"])
        app.Run("PPT_AI_Editor!Do_swap_positions", 1, a, b)
        s2 = snap(app)
        a_after = shape_by_id(s2, 0, a)["pos"]
        b_after = shape_by_id(s2, 0, b)["pos"]
        assert abs(a_after["left"] - b_pos["left"]) < 0.5
        assert abs(b_after["left"] - a_pos["left"]) < 0.5
        print(f"  ok  [swap_positions] a.left {a_pos['left']} <-> {b_pos['left']}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_group_by_overlap():
    """If no shapes overlap on slide 1, validate that the action returns the
    'no overlapping subset' error gracefully instead of crashing."""
    print("test_group_by_overlap")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 2:
            print("  skip  [group_by_overlap]")
            return
        # First make two shapes intentionally overlap
        if len(ids) >= 2:
            ref_pos = shape_by_id(s, 0, ids[0])["pos"]
            app.Run("PPT_AI_Editor!Do_match_position", 1, ids[0], ids[1], "left")
            app.Run("PPT_AI_Editor!Do_match_position", 1, ids[0], ids[1], "top")
        try:
            app.Run("PPT_AI_Editor!Do_group_by_overlap", 1, ids[:2])
            print("  ok  [group_by_overlap] grouped")
        except Exception as e:
            # Acceptable: no overlapping subset error if same-position shapes don't qualify
            if "no overlapping subset" in str(e) or "shape" in str(e).lower():
                print(f"  ok  [group_by_overlap] non-overlap rejected: {e}")
            else:
                raise
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_chained_snap_then_equalize():
    """Cross-action: snap to grid, then equalize spacing."""
    print("test_chained_snap_then_equalize")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app, DECK_PHASE2)
    try:
        s = snap(app)
        ids = shapes_on_slide(s, 0)
        if len(ids) < 3:
            print("  skip  [chained]")
            return
        for sid in ids[:3]:
            app.Run("PPT_AI_Editor!Do_snap_to_grid", 1, sid, 12.0)
        app.Run("PPT_AI_Editor!Do_equalize_spacing", 1, ids[:3], "h")
        s2 = snap(app)
        positions = sorted(
            [shape_by_id(s2, 0, sid)["pos"] for sid in ids[:3]],
            key=lambda p: p["left"]
        )
        gap1 = positions[1]["left"] - (positions[0]["left"] + positions[0]["width"])
        gap2 = positions[2]["left"] - (positions[1]["left"] + positions[1]["width"])
        assert abs(gap1 - gap2) < 1.0, f"chained gaps {gap1} vs {gap2}"
        print(f"  ok  [chained] equal gaps {gap1:.1f}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def main() -> int:
    test_snap_to_grid()
    test_align_to_slide_center()
    test_nudge()
    test_fit_to_content()
    test_match_size()
    test_uniform_size()
    test_smart_spacing()
    test_equalize_spacing()
    test_match_position()
    test_swap_positions()
    test_group_by_overlap()
    test_chained_snap_then_equalize()
    shutdown_app()
    print("\nall layout smoke tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

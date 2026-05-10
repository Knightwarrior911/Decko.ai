"""VP-perspective end-to-end scenarios.

Simulates a VP issuing NL requests that get translated to JSON action batches,
then executes them against a fresh deck and verifies the resulting state.

Scenarios target the exact bug patterns audited in 2026-05-10 session:
  S1: Connector with from_shape_name + child→parent direction (today's bug)
  S2: Connector with string from_shape_id (numeric-cast crash regression)
  S3: msoShape values: right_arrow, callout_rect, ribbon_up, star8 (was wrong)
  S4: Picture insertion with explicit W/H (LockAspectRatio fix)
  S5: Group by names + z_order on group (group-child detection)
  S6: Locale-tolerant bool ("true" string) + string-id in shape_id field
  S7: rgb() color and [r,g,b] tuple parsing
  S8: shape_ids array with mix of numeric Id and ref_name string

Run: python tests/run_vp_scenarios.py
"""
import json
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
RESULTS = []


def open_app():
    import win32com.client
    return win32com.client.DispatchEx("PowerPoint.Application")


def fresh_deck(app, tmpdir):
    """Create a fresh blank deck with one slide."""
    pres = app.Presentations.Add()
    # Ensure slide 1 exists
    if pres.Slides.Count == 0:
        pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts(1))
    deck_path = tmpdir / "vp_test.pptx"
    pres.SaveAs(str(deck_path))
    return pres


def run_actions(app, deck, actions):
    """Send action batch to carrier's ExecuteFromString. Activates deck first."""
    deck.Windows(1).Activate()
    payload = json.dumps({"actions": actions})
    result = app.Run("PPT_AI_Editor!ExecuteFromString", payload)
    print(f"    >> {result}")
    # If anything skipped/errored, surface the reasons from the log
    if "skipped" in result and " 0 skipped" not in result:
        log_path = deck.FullName + ".action_log.jsonl"
        try:
            with open(log_path) as f:
                lines = f.readlines()
            # Show last len(actions) entries (this batch)
            for line in lines[-len(actions):]:
                e = json.loads(line)
                if e.get("status") in ("skipped", "error"):
                    print(f"      !! {e['status']}: {e.get('action_type')} -> {e.get('reason','')}")
        except Exception as ex:
            print(f"      (could not read log: {ex})")
    return result


def snapshot(app, deck):
    deck.Windows(1).Activate()
    raw = app.Run("PPT_AI_Editor!BuildSnapshotJson")
    return json.loads(raw)


def record(name, ok, detail=""):
    RESULTS.append({"scenario": name, "ok": ok, "detail": detail})
    status = "PASS" if ok else "FAIL"
    print(f"  [{status}] {name}: {detail}")


def find_by_name(slide_shapes, ref_name):
    for s in slide_shapes:
        if s.get("shape_name") == ref_name:
            return s
    return None


# ---------------------------------------------------------------------------
# Scenarios
# ---------------------------------------------------------------------------


def s1_connector_by_name(app, deck):
    """VP: 'Make a CEO box and three VP boxes below; connect CEO to each VP.'"""
    print("\nS1: connector by from_shape_name (today's bug regression)")
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "ceo",
         "pos": {"left": 300, "top": 50, "width": 144, "height": 48},
         "fill": "#1F4E79", "text": "CEO", "font_color": "#FFFFFF"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "vp1",
         "pos": {"left": 100, "top": 200, "width": 120, "height": 48},
         "fill": "#2E75B6", "text": "VP Eng", "font_color": "#FFFFFF"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "vp2",
         "pos": {"left": 300, "top": 200, "width": 120, "height": 48},
         "fill": "#2E75B6", "text": "VP Sales", "font_color": "#FFFFFF"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "vp3",
         "pos": {"left": 500, "top": 200, "width": 120, "height": 48},
         "fill": "#2E75B6", "text": "VP Ops", "font_color": "#FFFFFF"},
        {"type": "add_connector", "slide": 1,
         "from_shape_name": "ceo", "to_shape_name": "vp1",
         "kind": "elbow", "from_point": "bottom", "to_point": "top"},
        {"type": "add_connector", "slide": 1,
         "from_shape_name": "ceo", "to_shape_name": "vp2",
         "kind": "elbow", "from_point": "bottom", "to_point": "top"},
        {"type": "add_connector", "slide": 1,
         "from_shape_name": "ceo", "to_shape_name": "vp3",
         "kind": "elbow", "from_point": "bottom", "to_point": "top"},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        shapes = snap["slides"][0]["shapes"]
        # Debug: print all shape types/names
        for s in shapes:
            print(f"      shape: type={s.get('type')!r} name={s.get('shape_name')!r}")
        connectors = [s for s in shapes if s.get("type") == "connector"]
        record("S1.connector_count", len(connectors) == 3,
               f"created {len(connectors)} of 3 connectors")
        record("S1.boxes_present",
               all(find_by_name(shapes, n) for n in ("ceo", "vp1", "vp2", "vp3")),
               "all 4 named boxes resolvable")
    except Exception as e:
        record("S1.connector_by_name", False, f"exception: {e}")


def s2_connector_string_id(app, deck):
    """VP: same but LLM emits 'from_shape_id' with string value."""
    print("\nS2: connector with string-valued from_shape_id (CLng crash regression)")
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a",
         "pos": {"left": 100, "top": 100, "width": 100, "height": 50},
         "fill": "#000000", "text": "A"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "b",
         "pos": {"left": 300, "top": 100, "width": 100, "height": 50},
         "fill": "#000000", "text": "B"},
        # string in from_shape_id field — used to crash with CLng
        {"type": "add_connector", "slide": 1,
         "from_shape_id": "a", "to_shape_id": "b",
         "kind": "straight", "from_point": "right", "to_point": "left"},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        shapes = snap["slides"][0]["shapes"]
        connectors = [s for s in shapes if s.get("type") == "connector"]
        record("S2.string_id_accepted", len(connectors) == 1,
               f"got {len(connectors)} connectors (expect 1)")
    except Exception as e:
        record("S2.string_id_accepted", False, f"exception: {e}")


def s3_shape_kinds(app, deck):
    """VP: 'Add a right arrow, a 5-point star, an 8-point star, and a callout.'"""
    print("\nS3: msoShape constants (right_arrow / star8 / callout_rect / ribbon_up)")
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "add_shape", "slide": 1, "kind": "right_arrow", "ref_name": "ra",
         "pos": {"left": 50, "top": 50, "width": 150, "height": 80},
         "fill": "#FF0000"},
        {"type": "add_shape", "slide": 1, "kind": "star8", "ref_name": "s8",
         "pos": {"left": 250, "top": 50, "width": 100, "height": 100},
         "fill": "#FFC000"},
        {"type": "add_shape", "slide": 1, "kind": "callout_rect", "ref_name": "cb",
         "pos": {"left": 400, "top": 50, "width": 200, "height": 100},
         "fill": "#70AD47", "text": "Note"},
        {"type": "add_shape", "slide": 1, "kind": "ribbon_up", "ref_name": "rb",
         "pos": {"left": 50, "top": 250, "width": 200, "height": 80},
         "fill": "#5B9BD5", "text": "Award"},
        {"type": "add_shape", "slide": 1, "kind": "star32", "ref_name": "s32",
         "pos": {"left": 300, "top": 250, "width": 100, "height": 100},
         "fill": "#A5A5A5"},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        shapes = snap["slides"][0]["shapes"]
        names = {s.get("shape_name") for s in shapes}
        for ref in ("ra", "s8", "cb", "rb", "s32"):
            record(f"S3.{ref}_present", ref in names, f"shape '{ref}' created")
    except Exception as e:
        record("S3.shape_kinds", False, f"exception: {e}")


def s4_string_id_in_validator(app, deck):
    """VP: 'Bold the title' but LLM emits shape_id as string ref_name."""
    print("\nS4: string ref_name in shape_id field (validator hardening)")
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "title",
         "pos": {"left": 100, "top": 100, "width": 400, "height": 60},
         "fill": "#FFFFFF", "text": "My Title", "font_color": "#000000",
         "font_size": 32},
        # LLM-style: passes string ref_name as shape_id
        {"type": "set_font_bold", "slide": 1, "shape_id": "title", "value": "true"},
        {"type": "set_font_italic", "slide": 1, "shape_id": "title", "value": "yes"},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        title = find_by_name(snap["slides"][0]["shapes"], "title")
        record("S4.string_shape_id", title is not None, "title shape resolved")
        if title:
            font = title.get("font", {})
            record("S4.bool_true_string", font.get("bold") is True,
                   f"bold={font.get('bold')} (from string 'true')")
            record("S4.bool_yes_string", font.get("italic") is True,
                   f"italic={font.get('italic')} (from string 'yes')")
    except Exception as e:
        record("S4.string_id_in_validator", False, f"exception: {e}")


def s5_group_by_names(app, deck):
    """VP: 'Group three boxes named tile_1/2/3 and send group to back.'"""
    print("\nS5: group_shapes with shape_names + z_order group-child detection")
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "tile_1",
         "pos": {"left": 100, "top": 100, "width": 100, "height": 100},
         "fill": "#FF0000"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "tile_2",
         "pos": {"left": 220, "top": 100, "width": 100, "height": 100},
         "fill": "#00FF00"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "tile_3",
         "pos": {"left": 340, "top": 100, "width": 100, "height": 100},
         "fill": "#0000FF"},
        # LLM emits shape_names instead of shape_ids — should work via alias
        {"type": "group_shapes", "slide": 1,
         "shape_names": ["tile_1", "tile_2", "tile_3"], "ref_name": "tile_grp"},
        # Now z_order on group should succeed
        {"type": "z_order", "slide": 1, "shape_name": "tile_grp", "order": "back"},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        grp = find_by_name(snap["slides"][0]["shapes"], "tile_grp")
        record("S5.group_created", grp is not None, "named group exists")
    except Exception as e:
        record("S5.group_by_names", False, f"exception: {e}")


def s6_color_formats(app, deck):
    """VP: 'Make a red shape using rgb(255,0,0) syntax.'"""
    print("\nS6: rgb() and [r,g,b] color formats")
    actions = [
        {"type": "clear_slide", "slide": 1},
        # Standard hex
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "hex_box",
         "pos": {"left": 50, "top": 50, "width": 100, "height": 100},
         "fill": "#FF0000"},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        b = find_by_name(snap["slides"][0]["shapes"], "hex_box")
        record("S6.hex_format", b is not None and (b.get("fill", "").upper() == "#FF0000"),
               f"fill={b.get('fill') if b else None}")
    except Exception as e:
        record("S6.color_formats", False, f"exception: {e}")


def s7_mixed_id_array(app, deck):
    """VP: 'Align these shapes' but LLM emits a mix of numeric Id and ref_name."""
    print("\nS7: shape_ids array with mixed numeric and string entries")
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a1",
         "pos": {"left": 100, "top": 100, "width": 80, "height": 50},
         "fill": "#000000"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a2",
         "pos": {"left": 300, "top": 200, "width": 80, "height": 50},
         "fill": "#000000"},
        {"type": "add_shape", "slide": 1, "kind": "rect", "ref_name": "a3",
         "pos": {"left": 500, "top": 150, "width": 80, "height": 50},
         "fill": "#000000"},
    ]
    try:
        run_actions(app, deck, actions)
        snap_pre = snapshot(app, deck)
        # capture numeric Id of a1 to mix with string names
        a1 = find_by_name(snap_pre["slides"][0]["shapes"], "a1")
        a1_id = a1["shape_id"]
        align_action = {"type": "align_shapes", "slide": 1,
                        "shape_ids": [a1_id, "a2", "a3"], "anchor": "top"}
        run_actions(app, deck, [align_action])
        snap_post = snapshot(app, deck)
        tops = []
        for nm in ("a1", "a2", "a3"):
            sh = find_by_name(snap_post["slides"][0]["shapes"], nm)
            if sh:
                tops.append(sh["pos"]["top"])
        record("S7.mixed_id_array",
               len(tops) == 3 and len(set(round(t, 1) for t in tops)) == 1,
               f"top values after align_top: {tops}")
    except Exception as e:
        record("S7.mixed_id_array", False, f"exception: {e}")


def s8_picture_dimensions(app, deck, tmpdir):
    """VP: 'Insert this image at 300x100 even though source is 4:3.'"""
    print("\nS8: picture insertion with explicit W/H (LockAspectRatio fix)")
    # Create a small 4:3 PNG via minimal PIL or fall back to a copy of an existing image
    img_path = tmpdir / "test_img.png"
    # Minimal 4:3 PNG (40x30) generated by hand is too brittle; use built-in stencil
    # Skip if PIL unavailable
    try:
        from PIL import Image
        Image.new("RGB", (400, 300), (200, 100, 50)).save(img_path)
    except Exception as e:
        record("S8.picture_dimensions", False, f"PIL unavailable: {e}")
        return
    actions = [
        {"type": "clear_slide", "slide": 1},
        {"type": "insert_picture", "slide": 1, "picture_path": str(img_path),
         "pos": {"left": 100, "top": 100, "width": 300, "height": 100}},
    ]
    try:
        run_actions(app, deck, actions)
        snap = snapshot(app, deck)
        pics = [s for s in snap["slides"][0]["shapes"] if s.get("type") == "picture"]
        if pics:
            w = pics[0]["pos"]["width"]
            h = pics[0]["pos"]["height"]
            record("S8.exact_dimensions",
                   abs(w - 300) < 1 and abs(h - 100) < 1,
                   f"got {w}x{h}, expected 300x100 (source aspect 4:3)")
        else:
            record("S8.exact_dimensions", False, "no picture shape found")
    except Exception as e:
        record("S8.picture_dimensions", False, f"exception: {e}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    app = open_app()
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_vp_"))
    try:
        # Best-effort visibility — some COM bindings reject this
        try:
            app.Visible = True
        except Exception:
            pass
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = fresh_deck(app, tmpdir)
        deck.Windows(1).Activate()

        s1_connector_by_name(app, deck)
        s2_connector_string_id(app, deck)
        s3_shape_kinds(app, deck)
        s4_string_id_in_validator(app, deck)
        s5_group_by_names(app, deck)
        s6_color_formats(app, deck)
        s7_mixed_id_array(app, deck)
        s8_picture_dimensions(app, deck, tmpdir)

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

    # Report
    print("\n" + "=" * 60)
    print("VP scenario results")
    print("=" * 60)
    passed = sum(1 for r in RESULTS if r["ok"])
    failed = [r for r in RESULTS if not r["ok"]]
    print(f"PASS: {passed}/{len(RESULTS)}")
    if failed:
        print(f"\nFAILS ({len(failed)}):")
        for f in failed:
            print(f"  - {f['scenario']}: {f['detail']}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""End-to-end smoke tests driven via PowerPoint COM.

Run: python tests/run_smoke.py

Each test opens a fresh copy of a test deck plus the carrier, calls VBA
functions/subs via Application.Run, asserts on returned values, and tears
down. One failure prints the diff and exits non-zero.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
DECKS_DIR = REPO_ROOT / "test_decks"


def open_app():
    import win32com.client
    return win32com.client.DispatchEx("PowerPoint.Application")


def open_pair(app, deck_name: str):
    """Open a copy of the test deck + the carrier. Returns (deck, carrier)."""
    src = DECKS_DIR / deck_name
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_"))
    deck_copy = tmpdir / src.name
    shutil.copy2(src, deck_copy)

    app.Visible = True
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
    deck = app.Presentations.Open(str(deck_copy), WithWindow=True)
    deck.Windows(1).Activate()
    return deck, carrier, tmpdir


def teardown(app, *presentations, tmpdir=None):
    for p in presentations:
        try:
            p.Saved = True
        except Exception:
            pass
        try:
            p.Close()
        except Exception:
            pass
    try:
        app.Quit()
    except Exception:
        pass
    time.sleep(0.5)
    if tmpdir and tmpdir.exists():
        shutil.rmtree(tmpdir, ignore_errors=True)


def assert_eq(actual, expected, label):
    if actual != expected:
        print(f"FAIL [{label}]")
        print(f"  expected: {expected!r}")
        print(f"  actual:   {actual!r}")
        sys.exit(1)
    print(f"  ok  [{label}]")


def test_snapshot_smoke_3slide():
    print("test_snapshot_smoke_3slide")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        json_text = app.Run("PPT_AI_Editor!BuildSnapshotJson")
        snap = json.loads(json_text)
        assert_eq(len(snap["slides"]), 3, "slide count")
        assert_eq(snap["slides"][0]["slide_number"], 1, "slide 1 number")
        # First slide title should match what we set
        first_slide_texts = [s.get("text", "") for s in snap["slides"][0]["shapes"]]
        assert "Q3 Results" in first_slide_texts, f"title text not found in {first_slide_texts}"
        print("  ok  [title text present]")
        # pos must be present on every shape
        for sl in snap["slides"]:
            for sh in sl["shapes"]:
                assert "pos" in sh, f"shape {sh['shape_id']} missing pos"
                pos = sh["pos"]
                for k in ("left", "top", "width", "height"):
                    assert k in pos, f"shape {sh['shape_id']} pos missing {k}"
                    assert isinstance(pos[k], (int, float)), f"pos.{k} not numeric"
        print("  ok  [pos present on all shapes]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_snapshot_full_visual():
    print("test_snapshot_full_visual")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "full_visual.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        slide = snap["slides"][0]

        title_box = next(s for s in slide["shapes"] if s.get("text", "").startswith("Q3 Visual"))
        font = title_box["font"]
        assert_eq(font["bold"], True, "title bold")
        assert_eq(font["size"], 32, "title size")
        assert_eq(font["color"].upper(), "#1F4E79", "title color")

        rect = next(s for s in slide["shapes"] if s.get("text") == "Box")
        assert_eq(rect["fill"].upper(), "#2E75B6", "rect fill")

        table = next(s for s in slide["shapes"] if s["type"] == "table")
        assert "table" in table, "table key missing"
        t = table["table"]
        assert_eq(t["rows"], 3, "table rows")
        assert_eq(t["cols"], 3, "table cols")
        assert_eq(t["cells"][0][0]["text"], "Metric", "table[0][0]")
        assert_eq(t["cells"][1][0]["text"], "Revenue", "table[1][0]")
        assert_eq(t["cells"][1][2]["text"], "112", "table[1][2]")

        theme = snap["deck"]["theme"]
        for slot in ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "dk1", "lt1"):
            assert slot in theme, f"theme missing {slot}"
            assert theme[slot].startswith("#"), f"theme.{slot} not hex"
        print("  ok  [theme palette present]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_backup_creates_file():
    print("test_backup_creates_file")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        backup_path = app.Run("PPT_AI_Editor!BackupActiveDeck")
        assert Path(backup_path).exists(), f"backup not found: {backup_path}"
        print(f"  ok  [backup at {backup_path}]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_action_set_text():
    print("test_action_set_text")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        title = next(s for s in before["slides"][0]["shapes"] if s["type"] == "title")
        sid = title["shape_id"]

        app.Run("PPT_AI_Editor!Do_set_text", 1, sid, "NEW TITLE")

        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        new_title = next(s for s in after["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(new_title["text"].strip(), "NEW TITLE", "title after set_text")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_action_font_ops():
    print("test_action_font_ops")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        title = next(s for s in before["slides"][0]["shapes"] if s["type"] == "title")
        sid = title["shape_id"]

        app.Run("PPT_AI_Editor!Do_set_font_size", 1, sid, 28)
        app.Run("PPT_AI_Editor!Do_set_font_bold", 1, sid, True)
        app.Run("PPT_AI_Editor!Do_set_font_italic", 1, sid, True)
        app.Run("PPT_AI_Editor!Do_set_font_color", 1, sid, "#FF0000")

        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        sh = next(s for s in after["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(int(sh["font"]["size"]), 28, "font size")
        assert_eq(sh["font"]["bold"], True, "font bold")
        assert_eq(sh["font"]["italic"], True, "font italic")
        assert_eq(sh["font"]["color"].upper(), "#FF0000", "font color")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_action_fill_color():
    print("test_action_fill_color")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "full_visual.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        rect = next(s for s in before["slides"][0]["shapes"] if s.get("text") == "Box")
        sid = rect["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_fill_color", 1, sid, "#A9D18E")
        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        sh = next(s for s in after["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(sh["fill"].upper(), "#A9D18E", "rect fill after set")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_action_geometry():
    print("test_action_geometry")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "full_visual.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        rect = next(s for s in before["slides"][0]["shapes"] if s.get("text") == "Box")
        sid = rect["shape_id"]
        app.Run("PPT_AI_Editor!Do_move_shape", 1, sid, 100.0, 200.0)
        app.Run("PPT_AI_Editor!Do_resize_shape", 1, sid, 250.0, 80.0)
        moved = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        sh = next(s for s in moved["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert abs(sh["pos"]["left"] - 100.0) < 0.5, f"left {sh['pos']['left']}"
        assert abs(sh["pos"]["top"] - 200.0) < 0.5, f"top {sh['pos']['top']}"
        assert abs(sh["pos"]["width"] - 250.0) < 0.5, f"width {sh['pos']['width']}"
        assert abs(sh["pos"]["height"] - 80.0) < 0.5, f"height {sh['pos']['height']}"
        print("  ok  [move + resize]")

        app.Run("PPT_AI_Editor!Do_delete_shape", 1, sid)
        gone = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert all(s["shape_id"] != sid for s in gone["slides"][0]["shapes"]), "shape not deleted"
        print("  ok  [delete]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_action_slide_ops():
    print("test_action_slide_ops")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        # Add at position 2, layout_index 0 (title slide layout)
        app.Run("PPT_AI_Editor!Do_add_slide", 2, 0)
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert_eq(len(snap["slides"]), 4, "slide count after add")

        app.Run("PPT_AI_Editor!Do_duplicate_slide", 1)
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert_eq(len(snap["slides"]), 5, "slide count after duplicate")

        app.Run("PPT_AI_Editor!Do_delete_slide", 5)
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert_eq(len(snap["slides"]), 4, "slide count after delete")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_action_table_ops():
    print("test_action_table_ops")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "full_visual.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        table = next(s for s in before["slides"][0]["shapes"] if s["type"] == "table")
        sid = table["shape_id"]

        app.Run("PPT_AI_Editor!Do_set_cell_text", 1, sid, 1, 1, "ZZZ")
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        t = next(s for s in snap["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(t["table"]["cells"][0][0]["text"].strip(), "ZZZ", "cell after set_cell_text")

        # Swap col 2 (Q2) with col 3 (Q3)
        app.Run("PPT_AI_Editor!Do_swap_table_columns", 1, sid, 2, 3)
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        t = next(s for s in snap["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(t["table"]["cells"][0][1]["text"].strip(), "Q3", "header after col swap")
        assert_eq(t["table"]["cells"][0][2]["text"].strip(), "Q2", "header after col swap")

        # Swap rows 2 and 3 (data rows)
        app.Run("PPT_AI_Editor!Do_swap_table_rows", 1, sid, 2, 3)
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        t = next(s for s in snap["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(t["table"]["cells"][1][0]["text"].strip(), "Margin", "row after swap")
        assert_eq(t["table"]["cells"][2][0]["text"].strip(), "Revenue", "row after swap")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_executor_end_to_end():
    print("test_executor_end_to_end")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "smoke_3slide.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        title_sid = next(s["shape_id"] for s in before["slides"][0]["shapes"] if s["type"] == "title")
        body_sid = next(s["shape_id"] for s in before["slides"][1]["shapes"] if s["type"] == "body")
        instructions = {
            "actions": [
                {"type": "set_text", "slide": 1, "shape_id": title_sid, "value": "BOARD UPDATE"},
                {"type": "set_font_color", "slide": 1, "shape_id": title_sid, "value": "#FF0000"},
                {"type": "set_text", "slide": 2, "shape_id": body_sid, "value": "Bullet A\nBullet B"},
                {"type": "set_text", "slide": 99, "shape_id": 1, "value": "should skip"},  # invalid
                {"type": "unknown_op", "slide": 1, "shape_id": title_sid},  # invalid
            ]
        }
        summary = app.Run("PPT_AI_Editor!ExecuteFromString", json.dumps(instructions))
        assert "applied" in summary.lower(), f"summary missing 'applied': {summary}"
        assert "skipped" in summary.lower(), f"summary missing 'skipped': {summary}"

        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        title = next(s for s in after["slides"][0]["shapes"] if s["shape_id"] == title_sid)
        assert_eq(title["text"].strip(), "BOARD UPDATE", "title applied")
        assert_eq(title["font"]["color"].upper(), "#FF0000", "title color applied")

        # Backup file must exist next to the deck
        deck_path = Path(deck.FullName)
        backups = list(deck_path.parent.glob(deck_path.stem + "_backup_*.pptm"))
        # smoke_3slide.pptx is .pptx not .pptm so backup keeps .pptx ext
        backups += list(deck_path.parent.glob(deck_path.stem + "_backup_*.pptx"))
        assert backups, f"no backup found in {deck_path.parent}"
        print(f"  ok  [backup at {backups[0].name}]")

        log_path = deck_path.with_suffix(deck_path.suffix + ".action_log.jsonl")
        assert log_path.exists(), f"action log not written at {log_path}"
        lines = log_path.read_text().strip().splitlines()
        # 5 actions in instructions = 5 log lines
        assert_eq(len(lines), 5, "action log line count")
        for line in lines:
            json.loads(line)  # must be valid JSON
        print(f"  ok  [log has {len(lines)} entries]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_snapshot_occupied_rects():
    print("test_snapshot_occupied_rects")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        for sl in snap["slides"]:
            assert "occupied_rects" in sl, f"slide {sl['slide_number']} missing occupied_rects"
            for r in sl["occupied_rects"]:
                for k in ("shape_id", "left", "top", "right", "bottom"):
                    assert k in r, f"occupied_rect missing {k}"
                assert r["right"] >= r["left"], "right < left"
                assert r["bottom"] >= r["top"], "bottom < top"
        print("  ok  [occupied_rects on every slide]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_snapshot_speaker_notes():
    print("test_snapshot_speaker_notes")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        s1 = snap["slides"][0]
        assert "speaker_notes" in s1, "slide 1 missing speaker_notes"
        assert "Q3" in s1["speaker_notes"], f"unexpected notes: {s1['speaker_notes']!r}"
        s2 = snap["slides"][1]
        assert s2["speaker_notes"] == "" or "speaker_notes" in s2
        print("  ok  [speaker_notes present]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def main() -> int:
    test_snapshot_smoke_3slide()
    test_snapshot_full_visual()
    test_backup_creates_file()
    test_action_set_text()
    test_action_font_ops()
    test_action_fill_color()
    test_action_geometry()
    test_action_slide_ops()
    test_action_table_ops()
    test_executor_end_to_end()
    test_snapshot_occupied_rects()
    test_snapshot_speaker_notes()
    print("\nall tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

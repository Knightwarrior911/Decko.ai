"""Smoke for granular text formatting actions (Tasks 5-20).

Tests: 11 run-level + 4 paragraph/frame + 2 cross-action = 17 total.
"""
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
DECK = REPO_ROOT / "test_decks" / "text_v3.pptx"
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"


_APP = None
_CARRIER = None


def open_app():
    """Lazy single-app singleton. Each test reuses the same PowerPoint instance
    and just reloads text_v3.pptx for isolation; this avoids COM lifecycle bugs
    where DispatchEx returns a half-released app handle between tests."""
    global _APP, _CARRIER
    if _APP is None:
        import win32com.client
        _APP = win32com.client.DispatchEx("PowerPoint.Application")
        _APP.Visible = True
        _CARRIER = _APP.Presentations.Open(str(CARRIER), WithWindow=True)
    return _APP


def fresh_deck(app):
    """Open a fresh copy of text_v3.pptx. Returns (deck, carrier, tmpdir).
    The carrier stays open across tests; only the deck cycles."""
    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_text_"))
    deck_copy = tmpdir / DECK.name
    shutil.copy2(DECK, deck_copy)
    deck = app.Presentations.Open(str(deck_copy), WithWindow=True)
    deck.Windows(1).Activate()
    return deck, _CARRIER, tmpdir


def teardown(app, *presentations, tmpdir=None):
    """Close per-test deck only; leave app + carrier alive for next test."""
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
    """Final cleanup at end of run."""
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


def assert_eq(actual, expected, label):
    if actual != expected:
        print(f"FAIL [{label}]")
        print(f"  expected: {expected!r}")
        print(f"  actual:   {actual!r}")
        sys.exit(1)
    print(f"  ok  [{label}]")


def shape_id_slide1(s):
    return s["slides"][0]["shapes"][1]["shape_id"]


# ---------- Run-level (11) ----------

def test_set_run_bold():
    print("test_set_run_bold")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_bold", 1, sid, 0, 0, True)
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][0]["font"]["bold"], True, "bold")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_italic():
    print("test_set_run_italic")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_italic", 1, sid, 0, 1, True)
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][1]["font"]["italic"], True, "italic")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_underline():
    print("test_set_run_underline")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_underline", 1, sid, 0, 0, True)
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][0]["font"]["underline"], True, "underline")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_subscript():
    print("test_set_run_subscript")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_subscript", 1, sid, 0, 0, True)
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][0]["font"]["subscript"], True, "subscript")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_superscript():
    print("test_set_run_superscript")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_superscript", 1, sid, 0, 2, True)
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][2]["font"]["superscript"], True, "superscript")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_font_color():
    print("test_set_run_font_color")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_font_color", 1, sid, 0, 0, "#FF0000")
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][0]["font"]["color"].upper(), "#FF0000", "color")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_font_size():
    print("test_set_run_font_size")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_font_size", 1, sid, 0, 1, 32)
        s2 = snap(app)
        sz = s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][1]["font"]["size"]
        assert abs(sz - 32) < 0.5, f"size {sz}"
        print(f"  ok  [size] {sz}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_font_name():
    print("test_set_run_font_name")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_font_name", 1, sid, 0, 0, "Cascadia Code")
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][0]["font"]["name"], "Cascadia Code", "font name")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_text():
    print("test_set_run_text")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_text", 1, sid, 0, 1, "DROVE 40%")
        s2 = snap(app)
        runs = s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"]
        assert_eq(runs[1]["text"], "DROVE 40%", "replaced text")
        # bold preserved on replaced run
        assert runs[1]["font"]["bold"] is True, "bold lost"
        print("  ok  [bold preserved]")
        # neighbours unchanged
        assert runs[0]["text"] == "Revenue ", runs[0]["text"]
        assert runs[2]["text"] == " in Q3", runs[2]["text"]
        print("  ok  [neighbours intact]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_run_hyperlink():
    print("test_set_run_hyperlink")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_run_hyperlink", 1, sid, 0, 1, "https://example.com")
        s2 = snap(app)
        h = s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][1]["hyperlink"]
        # PowerPoint may auto-append trailing slash to URLs without a path.
        assert h.rstrip("/") == "https://example.com", h
        print(f"  ok  [hyperlink set] {h}")
        # Clear via empty string. PowerPoint may keep the Hyperlink object but with
        # empty Address; either None or "" is acceptable as "cleared".
        app.Run("PPT_AI_Editor!Do_set_run_hyperlink", 1, sid, 0, 1, "")
        s3 = snap(app)
        cleared = s3["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][1]["hyperlink"]
        assert cleared is None or cleared == "", f"hyperlink should be cleared, got {cleared!r}"
        print(f"  ok  [hyperlink cleared] {cleared!r}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------- Paragraph + frame (4) ----------

def test_set_paragraph_alignment():
    print("test_set_paragraph_alignment")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_paragraph_alignment", 1, sid, 0, "center")
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["alignment"], "center", "alignment")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_paragraph_line_spacing():
    print("test_set_paragraph_line_spacing")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_paragraph_line_spacing", 1, sid, 0, 2.0)
        s2 = snap(app)
        ls = s2["slides"][0]["shapes"][1]["paragraphs"][0]["line_spacing"]
        assert abs(ls - 2.0) < 0.05, f"line_spacing {ls}"
        print(f"  ok  [line_spacing] {ls}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_text_vertical_align():
    print("test_set_text_vertical_align")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_text_vertical_align", 1, sid, "bottom")
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["text_frame"]["vertical_align"], "bottom", "vertical_align")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_set_text_margin():
    print("test_set_text_margin")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        app.Run("PPT_AI_Editor!Do_set_text_margin", 1, sid, 24.0, 24.0, 12.0, 12.0)
        s2 = snap(app)
        m = s2["slides"][0]["shapes"][1]["text_frame"]["margin"]
        assert abs(m["left"] - 24.0) < 0.5 and abs(m["top"] - 12.0) < 0.5, m
        print(f"  ok  [margin] {m}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


# ---------- Cross-action (2) ----------

def test_reverse_order_run_text():
    """Three set_run_text actions on runs 0, 1, 2 in one batch via ExecuteFromString.
    Without reverse-order processing, run 0 replacement shifts run 1's index.
    With it, all three apply correctly."""
    print("test_reverse_order_run_text")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        instr = json.dumps({
            "actions": [
                {"type": "set_run_text", "slide": 1, "shape_id": sid, "paragraph_index": 0, "run_index": 0, "value": "Sales "},
                {"type": "set_run_text", "slide": 1, "shape_id": sid, "paragraph_index": 0, "run_index": 1, "value": "soared 50%"},
                {"type": "set_run_text", "slide": 1, "shape_id": sid, "paragraph_index": 0, "run_index": 2, "value": " in Q4"},
            ]
        })
        result = app.Run("PPT_AI_Editor!ExecuteFromString", instr)
        assert "3 applied, 0 skipped" in result, result
        print(f"  ok  [batch] {result}")
        s2 = snap(app)
        runs = s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"]
        assert runs[0]["text"] == "Sales ", runs[0]["text"]
        assert runs[1]["text"] == "soared 50%", runs[1]["text"]
        assert runs[2]["text"] == " in Q4", runs[2]["text"]
        print("  ok  [all 3 runs correct]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_find_replace_preserves_formatting():
    """find_replace_text MUST keep mid-paragraph bold + other paragraphs intact."""
    print("test_find_replace_preserves_formatting")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        # text_v3 slide 1 paragraph 0: "Revenue grew 23% in Q3" with "grew 23%" bold (run index 1)
        before = snap(app)
        bold_run_text_before = before["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][1]["text"]
        # Replace something OUTSIDE the bold span: "Revenue " -> "Sales "
        instr = json.dumps({"actions": [
            {"type": "find_replace_text", "scope": "slide:1", "find": "Revenue", "replace": "Sales"}
        ]})
        result = app.Run("PPT_AI_Editor!ExecuteFromString", instr)
        assert "1 applied, 0 skipped" in result, result
        after = snap(app)
        para = after["slides"][0]["shapes"][1]["paragraphs"][0]
        # Look for bold run with the original bold text intact
        bold_runs = [r for r in para["runs"] if r["font"]["bold"] and "grew" in r["text"]]
        assert len(bold_runs) >= 1, f"bold span lost; runs: {para['runs']}"
        # And replacement applied
        assert any("Sales" in r["text"] for r in para["runs"]), f"replacement missing; runs: {para['runs']}"
        print(f"  ok  [formatting preserved] bold_text={bold_runs[0]['text']!r}")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def test_run_bold_idempotent():
    """Bold -> unbold -> bold again is idempotent."""
    print("test_run_bold_idempotent")
    app = open_app()
    deck, carrier, tmpdir = fresh_deck(app)
    try:
        s = snap(app)
        sid = shape_id_slide1(s)
        for v in (True, False, True):
            app.Run("PPT_AI_Editor!Do_set_run_bold", 1, sid, 0, 0, v)
        s2 = snap(app)
        assert_eq(s2["slides"][0]["shapes"][1]["paragraphs"][0]["runs"][0]["font"]["bold"], True, "final bold")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)


def main() -> int:
    test_set_run_bold()
    test_set_run_italic()
    test_set_run_underline()
    test_set_run_subscript()
    test_set_run_superscript()
    test_set_run_font_color()
    test_set_run_font_size()
    test_set_run_font_name()
    test_set_run_text()
    test_set_run_hyperlink()
    test_set_paragraph_alignment()
    test_set_paragraph_line_spacing()
    test_set_text_vertical_align()
    test_set_text_margin()
    test_reverse_order_run_text()
    test_find_replace_preserves_formatting()
    test_run_bold_idempotent()
    shutdown_app()
    print("\nall text smoke tests passed")
    return 0


if __name__ == "__main__":
    sys.exit(main())

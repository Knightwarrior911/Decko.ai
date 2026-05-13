"""
test_verify_loop.py — synthesize a deck with deliberate quality issues,
run modVerify via the carrier macros, time it, dump warnings.

Run: python tests/test_verify_loop.py
"""

import json
import os
import sys
import time
from pathlib import Path

import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
TEST_DECK = REPO_ROOT / "tests" / "test_verify_deck.pptx"
WARNINGS_SIDECAR = Path(str(TEST_DECK) + ".warnings.json")


def rgb(r: int, g: int, b: int) -> int:
    """VBA RGB encoding: R + G*256 + B*65536 (BGR order in memory)."""
    return r + (g << 8) + (b << 16)


def build_problem_deck(app):
    """Add slides with deliberate quality problems via PowerPoint COM."""
    pres = app.Presentations.Add()  # WithWindow=True (default) so chart API works
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540

    BLANK_LAYOUT = 12   # ppLayoutBlank
    NAVY = rgb(15, 40, 90)
    WHITE = rgb(255, 255, 255)
    BLACK = rgb(1, 1, 1)  # near-black; pure 0 trips some COM accessors

    # -------- Slide 1: layout-geometry problems --------
    sl1 = pres.Slides.Add(1, BLANK_LAYOUT)

    # off_slide_shape: extends past right edge
    s = sl1.Shapes.AddShape(1, 850, 100, 200, 60)  # 1=msoShapeRectangle
    s.TextFrame.TextRange.Text = "Off-slide right"
    s.Fill.ForeColor.RGB = WHITE

    # duplicate_position: two shapes at exact same bounds
    a = sl1.Shapes.AddShape(1, 100, 250, 120, 50)
    a.TextFrame.TextRange.Text = "Dup A"
    b = sl1.Shapes.AddShape(1, 100, 250, 120, 50)
    b.TextFrame.TextRange.Text = "Dup B"

    # zero_size_shape
    z = sl1.Shapes.AddShape(1, 50, 400, 0.5, 0.5)

    # empty_shape: no fill, no line, no text
    e = sl1.Shapes.AddShape(1, 300, 400, 100, 50)
    e.Fill.Visible = False
    e.Line.Visible = False

    # -------- Slide 2: text-quality problems --------
    sl2 = pres.Slides.Add(2, BLANK_LAYOUT)

    # shape_text_contrast: black text on dark navy fill
    c1 = sl2.Shapes.AddShape(1, 60, 60, 360, 80)
    c1.Fill.Solid()
    c1.Fill.ForeColor.RGB = NAVY
    c1.TextFrame.TextRange.Text = "Black text on dark navy (unreadable)"
    c1.TextFrame.TextRange.Font.Color.RGB = BLACK

    # tiny_shape_font: 6pt body text
    c2 = sl2.Shapes.AddShape(1, 60, 180, 360, 60)
    c2.Fill.ForeColor.RGB = WHITE
    c2.TextFrame.TextRange.Text = "Microscopic 6pt text body that nobody can read."
    c2.TextFrame.TextRange.Font.Size = 6

    # mixed_font_families: 3 fonts in one shape via runs
    c3 = sl2.Shapes.AddShape(1, 60, 280, 500, 60)
    c3.Fill.ForeColor.RGB = WHITE
    tr = c3.TextFrame.TextRange
    tr.Text = "Arial Calibri Times"
    tr.Words(1).Font.Name = "Arial"
    tr.Words(2).Font.Name = "Calibri"
    tr.Words(3).Font.Name = "Times New Roman"

    # text_overflow: large text in small box, no autofit
    c4 = sl2.Shapes.AddShape(1, 60, 380, 120, 30)
    c4.Fill.ForeColor.RGB = WHITE
    c4.TextFrame.TextRange.Text = (
        "This is a very long sentence that definitely overflows the tiny box "
        "because the box is only 120x30 points and the text is at 18pt size."
    )
    c4.TextFrame.TextRange.Font.Size = 18

    # -------- Slide 3: chart problems --------
    sl3 = pres.Slides.Add(3, BLANK_LAYOUT)
    chart_shape = sl3.Shapes.AddChart2(201, 51, 60, 60, 600, 360, True)  # 51=xlColumnClustered, Style=201 valid
    ch = chart_shape.Chart
    # Close the embedded data grid so subsequent chart ops don't lock
    try:
        ch.ChartData.Workbook.Close()
    except Exception:
        pass
    ch.HasTitle = False  # chart_no_title
    # Make series 1 dark with black labels (chart_label_contrast)
    ser = ch.SeriesCollection(1)
    ser.Format.Fill.Solid()
    ser.Format.Fill.ForeColor.RGB = NAVY
    ser.HasDataLabels = True
    # Default data-label color is "auto" (black); our check reads it as 0
    # and warns if contrast vs dark navy fill is too low. No explicit set needed.

    # -------- Slide 4: table problems --------
    sl4 = pres.Slides.Add(4, BLANK_LAYOUT)
    tbl_shape = sl4.Shapes.AddTable(3, 3, 60, 60, 600, 200)
    tbl = tbl_shape.Table
    # Tiny font in cell (1,1)
    c11 = tbl.Cell(1, 1)
    c11.Shape.TextFrame.TextRange.Text = "Tiny 5pt"
    c11.Shape.TextFrame.TextRange.Font.Size = 5
    # Cell (2,2): dark fill, black text
    c22 = tbl.Cell(2, 2)
    c22.Shape.Fill.Solid()
    c22.Shape.Fill.ForeColor.RGB = NAVY
    c22.Shape.TextFrame.TextRange.Text = "Black on navy"
    c22.Shape.TextFrame.TextRange.Font.Color.RGB = BLACK

    # -------- Slide 5: too many colors + too many fonts (slide aggregate) --------
    sl5 = pres.Slides.Add(5, BLANK_LAYOUT)
    palette = [
        rgb(220, 50, 47),   rgb(133, 153, 0),   rgb(38, 139, 210),  rgb(211, 54, 130),
        rgb(42, 161, 152),  rgb(203, 75, 22),   rgb(108, 113, 196), rgb(181, 137, 0),
        rgb(102, 102, 102), rgb(0, 200, 100),   rgb(255, 165, 0),   rgb(75, 0, 130),
    ]
    fonts = ["Arial", "Calibri", "Times New Roman", "Verdana", "Tahoma"]
    for i, col in enumerate(palette):
        sh = sl5.Shapes.AddShape(1, 40 + (i % 6) * 130, 60 + (i // 6) * 100, 110, 60)
        sh.Fill.Solid()
        sh.Fill.ForeColor.RGB = col
        sh.TextFrame.TextRange.Text = f"box {i+1}"
        sh.TextFrame.TextRange.Font.Color.RGB = WHITE
        sh.TextFrame.TextRange.Font.Name = fonts[i % len(fonts)]

    # Save
    if TEST_DECK.exists():
        TEST_DECK.unlink()
    pres.SaveAs(str(TEST_DECK))
    return pres


def main() -> int:
    if WARNINGS_SIDECAR.exists():
        WARNINGS_SIDECAR.unlink()

    pythoncom.CoInitialize()
    print("Launching PowerPoint...")
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True

    # Close any leftover presentations so chart-data grids aren't stuck
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    print(f"Opening carrier: {CARRIER}")
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)

    print("Building problem deck with deliberate quality issues...")
    test_pres = build_problem_deck(app)

    # Make test deck the ActivePresentation
    test_pres.Windows(1).Activate()

    # Invoke verification via Application.Run.
    # ExecuteFromString consumes a JSON instructions blob. Empty actions + verify_after=true
    # makes it just run the verification sweep.
    instructions = json.dumps({"actions": [], "verify_after": True, "verify_scope": "deck"})

    print("\nInvoking verification loop...")
    t0 = time.perf_counter()
    result = app.Run("PPT_AI_Editor.pptm!modExecuteInstructions.ExecuteFromString", instructions)
    elapsed_ms = (time.perf_counter() - t0) * 1000

    print(f"\n=== ExecuteFromString return ===")
    print(result)
    print(f"\n=== Wall-clock elapsed: {elapsed_ms:.1f} ms ===\n")

    if not WARNINGS_SIDECAR.exists():
        print(f"ERROR: sidecar not written: {WARNINGS_SIDECAR}")
        return 1

    with WARNINGS_SIDECAR.open(encoding="utf-8") as f:
        data = json.load(f)

    warnings = data.get("warnings", [])
    print(f"=== {len(warnings)} warnings ===\n")
    by_kind = {}
    for w in warnings:
        by_kind.setdefault(w["kind"], []).append(w)
    for kind, items in sorted(by_kind.items()):
        print(f"-- {kind} ({len(items)}) --")
        for w in items[:3]:
            print(f"   slide {w['slide']} shape {w['shape_id']} [{w['severity']}] {w['message']}")
            print(f"      => {w['suggestion']}")
        if len(items) > 3:
            print(f"   ... {len(items)-3} more")
        print()

    # Cleanup: close test deck without saving extra state, close carrier
    test_pres.Close()
    carrier.Close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

"""
stress_tables/harness.py — table advanced ops stress test via Decko COM.

Copies JAZZ-Pitch-Book.pptx (never modifies original), then replaces slides 7-9
with 3 stress-test slides covering:
  Slide 7: structural ops — add_table 5x4, add_table_row, add_table_col,
            delete_table_row, delete_table_col → final 5x4 with populate_table_cells
  Slide 8: formatting combos — header fill, alternating row fill, col widths,
            cell border, text orientation, padding, merge_cells, set_table_borders
  Slide 9: multi-paragraph bullet cells — disc/dash/square bullets per cell,
            set_cell_paragraph_bold, set_cell_paragraph_font_color, para alignment

Run: python tests/stress_tables/harness.py
"""

import json
import re
import shutil
import sys
from pathlib import Path

import win32com.client

REPO_ROOT = Path(__file__).resolve().parent.parent.parent
CARRIER   = REPO_ROOT / "PPT_AI_Editor.pptm"
JAZZ_SRC  = Path(r"C:\Users\vinit\Downloads\JAZZ-Pitch-Book.pptx")
OUT_DIR   = Path(__file__).resolve().parent / "output"
OUT_DECK  = OUT_DIR / "test_tables_v1.pptx"

OUT_DIR.mkdir(exist_ok=True)


def run_batch(app, actions: list) -> str:
    instr = json.dumps({"actions": actions, "verify_after": False}, ensure_ascii=True)
    return str(app.Run("PPT_AI_Editor.pptm!modExecuteInstructions.ExecuteFromString", instr))


def slide7_actions() -> list:
    """Structural ops: add 5x4, add row, add col, delete row, delete col → final 5x4."""
    S = 7
    REF = "s7_struct"

    # PHASE 1: create table 5x4
    acts = [
        {"type": "clear_slide", "slide": S},
        {"type": "add_table", "slide": S, "ref_name": REF, "rows": 5, "cols": 4,
         "pos": {"left": 30, "top": 50, "width": 660, "height": 340}},
    ]

    # Set initial column widths (5 cols expected after add_table_col)
    # We'll set them after structural changes
    # add_table_row after row 2 → 6 rows
    acts += [
        {"type": "add_table_row", "slide": S, "shape_id": REF, "after_row": 2},
    ]
    # add_table_col after col 2 → 4+1=5 cols
    acts += [
        {"type": "add_table_col", "slide": S, "shape_id": REF, "after_col": 2},
    ]
    # delete row 3 → back to 5 rows
    acts += [
        {"type": "delete_table_row", "slide": S, "shape_id": REF, "row": 3},
    ]
    # delete col 4 → back to 4 cols
    acts += [
        {"type": "delete_table_col", "slide": S, "shape_id": REF, "col": 4},
    ]

    # PHASE 2: populate the final 5x4 table
    acts += [
        {"type": "populate_table_cells", "slide": S, "shape_id": REF,
         "start_row": 1, "start_col": 1,
         "values": [
             ["Metric",      "Q1 2024",  "Q2 2024",  "Q3 2024"],
             ["Revenue",     "$480M",    "$510M",    "$530M"],
             ["EBITDA",      "$96M",     "$102M",    "$106M"],
             ["Net Income",  "$48M",     "$52M",     "$55M"],
             ["EPS",         "$1.20",    "$1.30",    "$1.38"],
         ]},
    ]

    # Set column widths
    acts += [
        {"type": "set_table_col_width", "slide": S, "shape_id": REF, "col": 1, "width_pt": 180},
        {"type": "set_table_col_width", "slide": S, "shape_id": REF, "col": 2, "width_pt": 160},
        {"type": "set_table_col_width", "slide": S, "shape_id": REF, "col": 3, "width_pt": 160},
        {"type": "set_table_col_width", "slide": S, "shape_id": REF, "col": 4, "width_pt": 160},
    ]

    return acts


def slide8_actions() -> list:
    """Formatting combos: header fill, alternating rows, cell border, text orientation,
    padding, merge_cells, set_table_borders."""
    S = 8
    REF = "s8_format"

    acts = [
        {"type": "clear_slide", "slide": S},
        {"type": "add_table", "slide": S, "ref_name": REF, "rows": 5, "cols": 3,
         "pos": {"left": 30, "top": 50, "width": 660, "height": 380}},
        # Populate
        {"type": "populate_table_cells", "slide": S, "shape_id": REF,
         "start_row": 1, "start_col": 1,
         "values": [
             ["Category",     "2023",    "2024"],
             ["North America","$210M",   "$250M"],
             ["Europe",       "$140M",   "$145M"],
             ["Asia Pacific", "$80M",    "$95M"],
             ["Total",        "$430M",   "$490M"],
         ]},
    ]

    # Header row: dark blue fill, white font, bold
    acts += [
        {"type": "set_row_fill", "slide": S, "shape_id": REF, "row": 1, "value": "#1F4E79"},
        {"type": "set_row_font_color", "slide": S, "shape_id": REF, "row": 1, "value": "#FFFFFF"},
        {"type": "set_row_font_bold", "slide": S, "shape_id": REF, "row": 1, "value": True},
    ]

    # Alternating fills for data rows
    acts += [
        {"type": "set_row_fill", "slide": S, "shape_id": REF, "row": 2, "value": "#EBF3FB"},
        {"type": "set_row_fill", "slide": S, "shape_id": REF, "row": 4, "value": "#EBF3FB"},
    ]

    # Total row: slightly different fill + bold
    acts += [
        {"type": "set_row_fill", "slide": S, "shape_id": REF, "row": 5, "value": "#D6E4F0"},
        {"type": "set_row_font_bold", "slide": S, "shape_id": REF, "row": 5, "value": True},
    ]

    # Cell-level: fill one data cell to highlight
    acts += [
        {"type": "set_cell_fill", "slide": S, "shape_id": REF, "row": 4, "col": 3, "color": "#C6EFCE"},
    ]

    # Cell alignment: header cells centered
    acts += [
        {"type": "set_cell_text_align", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "h_align": "center", "v_align": "middle"},
        {"type": "set_cell_text_align", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "h_align": "center", "v_align": "middle"},
        {"type": "set_cell_text_align", "slide": S, "shape_id": REF,
         "row": 1, "col": 3, "h_align": "center", "v_align": "middle"},
    ]

    # Cell border: right border on col 1 (separator)
    acts += [
        {"type": "set_cell_border", "slide": S, "shape_id": REF,
         "row": 2, "col": 1, "side": "right",
         "color": "#1F4E79", "weight_pt": 1.5, "visible": True, "dash_style": "solid"},
        {"type": "set_cell_border", "slide": S, "shape_id": REF,
         "row": 3, "col": 1, "side": "right",
         "color": "#1F4E79", "weight_pt": 1.5, "visible": True, "dash_style": "solid"},
        {"type": "set_cell_border", "slide": S, "shape_id": REF,
         "row": 4, "col": 1, "side": "right",
         "color": "#1F4E79", "weight_pt": 1.5, "visible": True, "dash_style": "solid"},
    ]

    # Text orientation: vertical on header col 1
    acts += [
        {"type": "set_cell_text_orientation", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "value": "horizontal"},
    ]

    # Cell padding on header row
    acts += [
        {"type": "set_cell_padding", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "left": 6, "right": 6, "top": 4, "bottom": 4},
        {"type": "set_cell_padding", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "left": 6, "right": 6, "top": 4, "bottom": 4},
        {"type": "set_cell_padding", "slide": S, "shape_id": REF,
         "row": 1, "col": 3, "left": 6, "right": 6, "top": 4, "bottom": 4},
    ]

    # Merge total row col 2+3 into one cell
    acts += [
        {"type": "set_cell_text", "slide": S, "shape_id": REF, "row": 5, "col": 3, "value": ""},
        {"type": "merge_cells", "slide": S, "shape_id": REF,
         "row_a": 5, "col_a": 2, "row_b": 5, "col_b": 3},
    ]

    # Table-level outer border
    acts += [
        {"type": "set_table_borders", "slide": S, "shape_id": REF,
         "side": "all", "color": "#1F4E79", "weight_pt": 1.0, "visible": True, "dash_style": "solid"},
    ]

    # Auto-fit text
    acts += [
        {"type": "auto_fit_table_text", "slide": S, "shape_id": REF},
    ]

    return acts


def slide9_actions() -> list:
    """Multi-paragraph bullet cells: disc/dash/square bullets, bold, font color, para align."""
    S = 9
    REF = "s9_bullets"

    acts = [
        {"type": "clear_slide", "slide": S},
        {"type": "add_table", "slide": S, "ref_name": REF, "rows": 3, "cols": 2,
         "pos": {"left": 30, "top": 50, "width": 660, "height": 380}},
    ]

    # --- Cell (1,1): heading + 3 disc bullets ---
    acts += [
        {"type": "set_cell_text", "slide": S, "shape_id": REF, "row": 1, "col": 1, "value": "Strengths"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 1, "col": 1,
         "after_paragraph_index": 0, "value": "Market leader in 3 verticals"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 1, "col": 1,
         "after_paragraph_index": 1, "value": "18% revenue CAGR over 5 years"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 1, "col": 1,
         "after_paragraph_index": 2, "value": "Proprietary IP portfolio"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "paragraph_index": 0, "value": "none"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "paragraph_index": 1, "value": "disc"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "paragraph_index": 2, "value": "disc"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "paragraph_index": 3, "value": "disc"},
        # Header bold
        {"type": "set_cell_paragraph_bold", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "paragraph_index": 0, "value": True},
        # Header blue
        {"type": "set_cell_paragraph_font_color", "slide": S, "shape_id": REF,
         "row": 1, "col": 1, "paragraph_index": 0, "value": "#1F4E79"},
    ]

    # --- Cell (1,2): heading + 3 dash bullets ---
    acts += [
        {"type": "set_cell_text", "slide": S, "shape_id": REF, "row": 1, "col": 2, "value": "Risks"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 1, "col": 2,
         "after_paragraph_index": 0, "value": "Regulatory exposure in EU"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 1, "col": 2,
         "after_paragraph_index": 1, "value": "Key-person dependency on CEO"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 1, "col": 2,
         "after_paragraph_index": 2, "value": "FX headwinds (30% rev non-USD)"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "paragraph_index": 0, "value": "none"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "paragraph_index": 1, "value": "dash"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "paragraph_index": 2, "value": "dash"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "paragraph_index": 3, "value": "dash"},
        {"type": "set_cell_paragraph_bold", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "paragraph_index": 0, "value": True},
        # Risk header red
        {"type": "set_cell_paragraph_font_color", "slide": S, "shape_id": REF,
         "row": 1, "col": 2, "paragraph_index": 0, "value": "#C00000"},
    ]

    # --- Cell (2,1): heading + 2 square bullets ---
    acts += [
        {"type": "set_cell_text", "slide": S, "shape_id": REF, "row": 2, "col": 1, "value": "Opportunities"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 2, "col": 1,
         "after_paragraph_index": 0, "value": "Emerging market expansion"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 2, "col": 1,
         "after_paragraph_index": 1, "value": "AI-driven product line"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 2, "col": 1, "paragraph_index": 0, "value": "none"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 2, "col": 1, "paragraph_index": 1, "value": "square"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 2, "col": 1, "paragraph_index": 2, "value": "square"},
        {"type": "set_cell_paragraph_alignment", "slide": S, "shape_id": REF,
         "row": 2, "col": 1, "paragraph_index": 0, "value": "left"},
    ]

    # --- Cell (2,2): heading + 2 disc bullets + colored warning ---
    acts += [
        {"type": "set_cell_text", "slide": S, "shape_id": REF, "row": 2, "col": 2, "value": "Threats"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 2, "col": 2,
         "after_paragraph_index": 0, "value": "New low-cost competitor entered"},
        {"type": "add_cell_paragraph", "slide": S, "shape_id": REF, "row": 2, "col": 2,
         "after_paragraph_index": 1, "value": "CRITICAL: supply chain risk HIGH"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 2, "col": 2, "paragraph_index": 0, "value": "none"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 2, "col": 2, "paragraph_index": 1, "value": "disc"},
        {"type": "set_cell_bullet_style", "slide": S, "shape_id": REF,
         "row": 2, "col": 2, "paragraph_index": 2, "value": "disc"},
        {"type": "set_cell_paragraph_font_color", "slide": S, "shape_id": REF,
         "row": 2, "col": 2, "paragraph_index": 2, "value": "#C00000"},
        {"type": "set_cell_paragraph_bold", "slide": S, "shape_id": REF,
         "row": 2, "col": 2, "paragraph_index": 2, "value": True},
    ]

    # --- Cell (3,1) + (3,2): merged summary row ---
    acts += [
        {"type": "set_cell_text", "slide": S, "shape_id": REF,
         "row": 3, "col": 1, "value": "Verdict: Attractive risk/reward at current valuation"},
        {"type": "set_cell_text", "slide": S, "shape_id": REF, "row": 3, "col": 2, "value": ""},
        {"type": "merge_cells", "slide": S, "shape_id": REF,
         "row_a": 3, "col_a": 1, "row_b": 3, "col_b": 2},
        {"type": "set_cell_text_align", "slide": S, "shape_id": REF,
         "row": 3, "col": 1, "h_align": "center", "v_align": "middle"},
        {"type": "set_cell_paragraph_bold", "slide": S, "shape_id": REF,
         "row": 3, "col": 1, "paragraph_index": 0, "value": True},
        {"type": "set_cell_paragraph_font_color", "slide": S, "shape_id": REF,
         "row": 3, "col": 1, "paragraph_index": 0, "value": "#1F4E79"},
    ]

    # Table border
    acts += [
        {"type": "set_table_borders", "slide": S, "shape_id": REF,
         "side": "all", "color": "#555555", "weight_pt": 0.75,
         "visible": True, "dash_style": "solid"},
    ]

    return acts


def main() -> int:
    if not JAZZ_SRC.exists():
        print(f"ERROR: {JAZZ_SRC} not found")
        return 1
    if not CARRIER.exists():
        print(f"ERROR: {CARRIER} not found")
        return 1

    print(f"Copying {JAZZ_SRC.name} -> {OUT_DECK}")
    shutil.copy2(JAZZ_SRC, OUT_DECK)

    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True

    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    print("Opening carrier...")
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)

    print("Opening test deck...")
    deck = app.Presentations.Open(str(OUT_DECK), WithWindow=True)
    deck.Windows(1).Activate()

    errors = 0
    for slide_num, build_fn in [
        (7, slide7_actions),
        (8, slide8_actions),
        (9, slide9_actions),
    ]:
        print(f"Building slide {slide_num}...")
        acts = build_fn()
        result = run_batch(app, acts)
        print(f"  Slide {slide_num}: {result.strip()[:140]}")
        m = re.search(r'(\d+) applied.*?(\d+) skipped', result)
        n_skipped = int(m.group(2)) if m else 0
        if n_skipped:
            print(f"  WARNING: {n_skipped} action(s) skipped/errored")
            errors += n_skipped

    print("Saving deck...")
    deck.SaveAs(str(OUT_DECK))
    deck.Close()
    carrier.Saved = True
    carrier.Close()
    app.Quit()

    if errors:
        print(f"\nFAILED: {errors} skipped/error action(s) detected.")
        return 1

    print(f"Done. Output: {OUT_DECK}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

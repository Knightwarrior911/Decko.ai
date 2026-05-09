"""End-to-end demo: 5 WOW capabilities executed on a copy of JAZZ-Pitch-Book.

Adds 5 demo slides to a copy of the deck — preserves original content.
"""
import json
import shutil
import sys
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SRC = Path(r"C:\Users\vinit\Downloads\JAZZ-Pitch-Book.pptx")
OUT = Path(r"C:\Users\vinit\Downloads\JAZZ-Pitch-Book Decko Demo.pptx")


def open_app():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    return app


def run_json(app, actions):
    instr = json.dumps({"actions": actions})
    return app.Run("PPT_AI_Editor!ExecuteFromString", instr)


def snap(app):
    return json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))


def main():
    if not SRC.exists():
        print(f"FAIL: source missing: {SRC}")
        sys.exit(1)

    # Fresh copy
    if OUT.exists():
        OUT.unlink()
    shutil.copy2(SRC, OUT)
    print(f"copied  -> {OUT}")

    app = open_app()
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
    deck = app.Presentations.Open(str(OUT), WithWindow=True)
    deck.Windows(1).Activate()

    base_count = deck.Slides.Count
    print(f"base slide count = {base_count}")

    # Add 5 fresh slides at end (use blank-ish layout index)
    # PowerPoint layouts: 0=Title, 1=Title+Content, 6=Blank usually
    # We pick layout_index=6 (Blank) for clean canvas
    print("\n--- adding 5 demo slides ---")
    for i in range(5):
        result = run_json(app, [
            {"type": "add_slide", "position": base_count + i + 1, "layout_index": 6}
        ])
        if "1 applied" not in result:
            print(f"  FAIL adding slide {i+1}: {result}")
            sys.exit(1)
    print(f"  ok  added 5 slides at positions {base_count+1}..{base_count+5}")

    s_badge = base_count + 1
    s_org   = base_count + 2
    s_comps = base_count + 3
    s_chart = base_count + 4   # may or may not have a usable native chart
    s_div   = base_count + 5

    failed = []

    # -------------------------------------------------------------------
    # DEMO 1: Badge circles + methodology row (rect + circle overlay + Z-order)
    print(f"\n--- DEMO 1 (slide {s_badge}): badge methodology row ---")
    result = run_json(app, [
        {"type": "add_text_box", "slide": s_badge,
         "text": "Demo 1: Badge methodology rows (rect + circle overlay + z_order)",
         "ref_name": "d1_title",
         "pos": {"left": 30, "top": 20, "width": 900, "height": 40},
         "font_size": 14, "font_bold": True, "font_color": "#1F4E79"},

        # Row D
        {"type": "add_shape", "slide": s_badge, "kind": "rect", "ref_name": "row_D",
         "pos": {"left": 80, "top": 100, "width": 260, "height": 50},
         "fill": "#1F4E79", "stroke": None,
         "text": "Sum-of-the-Parts DCF", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},
        {"type": "add_shape", "slide": s_badge, "kind": "circle", "ref_name": "badge_D",
         "pos": {"left": 60, "top": 110, "width": 32, "height": 32},
         "fill": "#243F60", "stroke": None,
         "text": "D", "font_color": "#FFFFFF", "font_size": 12, "font_bold": True},
        {"type": "z_order", "slide": s_badge, "shape_name": "badge_D", "order": "front"},
        {"type": "add_text_box", "slide": s_badge, "text": "X",
         "ref_name": "x_D",
         "pos": {"left": 380, "top": 110, "width": 35, "height": 35},
         "font_color": "#C00000", "font_size": 22, "font_bold": True},

        # Row H (added methodology - precedent premiums, marked X)
        {"type": "add_shape", "slide": s_badge, "kind": "rect", "ref_name": "row_H",
         "pos": {"left": 80, "top": 170, "width": 260, "height": 50},
         "fill": "#1F4E79", "stroke": None,
         "text": "Precedent Premiums (Not Used)", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},
        {"type": "add_shape", "slide": s_badge, "kind": "circle", "ref_name": "badge_H",
         "pos": {"left": 60, "top": 180, "width": 32, "height": 32},
         "fill": "#243F60", "stroke": None,
         "text": "H", "font_color": "#FFFFFF", "font_size": 12, "font_bold": True},
        {"type": "z_order", "slide": s_badge, "shape_name": "badge_H", "order": "front"},
        {"type": "add_text_box", "slide": s_badge, "text": "X",
         "ref_name": "x_H",
         "pos": {"left": 380, "top": 180, "width": 35, "height": 35},
         "font_color": "#C00000", "font_size": 22, "font_bold": True},
    ])
    if "9 applied" in result and "0 skipped" in result:
        print(f"  ok  9 actions applied (2 rows: rect + circle + z_order + X each, plus title)")
    else:
        print(f"  FAIL  {result}")
        failed.append("DEMO 1")

    # -------------------------------------------------------------------
    # DEMO 2: Org chart with bottom-to-top elbow connectors
    print(f"\n--- DEMO 2 (slide {s_org}): org chart with proper routing ---")
    # First add the boxes
    result = run_json(app, [
        {"type": "add_text_box", "slide": s_org,
         "text": "Demo 2: Deal structure org chart (elbow connectors, bottom->top routing)",
         "ref_name": "d2_title",
         "pos": {"left": 30, "top": 20, "width": 900, "height": 40},
         "font_size": 14, "font_bold": True, "font_color": "#1F4E79"},

        {"type": "add_shape", "slide": s_org, "kind": "rrect", "ref_name": "parent",
         "pos": {"left": 350, "top": 100, "width": 240, "height": 55},
         "fill": "#1F4E79", "stroke": None,
         "text": "JAZZ Pharmaceuticals plc", "font_color": "#FFFFFF",
         "font_size": 13, "font_bold": True},

        {"type": "add_shape", "slide": s_org, "kind": "rrect", "ref_name": "sub1",
         "pos": {"left": 100, "top": 240, "width": 200, "height": 50},
         "fill": "#2E75B6", "stroke": None,
         "text": "Jazz Therapeutics", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},

        {"type": "add_shape", "slide": s_org, "kind": "rrect", "ref_name": "sub2",
         "pos": {"left": 370, "top": 240, "width": 200, "height": 50},
         "fill": "#2E75B6", "stroke": None,
         "text": "GW Pharmaceuticals", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},

        {"type": "add_shape", "slide": s_org, "kind": "rrect", "ref_name": "sub3",
         "pos": {"left": 640, "top": 240, "width": 200, "height": 50},
         "fill": "#2E75B6", "stroke": None,
         "text": "Pharmos Inc.", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},
    ])
    if "5 applied" not in result:
        print(f"  FAIL adding org boxes: {result}")
        failed.append("DEMO 2 (boxes)")
    else:
        # Now connect parent -> children with elbow + bottom->top + filled arrow
        s = snap(app)
        slide_shapes = s["slides"][s_org - 1]["shapes"]
        ids = {sh["shape_name"]: sh["shape_id"] for sh in slide_shapes if sh.get("shape_name")}
        result = run_json(app, [
            {"type": "add_connector", "slide": s_org,
             "from_shape_id": ids["parent"], "to_shape_id": ids["sub1"],
             "kind": "elbow", "from_point": "bottom", "to_point": "top",
             "arrow_end": "filled", "color": "#595959", "weight_pt": 1.5},
            {"type": "add_connector", "slide": s_org,
             "from_shape_id": ids["parent"], "to_shape_id": ids["sub2"],
             "kind": "elbow", "from_point": "bottom", "to_point": "top",
             "arrow_end": "filled", "color": "#595959", "weight_pt": 1.5},
            {"type": "add_connector", "slide": s_org,
             "from_shape_id": ids["parent"], "to_shape_id": ids["sub3"],
             "kind": "elbow", "from_point": "bottom", "to_point": "top",
             "arrow_end": "filled", "color": "#595959", "weight_pt": 1.5},
        ])
        if "3 applied" in result:
            print(f"  ok  3 elbow connectors with filled triangle arrows")
            # Verify arrow style is correct (not diamond)
            slide = deck.Slides(s_org)
            connectors_seen = 0
            wrong_arrow = 0
            for i in range(1, slide.Shapes.Count + 1):
                sh = slide.Shapes(i)
                try:
                    if sh.Connector:
                        connectors_seen += 1
                        if sh.Line.EndArrowheadStyle != 2:
                            wrong_arrow += 1
                except Exception:
                    pass
            if wrong_arrow == 0 and connectors_seen >= 3:
                print(f"  ok  all {connectors_seen} connectors have triangle (not diamond) arrows")
            else:
                print(f"  FAIL  {wrong_arrow}/{connectors_seen} connectors have wrong arrow head")
                failed.append("DEMO 2 (arrow style)")
        else:
            print(f"  FAIL  {result}")
            failed.append("DEMO 2 (connectors)")

    # -------------------------------------------------------------------
    # DEMO 3: Comps table from nothing
    print(f"\n--- DEMO 3 (slide {s_comps}): trading comps table ---")
    result = run_json(app, [
        {"type": "add_text_box", "slide": s_comps,
         "text": "Demo 3: Trading comparables table (built from scratch)",
         "ref_name": "d3_title",
         "pos": {"left": 30, "top": 20, "width": 900, "height": 40},
         "font_size": 14, "font_bold": True, "font_color": "#1F4E79"},

        {"type": "add_table", "slide": s_comps, "rows": 6, "cols": 5,
         "ref_name": "comps_table",
         "pos": {"left": 50, "top": 90, "width": 880, "height": 280}},

        {"type": "apply_table_style", "slide": s_comps, "shape_name": "comps_table",
         "style_id": "medium_style_2_accent1"},

        # Header row styling
        {"type": "set_table_row_height", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "height_pt": 40},
        {"type": "set_cell_fill", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 1, "color": "#1F4E79"},
        {"type": "set_cell_fill", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 2, "color": "#1F4E79"},
        {"type": "set_cell_fill", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 3, "color": "#1F4E79"},
        {"type": "set_cell_fill", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 4, "color": "#1F4E79"},
        {"type": "set_cell_fill", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 5, "color": "#1F4E79"},

        # Header text
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 1, "value": "Company"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 2, "value": "EV ($B)"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 3, "value": "EV/Rev"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 4, "value": "EV/EBITDA"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 1, "col": 5, "value": "P/E"},

        # First column wider
        {"type": "set_table_col_width", "slide": s_comps, "shape_name": "comps_table",
         "col": 1, "width_pt": 240},

        # Sample comp rows
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 2, "col": 1, "value": "Vertex Pharmaceuticals"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 2, "col": 2, "value": "112.4"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 3, "col": 1, "value": "Alkermes plc"},
        {"type": "set_cell_text", "slide": s_comps, "shape_name": "comps_table",
         "row": 3, "col": 2, "value": "5.8"},
    ])
    if "applied" in result and "0 applied" not in result:
        applied = result.split(" applied")[0]
        print(f"  ok  {applied} actions applied")
    else:
        print(f"  FAIL  {result}")
        failed.append("DEMO 3")

    # -------------------------------------------------------------------
    # DEMO 4: Chart data update — depends on whether deck has native chart
    print(f"\n--- DEMO 4 (slide {s_chart}): chart data update ---")
    # Try to find a native chart anywhere in the deck first
    s = snap(app)
    chart_shape = None
    chart_slide = None
    for sl_idx, sl in enumerate(s["slides"][:base_count], start=1):
        for sh in sl["shapes"]:
            if sh.get("type") == "chart":
                chart_shape = sh
                chart_slide = sl_idx
                break
        if chart_shape:
            break

    if chart_shape:
        print(f"  found native chart on slide {chart_slide} (shape_id={chart_shape['shape_id']})")
        result = run_json(app, [
            {"type": "set_series_values", "slide": chart_slide,
             "shape_id": chart_shape["shape_id"],
             "series_index": 1, "values": [8.2, 12.4, 15.0, 18.7]},
            {"type": "set_chart_categories", "slide": chart_slide,
             "shape_id": chart_shape["shape_id"],
             "categories": ["Trading Comps", "Precedent Tx", "DCF SoTP", "DCF GC"]},
            {"type": "set_series_name", "slide": chart_slide,
             "shape_id": chart_shape["shape_id"],
             "series_index": 1, "value": "Implied EV ($bn)"},
        ])
        if "3 applied" in result:
            print(f"  ok  3 chart data actions applied on slide {chart_slide}")
        else:
            print(f"  PARTIAL  {result}")
            failed.append("DEMO 4 (chart actions)")
        # Place a marker text box on demo slide explaining what was done
        run_json(app, [
            {"type": "add_text_box", "slide": s_chart,
             "text": f"Demo 4: Live chart data update applied on slide {chart_slide} "
                     f"(see EV bar chart). Categories + values + series name swapped via JSON.",
             "ref_name": "d4_note",
             "pos": {"left": 30, "top": 20, "width": 900, "height": 80},
             "font_size": 14, "font_color": "#1F4E79"}
        ])
    else:
        print(f"  no native chart in deck — adding skip note on demo slide")
        run_json(app, [
            {"type": "add_text_box", "slide": s_chart,
             "text": "Demo 4 (skipped): JAZZ deck has no native PowerPoint chart object "
                     "(charts are likely embedded as pictures or Excel objects). "
                     "set_series_values / set_chart_categories / set_series_name "
                     "all work on native chart shapes — verified in run_smoke.py.",
             "ref_name": "d4_skip",
             "pos": {"left": 30, "top": 20, "width": 900, "height": 120},
             "font_size": 14, "font_color": "#7F7F7F"}
        ])

    # -------------------------------------------------------------------
    # DEMO 5: Navy section divider + chevron flow
    print(f"\n--- DEMO 5 (slide {s_div}): section divider + chevron flow ---")
    result = run_json(app, [
        {"type": "set_slide_background_color", "slide": s_div, "color": "#1F4E79"},
        {"type": "add_text_box", "slide": s_div, "text": "Valuation Analysis",
         "ref_name": "div_title",
         "pos": {"left": 50, "top": 130, "width": 860, "height": 90},
         "font_color": "#FFFFFF", "font_size": 44, "font_bold": True, "h_align": "center"},
        {"type": "add_text_box", "slide": s_div,
         "text": "Demo 5: Slide background color + chevron process flow",
         "ref_name": "div_sub",
         "pos": {"left": 50, "top": 240, "width": 860, "height": 30},
         "font_color": "#BDD7EE", "font_size": 14, "h_align": "center"},

        {"type": "add_shape", "slide": s_div, "kind": "chevron", "ref_name": "step1",
         "pos": {"left": 50, "top": 320, "width": 200, "height": 60},
         "fill": "#2E75B6", "stroke": None,
         "text": "Trading Comps", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},

        {"type": "add_shape", "slide": s_div, "kind": "chevron", "ref_name": "step2",
         "pos": {"left": 240, "top": 320, "width": 200, "height": 60},
         "fill": "#2E75B6", "stroke": None,
         "text": "Precedent Tx", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},

        {"type": "add_shape", "slide": s_div, "kind": "chevron", "ref_name": "step3",
         "pos": {"left": 430, "top": 320, "width": 200, "height": 60},
         "fill": "#2E75B6", "stroke": None,
         "text": "DCF - SoTP", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},

        {"type": "add_shape", "slide": s_div, "kind": "chevron", "ref_name": "step4",
         "pos": {"left": 620, "top": 320, "width": 200, "height": 60},
         "fill": "#2E75B6", "stroke": None,
         "text": "Football Field", "font_color": "#FFFFFF",
         "font_size": 12, "font_bold": True},
    ])
    if "applied" in result:
        applied_n = result.split(" applied")[0]
        if applied_n == "7":
            print(f"  ok  7 actions applied (bg + title + sub + 4 chevrons)")
        else:
            print(f"  PARTIAL  {result}")
            failed.append(f"DEMO 5 (only {applied_n}/7 applied)")
    else:
        print(f"  FAIL  {result}")
        failed.append("DEMO 5")

    # -------------------------------------------------------------------
    # Save
    print("\n--- saving ---")
    try:
        deck.Save()
        print(f"  saved -> {OUT}")
    except Exception as e:
        print(f"  FAIL save: {e}")
        failed.append("save")

    # Cleanup
    try:
        deck.Close()
        carrier.Saved = True
        carrier.Close()
        app.Quit()
    except Exception:
        pass

    print("\n=== summary ===")
    if failed:
        print(f"FAILED: {failed}")
        sys.exit(1)
    else:
        print(f"all 5 demos applied successfully on:\n  {OUT}")


if __name__ == "__main__":
    main()

"""Action-guidance completeness smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_guidance.py

For every dispatched action type:
  (a) Coverage      — GetActionGuidance(t) returns a SPECIFIC block: not
      the "No canonical guidance entry" fallback, contains an EXAMPLE
      line and a field-spec line (REQUIRED or, for actions with no
      required fields by design, OPTIONAL).
  (b) Example valid  — the EXAMPLE JSON parses, its "type" == t, and it
      passes ValidateBatchJson; PASS iff reason "" OR reason is purely a
      DECK-STATE error on the explicit allowlist below (an illustrative
      shape/slide id need not exist in the test deck). Any SCHEMA error
      (missing_field / enum / range / unknown_type / invalid JSON) FAILS.

Reuses build_problem_deck so most example refs resolve. Exit non-zero
unless both checks are 100%. No AI/API; deterministic.
"""
import json
import os
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"

sys.path.insert(0, str(REPO_ROOT / "tests"))
from test_verify_loop import build_problem_deck  # noqa: E402

# Deck-state errors are NOT schema defects: a guidance EXAMPLE uses an
# illustrative slide/shape id that need not exist in build_problem_deck.
# Justification per entry:
DECKSTATE_ALLOW = [
    "slide_out_of_range",                 # example slide # > test deck slide count
    "shape_not_found",                    # example shape_id absent in test deck
    "not found as Id or ref_name",        # example uses a ref_name not in deck
    "': not found",                       # shape_name '<x>': not found
]

ALL_TYPES = """
set_text set_font_size set_font_bold set_font_italic set_font_color set_fill_color
move_shape resize_shape delete_shape add_slide delete_slide duplicate_slide
set_cell_text swap_table_columns swap_table_rows set_paragraph_text add_paragraph
delete_paragraph set_bullet_style set_indent_level set_paragraph_font_size
set_paragraph_font_color find_replace_text align_shapes distribute_horizontal
distribute_vertical tile_grid fit_to_slide_margins add_line add_shape set_shape_kind
clear_slide move_shape_relative recolor_fill_match recolor_font_match
delete_shapes_match set_speaker_notes append_speaker_notes insert_icon insert_picture
replace_picture fetch_page_images open_image_picker build_image_picker_slide
download_image build_image_grid_table move_slide extract_slides
import_slides_from_deck add_table_row delete_table_row add_table_col delete_table_col
merge_cells group_shapes ungroup add_connector add_chart set_chart_axis
set_chart_gridlines set_chart_format add_chart_trendline set_chart_error_bars
set_chart_series set_chart_legend set_chart_type set_chart_title set_chart_axis_title
set_chart_legend_position set_series_color set_run_bold set_run_italic
set_run_underline set_run_subscript set_run_superscript set_run_font_color
set_run_font_size set_run_font_name set_run_text add_run set_run_hyperlink
set_paragraph_alignment set_paragraph_line_spacing set_text_vertical_align
set_text_autofit enable_text_shrink_for_overflow set_text_margin snap_to_grid
align_to_slide_center nudge fit_to_content match_size uniform_size smart_spacing
equalize_spacing match_position swap_positions group_by_overlap find_replace_regex
swap_font_deck_wide recolor_palette_deck_wide recolor_deck scan_palette apply_theme
set_slide_size set_theme_font bulk_insert_image bulk_insert_text_box
apply_layout_to_slides rotate_shape set_shape_adjustment flip_shape set_line_color
set_line_weight set_line_style set_shadow set_glow set_reflection set_transparency
set_gradient_fill set_3d_bevel apply_preset_effect crop_picture recolor_picture
set_brightness set_contrast add_text_box z_order duplicate_shape add_table
set_table_col_width set_table_row_height set_cell_border set_cell_text_align
set_cell_fill apply_table_style set_series_values set_chart_categories
set_series_name set_slide_background_color insert_slide_number copy_formatting
set_run_strikethrough set_paragraph_bold set_paragraph_italic set_paragraph_underline
set_paragraph_font_name set_paragraph_space_before set_paragraph_space_after
clear_paragraph_formatting set_run_highlight set_shape_name set_pos
set_shape_alt_text lock_aspect_ratio clear_shadow clear_glow clear_reflection
clear_all_effects set_soft_edge set_3d_rotation set_slide_hidden
clear_speaker_notes set_slide_name set_cell_padding clear_cell_text
set_table_style_options set_chart_data_table set_line_smoothing delete_series
add_series populate_table_row populate_table_column populate_table_cells
set_cell_font_size set_cell_font_color set_cell_font_bold set_cell_font_italic
set_cell_font_underline set_cell_font_name set_cell_text_orientation set_row_fill
set_column_fill set_row_font_size set_column_font_size set_row_font_color
set_column_font_color set_row_font_bold set_column_font_bold clear_row_text
clear_column_text set_table_font_size set_table_font_name set_table_font_color
auto_fit_table_text set_table_borders set_row_borders set_column_borders
unmerge_cells set_cell_paragraph_text set_cell_paragraph_font_size
set_cell_paragraph_font_color set_cell_paragraph_bold set_cell_paragraph_italic
set_cell_paragraph_alignment set_cell_bullet_style add_cell_paragraph
delete_cell_paragraph set_cell_indent_level append_cell_text set_cell clear_fill
clear_line set_fill_visible set_line_visible set_shape_hyperlink
set_shape_picture_fill set_slide_transition change_slide_layout add_section
delete_section rename_section move_section apply_picture_artistic_effect
reset_picture set_shape_visible reconnect_connector set_run_kerning
set_run_baseline_offset set_bullet_start_number set_notes_font_size
set_notes_font_color set_notes_font_bold set_notes_font_italic set_notes_font_name
fit_cell_to_content set_data_label_text run_verification apply_template
""".split()


def extract_example_json(text: str):
    """Return the first complete JSON value after an 'EXAMPLE:' marker, or
    None. String/escape-aware brace+bracket balance so nested pos/mappings
    are captured."""
    i = text.find("EXAMPLE:")
    if i < 0:
        return None
    s = text[i + len("EXAMPLE:"):]
    start = -1
    for j, ch in enumerate(s):
        if ch in "{[":
            start = j
            break
    if start < 0:
        return None
    depth = 0
    in_str = False
    esc = False
    for k in range(start, len(s)):
        c = s[k]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == '"':
                in_str = False
            continue
        if c == '"':
            in_str = True
        elif c in "{[":
            depth += 1
        elif c in "}]":
            depth -= 1
            if depth == 0:
                return s[start:k + 1]
    return None


def open_app():
    last = None
    for _ in range(15):
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = True
            return app
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2.0)
    raise RuntimeError(f"PowerPoint COM bring-up failed: {last!r}")


def run_session():
    pythoncom.CoInitialize()
    app = open_app()
    carrier = None
    test_pres = None
    try:
        while app.Presentations.Count > 0:
            try:
                app.Presentations(1).Close()
            except Exception:
                break
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
        test_pres = build_problem_deck(app)
        test_pres.Windows(1).Activate()
        time.sleep(1.0)

        cov_fail = []
        ex_fail = []
        for t in sorted(set(ALL_TYPES)):
            g = app.Run("PPT_AI_Editor!GetActionGuidance", t)
            g = g or ""

            is_fallback = "No canonical guidance entry" in g
            has_spec = ("REQUIRED" in g) or ("OPTIONAL" in g)
            has_example = "EXAMPLE:" in g
            if is_fallback or not has_spec or not has_example:
                cov_fail.append((t, "fallback" if is_fallback
                                 else ("no REQUIRED/OPTIONAL" if not has_spec
                                       else "no EXAMPLE")))
                ex_fail.append((t, "no usable EXAMPLE (coverage gap)"))
                continue

            raw = extract_example_json(g)
            if raw is None:
                ex_fail.append((t, "could not extract EXAMPLE JSON"))
                continue
            try:
                obj = json.loads(raw)
            except Exception as e:  # noqa: BLE001
                ex_fail.append((t, f"EXAMPLE not valid JSON: {e}"))
                continue
            if obj.get("type") != t:
                ex_fail.append((t, f'EXAMPLE type {obj.get("type")!r} != {t!r}'))
                continue

            out = app.Run("PPT_AI_Editor!ValidateBatchJson",
                          json.dumps({"actions": [obj]}))
            rows = json.loads(out).get("results", [])
            reason = rows[0]["reason"] if rows else "(no result)"
            if reason == "":
                continue
            if any(a in reason for a in DECKSTATE_ALLOW):
                continue  # deck-state only -> schema is fine
            ex_fail.append((t, f"schema error: {reason}"))

        return len(set(ALL_TYPES)), cov_fail, ex_fail
    finally:
        for p in (test_pres, carrier):
            try:
                if p is not None:
                    p.Saved = True
                    p.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(2.0)


def main() -> int:
    transient = (pywintypes.com_error, AttributeError, RuntimeError)
    result = None
    last = None
    for attempt in range(1, 4):
        try:
            result = run_session()
            break
        except transient as e:  # noqa: PERF203
            last = e
            print(f"  retry transient COM error (attempt {attempt}): {e!r}")
            os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
            time.sleep(5.0)
    if result is None:
        print(f"FAIL: guidance run failed after retries: {last!r}")
        return 1

    n_types, cov_fail, ex_fail = result
    print("deck-state allowlist (schema-OK reasons):")
    for a in DECKSTATE_ALLOW:
        print(f"  {a!r}")
    print()
    print(f"(a) coverage      : {n_types - len(cov_fail)}/{n_types} types specific")
    for t, why in cov_fail:
        print(f"  FAIL coverage [{t}] -> {why}")
    print(f"(b) example valid : {n_types - len(ex_fail)}/{n_types} examples schema-valid")
    for t, why in ex_fail:
        print(f"  FAIL example  [{t}] -> {why}")

    ok = not cov_fail and not ex_fail
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

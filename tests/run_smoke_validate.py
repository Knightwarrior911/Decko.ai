"""Action-validation smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_validate.py

Reuses build_problem_deck (known shape IDs: slide1 shapes 2-6, slide3
chart=2, slide4 table=2, >=6 slides) as the active deck, then exercises
the headless ValidateBatchJson wrapper. Three checks, all must be 100%:

  (a) Recognition  — every dispatched action type is recognized
      (reason never starts "unknown_type") for a minimal instance.
  (b) Rejection    — frozen corpus of malformed actions -> a SPECIFIC
      reason containing the expected field/constraint substring.
  (c) No false rej — frozen corpus of minimal VALID actions (referencing
      real shapes on the built deck) -> empty reason.

Exit non-zero unless all three hit 100%. No AI/API; deterministic.
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
MACRO = "PPT_AI_Editor!ValidateBatchJson"

sys.path.insert(0, str(REPO_ROOT / "tests"))
from test_verify_loop import build_problem_deck  # noqa: E402

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
build_deck_from_spec extract_spec generate_variants
""".split()


def batch(actions):
    return json.dumps({"actions": actions})


# (name, action_dict, expected_reason_substring)
REJECT = [
    ("set_text_no_value", {"type": "set_text", "slide": 1, "shape_id": 2}, "missing_field: value"),
    ("set_font_size_no_value", {"type": "set_font_size", "slide": 1, "shape_id": 2}, "missing_field: value"),
    ("resize_no_height", {"type": "resize_shape", "slide": 1, "shape_id": 2, "width": 10}, "missing_field: height"),
    ("move_no_top", {"type": "move_shape", "slide": 1, "shape_id": 2, "left": 1}, "missing_field: top"),
    ("add_chart_no_type", {"type": "add_chart", "slide": 1}, "missing_field: chart_type"),
    ("find_replace_no_scope", {"type": "find_replace_text"}, "missing_field: scope"),
    ("para_align_bad_enum", {"type": "set_paragraph_alignment", "slide": 1, "shape_id": 2,
                             "paragraph_index": 0, "value": "sideways"},
     "must be one of left, center, right, justify"),
    ("transparency_range", {"type": "set_transparency", "slide": 1, "shape_id": 2, "value": 5},
     "value: must be 0..1"),
    ("autofit_bad_mode", {"type": "set_text_autofit", "slide": 1, "shape_id": 2, "mode": "grow"},
     "mode: must be none/shrink/resize"),
    ("flip_bad_axis", {"type": "flip_shape", "slide": 1, "shape_id": 2, "axis": "x"},
     "axis: must be h or v"),
    ("unknown_type", {"type": "frobnicate"}, "unknown_type: frobnicate"),
    ("missing_type", {}, "missing_field: type"),
    ("run_font_size_neg", {"type": "set_run_font_size", "slide": 1, "shape_id": 2,
                           "paragraph_index": 0, "run_index": 0, "value": -3},
     "value: must be a positive number"),
    ("slide_size_both", {"type": "set_slide_size", "width_pt": 100, "height_pt": 100, "preset": "16:9"},
     "specify dims OR preset, not both"),
    ("slide_size_none", {"type": "set_slide_size"}, "width_pt+height_pt or preset"),
    ("recolor_deck_empty", {"type": "recolor_deck", "mappings": []}, "mappings: must be non-empty"),
    ("bevel_bad_type", {"type": "set_3d_bevel", "slide": 1, "shape_id": 2,
                        "bevel_type": "wedge", "depth_pt": 3},
     "bevel_type: must be circle/slope/cross/angle/softround"),
    ("bevel_no_depth", {"type": "set_3d_bevel", "slide": 1, "shape_id": 2,
                        "bevel_type": "circle"}, "missing_field: depth_pt"),
    ("preset_effect_range", {"type": "apply_preset_effect", "slide": 1, "shape_id": 2, "preset_index": 99},
     "preset_index: must be 1..24"),
    ("recolor_pic_bad", {"type": "recolor_picture", "slide": 1, "shape_id": 2, "color_type": "neon"},
     "color_type: must be"),
    ("nudge_bad_dir", {"type": "nudge", "slide": 1, "shape_id": 2, "direction": "x", "amount_pt": 3},
     "direction: must be l, r, u, or d"),
    ("line_style_bad", {"type": "set_line_style", "slide": 1, "shape_id": 2, "style": "wavy"},
     "style: must be solid/dash/dot/dashdot"),
    ("align_center_bad_axis", {"type": "align_to_slide_center", "slide": 1, "shape_id": 2, "axis": "diag"},
     "axis: must be h, v, or both"),
    ("zorder_bad", {"type": "z_order", "slide": 1, "shape_id": 2, "order": "sideways"},
     "order: must be front/back/forward/backward"),
    ("brightness_range", {"type": "set_brightness", "slide": 1, "shape_id": 2, "value": 9},
     "value: must be -1..1"),
    ("connector_missing_from", {"type": "add_connector", "slide": 1},
     "from_shape_id or from_shape_name"),
    ("move_slide_missing", {"type": "move_slide"}, "from or from_slide"),
    ("theme_font_missing", {"type": "set_theme_font"}, "need major or minor"),
    ("vert_align_bad", {"type": "set_text_vertical_align", "slide": 1, "shape_id": 2, "value": "upper"},
     "must be one of top, middle, bottom"),
    ("set_pos_no_dims", {"type": "set_pos", "slide": 1, "shape_id": 2},
     "at least one of left/top/width/height"),
    ("rot3d_no_axis", {"type": "set_3d_rotation", "slide": 1, "shape_id": 2},
     "at least one of x/y/z"),
    ("set_cell_no_content", {"type": "set_cell", "slide": 4, "shape_id": 2, "row": 1, "col": 1},
     "pass at least one of"),
    ("tablestyle_no_opt", {"type": "set_table_style_options", "slide": 4, "shape_id": 2},
     "pass at least one of header_row"),
    ("shape_hyperlink_bad", {"type": "set_shape_hyperlink", "slide": 1, "shape_id": 2, "value": "ftp://x"},
     "value: invalid hyperlink URL"),
    ("smart_spacing_bad_axis", {"type": "smart_spacing", "slide": 1, "shape_ids": [2, 3],
                                "gap_pt": 5, "axis": "z"}, "axis: must be h or v"),
    ("equalize_bad_axis", {"type": "equalize_spacing", "slide": 1, "shape_ids": [2, 3], "axis": "z"},
     "axis: must be h or v"),
    ("match_pos_bad_edge", {"type": "match_position", "slide": 1, "ref_shape_id": 2,
                            "target_shape_id": 3, "edge": "corner"}, "edge: must be"),
    ("recolor_palette_bad", {"type": "recolor_palette_deck_wide", "from_hex": "#000",
                             "to_hex": "#fff", "target": "glow"}, "target: must be fill/font/both"),
    ("line_weight_zero", {"type": "set_line_weight", "slide": 1, "shape_id": 2, "weight_pt": 0},
     "weight_pt: must be > 0"),
    ("uniform_size_zero", {"type": "uniform_size", "slide": 1, "shape_ids": [2, 3],
                           "width_pt": 0, "height_pt": 5}, "width_pt/height_pt: must be > 0"),
    ("scan_palette_bad_scope", {"type": "scan_palette", "scope": "page"},
     "scope: must be 'deck' or 'slide:N'"),
    ("run_font_name_empty", {"type": "set_run_font_name", "slide": 1, "shape_id": 2,
                             "paragraph_index": 0, "run_index": 0, "value": "  "},
     "value: empty font name"),
    ("apply_theme_empty", {"type": "apply_theme", "theme_path": ""}, "theme_path: empty"),
    ("apply_template_bad_name", {"type": "apply_template", "template": "bogus",
                                 "content": {}}, "template: must be one of"),
    ("apply_template_missing_slot", {"type": "apply_template", "template": "title",
                                     "content": {"title": "x"}}, "content.subtitle: required"),
    ("build_spec_empty_deck", {"type": "build_deck_from_spec",
                               "spec": {"deck": []}}, "spec.deck: must be a non-empty array"),
    ("build_spec_no_spec", {"type": "build_deck_from_spec"}, "missing_field: spec"),
    ("variants_bad_n", {"type": "generate_variants", "template": "title",
                        "content": {"title": "a"}, "n": 9}, "n: must be an integer 2..6"),
]

# (name, action_dict)  -- expected reason == ""
VALID = [
    ("set_text", {"type": "set_text", "slide": 1, "shape_id": 2, "value": "Hi"}),
    ("set_font_size", {"type": "set_font_size", "slide": 1, "shape_id": 2, "value": 18}),
    ("set_font_bold", {"type": "set_font_bold", "slide": 1, "shape_id": 2, "value": True}),
    ("set_font_italic", {"type": "set_font_italic", "slide": 1, "shape_id": 2, "value": False}),
    ("set_font_color", {"type": "set_font_color", "slide": 1, "shape_id": 2, "value": "#112233"}),
    ("set_fill_color", {"type": "set_fill_color", "slide": 1, "shape_id": 2, "value": "#445566"}),
    ("move_shape", {"type": "move_shape", "slide": 1, "shape_id": 2, "left": 10, "top": 20}),
    ("resize_shape", {"type": "resize_shape", "slide": 1, "shape_id": 2, "width": 100, "height": 50}),
    ("delete_shape", {"type": "delete_shape", "slide": 1, "shape_id": 3}),
    ("set_pos", {"type": "set_pos", "slide": 1, "shape_id": 2, "left": 5}),
    ("rotate_shape", {"type": "rotate_shape", "slide": 1, "shape_id": 2, "degrees": 45}),
    ("set_shape_name", {"type": "set_shape_name", "slide": 1, "shape_id": 2, "value": "Box"}),
    ("set_shape_alt_text", {"type": "set_shape_alt_text", "slide": 1, "shape_id": 2, "value": "alt"}),
    ("lock_aspect_ratio", {"type": "lock_aspect_ratio", "slide": 1, "shape_id": 2, "value": True}),
    ("set_transparency", {"type": "set_transparency", "slide": 1, "shape_id": 2, "value": 0.5}),
    ("set_line_color", {"type": "set_line_color", "slide": 1, "shape_id": 2, "value": "#000000"}),
    ("set_line_weight", {"type": "set_line_weight", "slide": 1, "shape_id": 2, "weight_pt": 2}),
    ("set_line_style", {"type": "set_line_style", "slide": 1, "shape_id": 2, "style": "dash"}),
    ("set_3d_rotation", {"type": "set_3d_rotation", "slide": 1, "shape_id": 2, "x": 10}),
    ("clear_all_effects", {"type": "clear_all_effects", "slide": 1, "shape_id": 2}),
    ("z_order", {"type": "z_order", "slide": 1, "shape_id": 2, "order": "front"}),
    ("duplicate_shape", {"type": "duplicate_shape", "slide": 1, "shape_id": 2, "left": 30, "top": 30}),
    ("set_paragraph_alignment", {"type": "set_paragraph_alignment", "slide": 1, "shape_id": 2,
                                 "paragraph_index": 0, "value": "center"}),
    ("set_paragraph_bold", {"type": "set_paragraph_bold", "slide": 1, "shape_id": 2,
                            "paragraph_index": 0, "value": True}),
    ("set_run_bold", {"type": "set_run_bold", "slide": 1, "shape_id": 2,
                      "paragraph_index": 0, "run_index": 0, "value": True}),
    ("set_cell_text", {"type": "set_cell_text", "slide": 4, "shape_id": 2,
                       "row": 1, "col": 1, "value": "X"}),
    ("add_table_row", {"type": "add_table_row", "slide": 4, "shape_id": 2, "after_row": 1}),
    ("delete_table_col", {"type": "delete_table_col", "slide": 4, "shape_id": 2, "col": 1}),
    ("set_cell_fill", {"type": "set_cell_fill", "slide": 4, "shape_id": 2,
                       "row": 1, "col": 1, "color": "#abcdef"}),
    ("merge_cells", {"type": "merge_cells", "slide": 4, "shape_id": 2,
                     "row_a": 1, "col_a": 1, "row_b": 1, "col_b": 2}),
    ("set_table_col_width", {"type": "set_table_col_width", "slide": 4, "shape_id": 2,
                             "col": 1, "width_pt": 80}),
    ("set_chart_title", {"type": "set_chart_title", "slide": 3, "shape_id": 2, "value": "T"}),
    ("set_chart_type", {"type": "set_chart_type", "slide": 3, "shape_id": 2, "value": "bar"}),
    ("set_chart_legend_position", {"type": "set_chart_legend_position", "slide": 3,
                                   "shape_id": 2, "value": "bottom"}),
    ("find_replace_text", {"type": "find_replace_text", "scope": "deck", "find": "a", "replace": "b"}),
    ("move_slide", {"type": "move_slide", "from": 1, "to": 2}),
    ("add_section", {"type": "add_section", "before_slide": 1, "name": "Intro"}),
    ("set_slide_size", {"type": "set_slide_size", "preset": "16:9"}),
    ("recolor_deck", {"type": "recolor_deck", "mappings": [{"from": "#000000", "to": "#ffffff"}]}),
    ("swap_font_deck_wide", {"type": "swap_font_deck_wide", "from_name": "Arial", "to_name": "Calibri"}),
    ("run_verification", {"type": "run_verification"}),
    ("scan_palette", {"type": "scan_palette", "scope": "deck"}),
    ("set_theme_font", {"type": "set_theme_font", "major": "Arial"}),
    ("set_slide_hidden", {"type": "set_slide_hidden", "slide": 1, "value": True}),
    ("set_speaker_notes", {"type": "set_speaker_notes", "slide": 1, "value": "notes"}),
    ("set_3d_bevel", {"type": "set_3d_bevel", "slide": 1, "shape_id": 2,
                      "bevel_type": "circle", "depth_pt": 6}),
    ("apply_template", {"type": "apply_template", "template": "title",
                        "content": {"title": "a", "subtitle": "b"}}),
    ("build_deck_from_spec", {"type": "build_deck_from_spec",
                              "spec": {"deck": [{"template": "title",
                                                 "content": {"title": "a", "subtitle": "b"}}]}}),
    ("extract_spec", {"type": "extract_spec"}),
    ("generate_variants", {"type": "generate_variants", "template": "title",
                           "content": {"title": "a"}, "n": 3}),
]


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


def reason_of(app, action):
    out = app.Run(MACRO, batch([action]))
    rows = json.loads(out).get("results", [])
    return rows[0]["reason"] if rows else "(no result)"


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

        # (a) recognition
        cov_fail = []
        for t in sorted(set(ALL_TYPES)):
            r = reason_of(app, {"type": t})
            if r.lower().startswith("unknown_type") or r.startswith("ERROR: validator raised"):
                cov_fail.append((t, r))

        # (b) rejection
        rej_fail = []
        for name, act, sub in REJECT:
            r = reason_of(app, act)
            if not r or sub not in r or r.startswith("ERROR: validator raised"):
                rej_fail.append((name, sub, r))

        # (c) no false reject
        val_fail = []
        for name, act in VALID:
            r = reason_of(app, act)
            if r != "":
                val_fail.append((name, r))

        n_types = len(set(ALL_TYPES))
        return (n_types, cov_fail, len(REJECT), rej_fail, len(VALID), val_fail)
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
        print(f"FAIL: validate run failed after retries: {last!r}")
        return 1

    n_types, cov_fail, n_rej, rej_fail, n_val, val_fail = result
    print(f"(a) recognition : {n_types - len(cov_fail)}/{n_types} types recognized")
    for t, r in cov_fail:
        print(f"  FAIL recognition [{t}] -> {r!r}")
    print(f"(b) rejection   : {n_rej - len(rej_fail)}/{n_rej} malformed -> specific reason")
    for name, sub, r in rej_fail:
        print(f"  FAIL reject [{name}] want substr {sub!r} got {r!r}")
    print(f"(c) no false rej: {n_val - len(val_fail)}/{n_val} valid -> empty reason")
    for name, r in val_fail:
        print(f"  FAIL falsereject [{name}] got {r!r}")

    ok = not cov_fail and not rej_fail and not val_fail
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

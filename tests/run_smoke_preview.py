"""Plan-preview describer smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_preview.py

Checks BuildActionPlanSummary (headless, read-only describer):
  (a) Coverage  — every action type the executor dispatches yields a
      SPECIFIC sentence (never UNKNOWN/ERROR/empty/generic).
  (b) Exactness — a frozen corpus of (batch JSON -> expected summary)
      pairs match byte-for-byte (newlines normalized to \\n).

Exit non-zero unless 100% on both. No AI/API; fully deterministic.
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
MACRO = "PPT_AI_Editor!BuildActionPlanSummary"

# Every action type the executor dispatch (Select Case) handles.
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
capture_template list_templates delete_template rename_template
""".split()


def batch(actions):
    return json.dumps({"actions": actions})


# (name, raw_json_or_text, expected_summary)  -- newlines normalized to \n
CORPUS = [
    ("set_text",
     batch([{"type": "set_text", "slide": 1, "shape_id": 3, "text": "Hello"}]),
     '1. Slide 1: set text on shape #3 -> "Hello"'),
    ("set_font_size_named",
     batch([{"type": "set_font_size", "slide": 2, "shape_name": "Title", "size": 40}]),
     '1. Slide 2: set font size on "Title" -> 40pt'),
    ("set_font_bold_true",
     batch([{"type": "set_font_bold", "slide": 1, "shape_id": 2, "value": True}]),
     "1. Slide 1: set font bold on shape #2 -> true"),
    ("set_fill_color",
     batch([{"type": "set_fill_color", "slide": 1, "shape_id": 5, "value": "#FF0000"}]),
     "1. Slide 1: set fill color on shape #5 -> #FF0000"),
    ("move_shape",
     batch([{"type": "move_shape", "slide": 1, "shape_id": 4, "left": 100, "top": 200}]),
     "1. Slide 1: move shape on shape #4 -> left 100, top 200"),
    ("resize_shape",
     batch([{"type": "resize_shape", "slide": 1, "shape_id": 4, "width": 300, "height": 150}]),
     "1. Slide 1: resize on shape #4 -> 300 x 150"),
    ("delete_shape",
     batch([{"type": "delete_shape", "slide": 3, "shape_id": 7}]),
     "1. Slide 3: delete shape on shape #7"),
    ("add_slide_layout",
     batch([{"type": "add_slide", "layout": "Title and Content"}]),
     "1. add slide -> Title and Content"),
    ("find_replace_text",
     batch([{"type": "find_replace_text", "find": "foo", "replace": "bar"}]),
     '1. find replace text -> replace "foo" with "bar"'),
    ("run_verification",
     batch([{"type": "run_verification"}]),
     "1. run the slide-quality verification sweep"),
    ("unknown_type",
     batch([{"type": "frobnicate", "slide": 1, "shape_id": 2}]),
     "1. UNKNOWN ACTION frobnicate"),
    ("missing_type",
     batch([{"slide": 1}]),
     "1. UNKNOWN ACTION (missing type)"),
    ("multi_action_order",
     batch([{"type": "set_text", "slide": 1, "shape_id": 1, "text": "A"},
            {"type": "set_font_bold", "slide": 1, "shape_id": 1, "value": False}]),
     '1. Slide 1: set text on shape #1 -> "A"\n'
     "2. Slide 1: set font bold on shape #1 -> false"),
    ("messy_fence_prose_comma",
     "```json\nSure:\n"
     '{"actions":[{"type":"set_text","slide":1,"shape_id":2,"text":"Hi"},]}'
     "\n```",
     '1. Slide 1: set text on shape #2 -> "Hi"'),
    ("no_actions",
     batch([]),
     "(no actions)"),
    ("missing_actions_key",
     '{"foo":1}',
     "ERROR: missing top-level 'actions' array"),
    ("set_paragraph_alignment",
     batch([{"type": "set_paragraph_alignment", "slide": 1, "shape_id": 2, "value": "center"}]),
     "1. Slide 1: set paragraph alignment on shape #2 -> center"),
    ("set_cell_text",
     batch([{"type": "set_cell_text", "slide": 1, "shape_id": 3, "row": 2, "col": 1, "text": "X"}]),
     "1. Slide 1: set cell text in cell (2,1) on shape #3"),
    ("add_chart_no_salient",
     batch([{"type": "add_chart", "slide": 2, "chart_type": "columnclustered"}]),
     "1. Slide 2: add chart"),
    ("add_section",
     batch([{"type": "add_section", "name": "Intro"}]),
     '1. add section "Intro"'),
    ("recolor_deck",
     batch([{"type": "recolor_deck", "value": "#123456"}]),
     "1. recolor deck -> #123456 (deck-wide)"),
    ("set_speaker_notes",
     batch([{"type": "set_speaker_notes", "slide": 4, "text": "hi"}]),
     "1. Slide 4: set speaker notes"),
    ("align_shapes",
     batch([{"type": "align_shapes", "slide": 1, "value": "left"}]),
     "1. Slide 1: align shapes on selected shapes"),
    ("rotate_shape",
     batch([{"type": "rotate_shape", "slide": 1, "shape_id": 2, "angle": 45}]),
     "1. Slide 1: rotate shape on shape #2 -> angle 45"),
    ("set_run_bold",
     batch([{"type": "set_run_bold", "slide": 1, "shape_id": 2,
             "paragraph_index": 0, "run_index": 1, "value": True}]),
     "1. Slide 1: set run bold on shape #2 -> true"),
    ("set_font_color",
     batch([{"type": "set_font_color", "slide": 1, "shape_id": 2, "value": "#00FF00"}]),
     "1. Slide 1: set font color on shape #2 -> #00FF00"),
    ("apply_template",
     batch([{"type": "apply_template", "template": "title",
             "content": {"title": "X"}}]),
     '1. apply the "title" slide template'),
    ("build_deck_from_spec",
     batch([{"type": "build_deck_from_spec",
             "spec": {"deck": [{"template": "title", "content": {}}]}}]),
     "1. build the deck from the provided spec"),
    ("extract_spec",
     batch([{"type": "extract_spec"}]),
     "1. extract the live deck into a spec (.spec.json)"),
    ("generate_variants",
     batch([{"type": "generate_variants", "template": "title",
             "content": {}, "n": 3}]),
     '1. generate 3 layout variants of the "title" template'),
    ("capture_template",
     batch([{"type": "capture_template", "name": "my_kpi"}]),
     '1. capture this slide as reusable template "my_kpi"'),
    ("list_templates",
     batch([{"type": "list_templates"}]),
     "1. list your captured templates"),
    ("delete_template",
     batch([{"type": "delete_template", "name": "old"}]),
     '1. delete captured template "old"'),
    ("rename_template",
     batch([{"type": "rename_template", "from": "a", "to": "b"}]),
     '1. rename captured template "a" -> "b"'),
]


def norm(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\r", "\n")


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
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)

        # (a) coverage
        cov_pass = 0
        cov_fail = []
        for t in sorted(set(ALL_TYPES)):
            raw = batch([{"type": t, "slide": 1, "shape_id": 1}])
            out = norm(app.Run(MACRO, raw))
            line = out.split("\n", 1)[0]
            body = line[3:] if line.startswith(("1. ",)) else line
            ok = (line.startswith("1. ")
                  and not body.startswith("UNKNOWN ACTION")
                  and not body.startswith("ERROR")
                  and " " in body.strip()
                  and len(body.strip()) > 0)
            if ok:
                cov_pass += 1
            else:
                cov_fail.append((t, line))

        # (b) exactness
        ex_pass = 0
        ex_fail = []
        for name, raw, expected in CORPUS:
            out = norm(app.Run(MACRO, raw))
            if out == norm(expected):
                ex_pass += 1
            else:
                ex_fail.append((name, expected, out))

        return cov_pass, cov_fail, ex_pass, ex_fail
    finally:
        try:
            if carrier is not None:
                carrier.Saved = True
                carrier.Close()
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
        print(f"FAIL: preview run failed after retries: {last!r}")
        return 1

    cov_pass, cov_fail, ex_pass, ex_fail = result
    n_types = len(set(ALL_TYPES))
    n_corpus = len(CORPUS)

    print(f"coverage : {cov_pass}/{n_types} action types -> specific sentence")
    for t, line in cov_fail:
        print(f"  FAIL coverage [{t}] -> {line!r}")
    print(f"exactness: {ex_pass}/{n_corpus} corpus pairs exact match")
    for name, expected, out in ex_fail:
        print(f"  FAIL exact [{name}]")
        print(f"    expected: {expected!r}")
        print(f"    actual:   {out!r}")

    ok = (cov_pass == n_types and ex_pass == n_corpus)
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

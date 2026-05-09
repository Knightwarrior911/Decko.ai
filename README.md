# Decko.ai

VBA-based PowerPoint editor driven by natural language via an external LLM.

The user copies a JSON snapshot of the active deck into their LLM tool,
describes a desired change, copies the LLM's instructions JSON back, and
clicks Apply. Three macros, three UserForms, no API calls from PowerPoint.

See `docs/specs/2026-05-08-ppt-ai-editor-design.md` (Phase 1 design) and
`docs/specs/2026-05-08-ppt-ai-editor-phase2-design.md` (Phase 2 design).

## One-time developer setup

1. Install Python 3.10+ and run `pip install -r requirements.txt`.
2. In PowerPoint: File → Options → Trust Center → Trust Center Settings →
   Macro Settings → check **Trust access to the VBA project object model**.
3. From the repo root, run `python tools/build_carrier.py` to bootstrap
   `PPT_AI_Editor.pptm` (idempotent — skips if it already exists).
4. Run `python tools/build_forms.py` to author the two UserForms inside the
   carrier and export their `.frm`/`.frx` source files to `src/`.
5. Run `python update_macros.py` to sync `src/` modules into the carrier.

## Daily use

1. Open the deck you want to edit, plus `PPT_AI_Editor.pptm`.
2. Press Alt+F8 → run `ExportSnapshot`. The form opens with a JSON snapshot.
3. Click **Copy snapshot + prompt template**. Your clipboard now holds a
   prompt block ending with `[REPLACE THIS LINE WITH YOUR REQUEST]`.
4. Paste into your LLM tool. Replace the placeholder line with what you want
   changed in the deck. Submit.
5. The LLM returns an instructions JSON. Copy it.
6. Press Alt+F8 → run `ExecuteInstructions`. Paste into the textbox, click
   **Parse**. Review the action list — invalid rows are tagged.
7. Click **Apply**. The deck is backed up first, then valid actions are
   executed in order. A summary line shows applied / skipped counts and the
   paths to the backup file and JSONL action log.

## Action types

56 actions total across 9 modules.

### Core shape + slide (`modActions.bas`, 17)

| Action | Effect |
|---|---|
| `set_text` | Replace all text in a shape. |
| `set_font_size` | Set font size (pt) on the entire text frame. |
| `set_font_bold` | True / False. |
| `set_font_italic` | True / False. |
| `set_font_color` | `#RRGGBB`. |
| `set_fill_color` | `#RRGGBB` solid fill. |
| `move_shape` | Set `left` / `top` (pt). |
| `resize_shape` | Set `width` / `height` (pt). |
| `delete_shape` | Remove a shape. |
| `add_slide` | New slide at `position` using `layout_index`. |
| `delete_slide` | Remove a slide by 1-based number. |
| `duplicate_slide` | Clone a slide; copy lands immediately after source. |
| `set_cell_text` | Set text of a table cell `(row, col)`. |
| `swap_table_columns` | Swap two columns. |
| `swap_table_rows` | Swap two rows. |
| `set_speaker_notes` | Replace speaker notes on a slide. |
| `append_speaker_notes` | Append text to existing speaker notes. |

### Granular text (`modActionsText.bas`, 8)

| Action | Effect |
|---|---|
| `set_paragraph_text` | Replace text of paragraph N inside a shape. |
| `add_paragraph` | Insert a new paragraph at index. |
| `delete_paragraph` | Remove paragraph at index. |
| `set_bullet_style` | Set bullet type (none/bullet/number) for a paragraph. |
| `set_indent_level` | Set indent level (0-4) for a paragraph. |
| `set_paragraph_font_size` | Per-paragraph font size override. |
| `set_paragraph_font_color` | Per-paragraph color override. |
| `find_replace_text` | Scoped find/replace; scope = `deck` or `slide:N`. |

### Layout + alignment (`modActionsLayout.bas`, 13)

| Action | Effect |
|---|---|
| `align_shapes` | Align multiple shapes (left/right/top/bottom/hcenter/vcenter). |
| `distribute_horizontal` | Even horizontal spacing across selected shapes. |
| `distribute_vertical` | Even vertical spacing across selected shapes. |
| `tile_grid` | Arrange shapes into N-column grid with gap. |
| `fit_to_slide_margins` | Fit a shape inside slide minus margin. |
| `add_line` | Add a line connector between (x1,y1) and (x2,y2). |
| `add_shape` | Add an `msoAutoShape` (rect, oval, etc.) at given pos. |
| `set_shape_kind` | Change autoshape kind on existing shape. |
| `clear_slide` | Delete all shapes except those in `keep_ids`. |
| `move_shape_relative` | Nudge a shape by (dx, dy). |
| `recolor_fill_match` | Replace one fill color with another in scope. |
| `recolor_font_match` | Replace one font color with another in scope. |
| `delete_shapes_match` | Delete shapes matching kind/text filters in scope. |

### Tables (`modActionsTable.bas`, 5)

| Action | Effect |
|---|---|
| `add_table_row` | Insert row after row N. |
| `delete_table_row` | Remove row N. |
| `add_table_col` | Insert column after column N. |
| `delete_table_col` | Remove column N. |
| `merge_cells` | Merge cell range `(r1,c1)-(r2,c2)`. |

### Charts (`modActionsChart.bas`, 5)

| Action | Effect |
|---|---|
| `set_chart_type` | Change chart type (column/bar/line/pie/etc.). |
| `set_chart_title` | Set chart title text. |
| `set_chart_axis_title` | Set axis title (category/value). |
| `set_chart_legend_position` | Position legend (top/right/bottom/left/none). |
| `set_series_color` | Color a chart series by index. |

### Images (`modActionsImage.bas`, 2)

| Action | Effect |
|---|---|
| `insert_picture` | Insert image at position. |
| `replace_picture` | Replace existing picture, preserving frame. |

### Connectors + groups (`modActionsConnector.bas`, `modActionsGroup.bas`, 3)

| Action | Effect |
|---|---|
| `add_connector` | Add elbow/straight connector between two shapes. |
| `group_shapes` | Group shapes into one. |
| `ungroup` | Ungroup a group shape. |

### Slide structure (`modActionsSlide.bas`, 3)

| Action | Effect |
|---|---|
| `move_slide` | Reorder slide from index A to B. |
| `extract_slides` | Export selected slides to a new .pptx file. |
| `import_slides_from_deck` | Import slides from another deck at position. |

## Safety

- Every Apply runs an auto-backup first: `<deck>_backup_<timestamp>.<ext>`
  is written next to the deck before any edit.
- Every executed (or skipped) action is appended to
  `<deck>.action_log.jsonl` as one JSON line per action.
- One bad action does not abort the batch — it is logged and skipped, and
  the rest continue.
- Top-level invalid JSON aborts before any change.

## Editing the macros

The source of truth is in `src/`. To change behavior:

1. Edit a `.bas` file (or rebuild a `.frm` via `tools/build_forms.py`).
2. Run `python update_macros.py` — re-imports modules into the carrier.
3. Reopen the carrier in PowerPoint to test.

UserForms are built programmatically by `tools/build_forms.py`. To change
their layout or controls, edit that script and re-run it; the script
exports the resulting `.frm`/`.frx` to `src/` for `update_macros.py` to
pick up.

## Tests

```bash
python tests/make_test_decks.py     # regenerate test decks
python update_macros.py             # ensure carrier matches src/
python tests/run_smoke.py           # end-to-end COM-driven smoke
```

## Macros (Alt+F8)

| Macro | UserForm | Purpose |
|---|---|---|
| `ExportSnapshot` | `frmExport` | Build deck snapshot JSON; copy snapshot + prompt to clipboard. |
| `ExecuteInstructions` | `frmExecute` | Paste instructions JSON, parse, review, apply. |
| `ImportSlides` | `frmImportSlides` | Import slides from another deck at a given position. |

## Files

```
PPT_AI_Editor.pptm                ← carrier (regenerated from src/)
src/
  modUI.bas                       ← public macros (Alt+F8 entry points)
  modExportSnapshot.bas           ← snapshot JSON builder
  modExecuteInstructions.bas      ← parse + dispatch + apply pipeline
  modBackup.bas                   ← auto-backup helper
  modJSON.bas                     ← JSON parser/encoder
  modActions.bas                  ← core 17 actions (text/font/shape/slide/notes)
  modActionsText.bas              ← granular text actions (8)
  modActionsLayout.bas            ← layout/align/distribute/recolor (13)
  modActionsTable.bas             ← table row/col/merge (5)
  modActionsChart.bas             ← chart type/title/axis/legend/series (5)
  modActionsImage.bas             ← insert/replace picture (2)
  modActionsConnector.bas         ← connector (1)
  modActionsGroup.bas             ← group/ungroup (2)
  modActionsSlide.bas             ← move/extract/import slides (3)
  frmExport.frm/.frx              ← snapshot UserForm
  frmExecute.frm/.frx             ← instructions UserForm
  frmImportSlides.frm/.frx        ← import UserForm
update_macros.py                  ← sync src/ → carrier
tools/build_carrier.py            ← bootstrap empty carrier
tools/build_forms.py              ← rebuild UserForms in carrier; export .frm/.frx to src/
tools/inspect_form.py             ← inspect UserForm controls via COM
tools/screenshot_forms.py         ← DPI-aware screenshot of each UserForm
tools/rebuild_import_slides.py    ← escape hatch for frmImportSlides rebuild
tools/precheck_carrier.py         ← carrier sanity check
tests/                            ← smoke harness + deck generator
test_decks/                       ← deterministic test inputs
docs/specs/                       ← design specs (Phase 1 + Phase 2)
docs/superpowers/plans/           ← implementation plans
```

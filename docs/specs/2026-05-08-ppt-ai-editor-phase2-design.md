# PPT AI Editor — Phase 2 Design Spec

**Date:** 2026-05-08
**Status:** Approved (pending user review of this document)
**Builds on:** `docs/specs/2026-05-08-ppt-ai-editor-design.md` (V1, 15 actions)

## 1. Goals & non-goals

### Goals

Expand action surface from 15 to ~70 so that a senior reviewer (VP / MD level in investment banking) can issue PowerPoint edit comments in plain English and have them applied. Cover the realistic comment categories:

- **Granular text** — paragraph/bullet level edits, find-replace across deck
- **Layout/composition** — whitespace cleanup, alignment, distribute, replace one shape kind with another, add lines/dividers, rebuild slide
- **Cross-cutting batch** — recolor by match, find-replace by scope
- **Speaker notes** — read + write
- **Images** — insert + replace by local file path
- **Slide structure** — reorder, import from another deck, extract slides
- **Tables** — add/delete row/col, merge cells, format rows
- **Native charts** — type, title, axis title, legend position (no data edits this phase)
- **Groups** — group/ungroup, see internals
- **Connectors** — arrows between shapes

### Non-goals (V2)

- **Chart data / series edits.** Requires opening hidden Excel workbook (`Chart.ChartData.Workbook`). Defer to V3 if/when needed.
- **URL image fetch.** Only local file paths in V2.
- **Pasted-as-image charts.** Only `Shape.HasChart = True` shapes are eligible for chart actions; pasted images surface as `picture` type.
- **API integration.** Workflow stays copy-paste between user and their employer's LLM aggregator. No HTTP calls from VBA.
- **Vision-grounded LLM.** No PNG export per slide. Snapshot is text-only JSON. (User accepts JSON length growth.)
- **Per-run paragraph editing actions.** Snapshot will expose runs for inspection, but write actions operate at paragraph granularity, not run-by-run. Per-run write is V3.
- **Animations, transitions, comments, master/layout edits.**

## 2. Architecture

### 2.1 Module split

`modActions.bas` keeps growing; we split by bucket so each file stays small and a subagent can hold the whole file in context:

```
src/
├── modJSON.bas                       (unchanged from V1)
├── modBackup.bas                     (unchanged)
├── modExportSnapshot.bas             (extended — paragraphs, notes, gaps, chart, group)
├── modActions.bas                    (V1's 15 actions, unchanged)
├── modActionsText.bas      NEW
├── modActionsLayout.bas    NEW
├── modActionsTable.bas     NEW
├── modActionsChart.bas     NEW
├── modActionsImage.bas     NEW
├── modActionsSlide.bas     NEW
├── modActionsGroup.bas     NEW
├── modActionsConnector.bas NEW
├── modExecuteInstructions.bas        (extended — validate + dispatch ~70 actions)
├── modUI.bas                         (extended — ImportSlides entry point)
├── frmExport.frm/.frx                (PromptTemplate fn extended)
├── frmExecute.frm/.frx               (unchanged behavior)
└── frmImportSlides.frm/.frx  NEW     (file picker + slide range + target position)
```

Cross-module call rules: every `modActionsX.bas` may call `modActions.FindShape` and `modActions.HexToRgb` (V1 helpers). The dispatcher in `modExecuteInstructions.bas` knows about every Sub by name and routes via `Select Case act("type")`.

### 2.2 Build order

```
[2.0 Snapshot v2]  ← must land first
        │
        ├──► [2.A Granular text]      (depends on paragraphs[] in snapshot)
        ├──► [2.B Layout]             (uses occupied_rects from 2.0)
        ├──► [2.C Cross-cutting]      (depends on paragraphs[])
        ├──► [2.D Speaker notes]      (depends on speaker_notes in snapshot)
        ├──► [2.E Images]             (independent)
        ├──► [2.F Slide structure]    (independent)
        ├──► [2.G Tables]             (depends on merged-cell info)
        ├──► [2.H Groups]             (depends on group_children)
        ├──► [2.I Connectors]         (independent)
        ├──► [2.K Charts]             (depends on chart{} in snapshot)
        ├──► [2.UI frmImportSlides]   (independent)
        └──► [2.Z Prompt template + final smoke]
```

Within Phase 2.A through 2.K, work can run roughly in parallel by separate subagents — they touch different `.bas` files. Snapshot v2 is the bottleneck and must complete first.

## 3. Snapshot v2

### 3.1 New top-level keys per slide

```json
{
  "slide_number": 1,
  "layout_name": "Title Slide",
  "speaker_notes": "Talking points: highlight Q3 outperformance...",
  "occupied_rects": [
    {"shape_id": 3, "left": 50, "top": 40, "right": 910, "bottom": 120},
    {"shape_id": 5, "left": 50, "top": 150, "right": 910, "bottom": 450}
  ],
  "shapes": [ ... ]
}
```

`occupied_rects` is a flat list of bounding rectangles for every top-level shape (computed from existing `pos`). It lets the LLM reason about whitespace without us shipping a precomputed gap-map. Cheap to produce (one entry per shape).

### 3.2 New shape-level keys

```json
{
  "shape_id": 5,
  "type": "body",
  "pos": {...},
  "text": "Revenue up 12%\nMargins improved",
  "font": {"name": "Calibri", "size": 18, "bold": false, "italic": false, "color": "#000000"},
  "fill": null,

  "paragraphs": [
    {
      "index": 0,
      "text": "Revenue up 12%",
      "bullet_style": "disc",
      "indent_level": 0,
      "runs": [
        {"text": "Revenue up ", "font": {"name":"Calibri","size":18,"bold":false,"italic":false,"color":"#000000"}},
        {"text": "12%",         "font": {"name":"Calibri","size":18,"bold":true, "italic":false,"color":"#1F4E79"}}
      ]
    },
    {
      "index": 1,
      "text": "Margins improved",
      "bullet_style": "disc",
      "indent_level": 0,
      "runs": [{"text":"Margins improved","font":{...}}]
    }
  ],

  "group_children": [<shape_dict>, <shape_dict>],

  "chart": {
    "type": "barChart",
    "is_native": true,
    "title": "Quarterly Revenue",
    "axis_titles": {"x": "Quarter", "y": "Revenue ($M)"},
    "legend_position": "right",
    "series": [
      {"name": "FY24", "categories": ["Q1","Q2","Q3","Q4"], "values": [100,110,120,130]}
    ]
  },

  "table_extra": {
    "merged_cells": [{"row": 1, "col": 1, "row_span": 1, "col_span": 3}]
  }
}
```

#### Field semantics

- **`paragraphs`**: only for text-bearing shapes (placeholder, textbox, table cell). Each paragraph has a 0-based `index`, raw `text`, `bullet_style` (`none | disc | number | letter | square | dash | image`), `indent_level` (0-4 inclusive), and `runs`. A run is a contiguous span of identically-formatted text.
- **`group_children`**: only when `Shape.Type = msoGroup`. Each child is a full shape_dict (recursive); use `shape_id` to address — child IDs are unique inside the group's slide.
- **`chart`**: only when `Shape.HasChart = True`. `is_native: true` always (we don't emit `chart` for `Shape.Type = msoPicture` even if it visually looks like a chart). `series.values` is read via `Series.Values` — if reading triggers the embedded workbook (rare on some PowerPoint builds), set `values: null` and continue.
- **`table_extra`**: only when `Shape.HasTable = True`. Contains merged-cell info keyed by anchor cell `(row, col)` plus its `row_span` and `col_span`. Cells that are merged-into-anchor still appear in `table.cells[][]` but their `text` will be empty.

### 3.3 Things deliberately not in Snapshot v2

- Per-shape z-order / send-to-back / bring-to-front. Out of scope for our edits.
- Rotation. Out of scope.
- Gradient/pattern fills. Solid fills only (V1 already has this).
- Hyperlinks. Out of scope.
- Animations. Out of scope.
- Comments / review markup. Out of scope.

### 3.4 Snapshot size

V1 snapshot of a 10-slide deck = ~5-15 KB. V2 will be 30-80 KB depending on content. User accepts the LLM aggregator paste-box size growth — no toggle.

## 4. Action types — full V2 catalogue

Total: 15 (V1) + 55 (V2) = 70 actions across 11 buckets.

### 4.A Granular text (8 new)

| Action | Required keys | Effect |
|---|---|---|
| `set_paragraph_text` | slide, shape_id, paragraph_index, value | Replace one paragraph's text. |
| `add_paragraph` | slide, shape_id, after_paragraph_index, value | Insert a new paragraph after the given index (`-1` = beginning). |
| `delete_paragraph` | slide, shape_id, paragraph_index | Remove a paragraph. |
| `set_bullet_style` | slide, shape_id, paragraph_index, value (`none\|disc\|number\|letter\|square\|dash`) | Change bullet style. |
| `set_indent_level` | slide, shape_id, paragraph_index, value (0-4) | Change indent level. |
| `set_paragraph_font_size` | slide, shape_id, paragraph_index, value (Long) | Per-paragraph font size override. |
| `set_paragraph_font_color` | slide, shape_id, paragraph_index, value (#RRGGBB) | Per-paragraph font color override. |
| `find_replace_text` | scope (`deck` or `slide:N`), find, replace | Substring replace across all text shapes in scope. Case-sensitive. |

### 4.B Layout / composition (10 new)

| Action | Required keys | Effect |
|---|---|---|
| `align_shapes` | slide, shape_ids[], anchor (`left\|right\|top\|bottom\|hcenter\|vcenter`) | Align listed shapes' edges to the anchor of the first shape in the list. |
| `distribute_horizontal` | slide, shape_ids[] | Equal horizontal spacing between left edges, preserving leftmost and rightmost shapes' positions. |
| `distribute_vertical` | slide, shape_ids[] | Equal vertical spacing. |
| `tile_grid` | slide, shape_ids[], cols (Long), gap_pt (Single) | Lay out shapes in a row-major grid with given column count and inter-cell gap, anchored at first shape's top-left. |
| `fit_to_slide_margins` | slide, shape_id, margin_pt (Single, default 36) | Resize + center a single shape so it fits inside slide minus margin on all sides. |
| `add_line` | slide, x1, y1, x2, y2 (pt), color (#RRGGBB), weight_pt (Single) | Draw a straight line. |
| `add_shape` | slide, kind (`rect\|rrect\|oval\|circle\|capsule\|arrow\|diamond\|triangle`), pos {left,top,width,height}, fill (hex or null), stroke (hex or null), stroke_weight_pt (Single) | Add a new shape with the given geometry and styling. |
| `set_shape_kind` | slide, shape_id, kind | Replace an existing shape's auto-shape kind in place (preserves pos + fill + text). Implemented as Shape.AutoShapeType assignment. |
| `clear_slide` | slide, keep_shape_ids[] (default empty) | Delete every top-level shape on the slide except those in the keep list. Used by LLM as a "rebuild this slide" primitive. |
| `move_shape_relative` | slide, shape_id, dx_pt (Single), dy_pt (Single) | Convenience: move by delta. Equivalent to `move_shape` with computed absolute target. Useful for "move this down 50pt" comments. |

### 4.C Cross-cutting batch (3 new)

| Action | Required keys | Effect |
|---|---|---|
| `recolor_fill_match` | scope (`deck` or `slide:N`), from (#RRGGBB), to (#RRGGBB) | Change every solid fill matching `from` to `to`. |
| `recolor_font_match` | scope, from, to | Same for font color across text frames. |
| `delete_shapes_match` | scope, kind (optional), fill (optional, hex), text_contains (optional, string) | Delete shapes matching all provided criteria. At least one criterion required. |

LLM may also enumerate matches in the snapshot itself and emit N atomic actions; both styles work, batch helpers are convenience.

### 4.D Speaker notes (2 new)

| Action | Required keys | Effect |
|---|---|---|
| `set_speaker_notes` | slide, value | Replace notes text. |
| `append_speaker_notes` | slide, value | Append to existing notes (with newline separator). |

### 4.E Images (2 new)

| Action | Required keys | Effect |
|---|---|---|
| `insert_picture` | slide, path (local file, .png/.jpg/.jpeg/.gif/.bmp), pos {left,top,width,height} | Insert a new picture at the given pos. Errors if file does not exist. |
| `replace_picture` | slide, shape_id, path | Replace an existing picture's image source. Preserves pos. |

### 4.F Slide structure (3 new)

| Action | Required keys | Effect |
|---|---|---|
| `move_slide` | from (Long), to (Long) | Reorder a slide. |
| `extract_slides` | slide_indices[], output_path (file path) | Save the listed slides as a new .pptx at output_path. Preserves theme + master. |
| `import_slides_from_deck` | source_path, slide_indices[], target_position | Insert slides from another deck at target_position in current deck. (Also exposed via standalone `frmImportSlides` UserForm.) |

### 4.G Tables (5 new)

| Action | Required keys | Effect |
|---|---|---|
| `add_table_row` | slide, shape_id, after_row (Long, `0` = top) | Insert blank row after specified row index. |
| `delete_table_row` | slide, shape_id, row | Delete a row (1-based). |
| `add_table_col` | slide, shape_id, after_col | Insert blank column. |
| `delete_table_col` | slide, shape_id, col | Delete a column. |
| `merge_cells` | slide, shape_id, row_a, col_a, row_b, col_b | Merge a rectangular cell range. The top-left cell becomes the anchor. |

### 4.H Groups (2 new)

| Action | Required keys | Effect |
|---|---|---|
| `group_shapes` | slide, shape_ids[] | Group the listed shapes; the new group gets a fresh shape_id reflected on next snapshot. |
| `ungroup` | slide, shape_id | Ungroup; children become top-level shapes with their own IDs. |

### 4.I Connectors (1 new)

| Action | Required keys | Effect |
|---|---|---|
| `add_connector` | slide, from_shape_id, to_shape_id, kind (`straight\|elbow\|curved`), arrow_end (`none\|open\|filled`, default `filled`), color, weight_pt | Connect two shapes with an arrow. |

### 4.K Native charts (5 new — non-workbook ops only)

Skips when `Shape.HasChart = False` with `not_a_native_chart` reason.

| Action | Required keys | Effect |
|---|---|---|
| `set_chart_type` | slide, shape_id, value (PowerPoint XlChartType name, e.g. `xlBarClustered`, `xlLine`, `xlPie`, `xlColumnClustered`) | Change chart type. |
| `set_chart_title` | slide, shape_id, value (string), enabled (bool, optional, default true) | Set title text. If `enabled=false`, hide title. |
| `set_chart_axis_title` | slide, shape_id, axis (`x\|y`), value | Set axis title. |
| `set_chart_legend_position` | slide, shape_id, value (`right\|left\|top\|bottom\|corner\|none`) | Move or hide legend. |
| `set_series_color` | slide, shape_id, series_index (Long, 1-based), value (#RRGGBB) | Recolor a single series. Doesn't open workbook on most PowerPoint builds. |

If `set_series_color` triggers workbook-open in testing, demote to V3 and document. Other 4 are pure Chart object property assignments.

## 5. New UserForm: `frmImportSlides`

Standalone — no LLM round-trip.

### Controls

| Control | Name | Notes |
|---|---|---|
| Label | `lblPath` | Caption: "Source deck" |
| TextBox | `txtPath` | Read-only; populated by file picker |
| CommandButton | `btnBrowse` | Caption: "Browse..." → opens `Application.FileDialog(msoFileDialogFilePicker)` filtered to `.pptx`/`.pptm`. Fallback: InputBox if FileDialog unavailable. |
| Label | `lblRange` | Caption: "Slide range (e.g. 1-3,5,7-9)" |
| TextBox | `txtRange` | Free-form; parser converts to `[1,2,3,5,7,8,9]` |
| Label | `lblPosition` | Caption: "Insert at position" |
| TextBox | `txtPosition` | Numeric input |
| CommandButton | `btnImport` | Disabled until path + range filled; on click calls `modActionsSlide.Do_import_slides_from_deck(path, indices, position)` and shows summary |
| CommandButton | `btnCancel` | |
| Label | `lblStatus` | Final summary or error |

Range parser accepts `1-3,5,7-9`, comma-separated, hyphenated ranges. Validates numbers > 0 and ≤ source-deck slide count (after opening the source).

### Macro entry point

`Public Sub ImportSlides()` in `modUI.bas` shows `frmImportSlides`. Visible in Alt+F8.

## 6. Prompt template growth (frmExport)

V1 template lists 15 schemas. V2 will list 70. To stay under VBA's 24-line continuation limit (already learned the hard way), the template builder is a `Private Function PromptTemplate() As String` using `s = s & ...` per line — no `& _` chains.

Section structure of the new template:

```
[intro + snapshot placeholder + user request placeholder]

15 ATOMIC SCHEMAS (V1)
  set_text, set_font_*, set_fill_color, move_shape, resize_shape,
  delete_shape, add_slide, delete_slide, duplicate_slide,
  set_cell_text, swap_table_columns, swap_table_rows

GRANULAR TEXT
  set_paragraph_text, add_paragraph, delete_paragraph,
  set_bullet_style, set_indent_level, set_paragraph_font_size,
  set_paragraph_font_color, find_replace_text

LAYOUT / COMPOSITION
  align_shapes, distribute_horizontal, distribute_vertical, tile_grid,
  fit_to_slide_margins, add_line, add_shape, set_shape_kind,
  clear_slide, move_shape_relative

CROSS-CUTTING BATCH
  recolor_fill_match, recolor_font_match, delete_shapes_match
  (Or LLM may enumerate matches and emit N atomic actions.)

SPEAKER NOTES
  set_speaker_notes, append_speaker_notes

IMAGES
  insert_picture, replace_picture (LOCAL FILE PATHS ONLY — no URLs)

SLIDE STRUCTURE
  move_slide, extract_slides, import_slides_from_deck

TABLES
  add_table_row, delete_table_row, add_table_col, delete_table_col,
  merge_cells

GROUPS
  group_shapes, ungroup

CONNECTORS
  add_connector

NATIVE CHARTS
  set_chart_type, set_chart_title, set_chart_axis_title,
  set_chart_legend_position, set_series_color
  (Only for shapes with "chart": {...} in snapshot. Pasted images skipped.)

RULES
  - Use only shape_ids that exist in the snapshot.
  - Slide / row / col / paragraph / series numbers are 1-based EXCEPT
    paragraph_index which is 0-based to match snapshot.
  - Colors as #RRGGBB hex.
  - Lengths in points.
  - Booleans are JSON true / false.
  - One field name per action — never substitute aliases.
  - For "every X with property P → Y" requests: enumerate matching
    shape_ids in the snapshot, emit one action per match (or use the
    *_match helper if available).
  - For "rebuild this slide" requests: optionally use clear_slide first,
    then emit a sequence of add_shape / set_text / move_shape actions.
```

Each schema section shows one canonical example per action with all field names. ~80 example lines total.

## 7. Validation + dispatch updates

`modExecuteInstructions.bas` extends `ValidateAction` with cases for the 55 new types. Each case:
1. `RequireFields(act, Array(...))` — same pattern as V1.
2. `ValidateShape(act)` or `ValidateSlide(act)` where applicable.
3. Type-specific guard (e.g. `set_chart_type` requires the shape to have `HasChart = True`).

`DispatchAction` extends with `Case` arms calling the right `modActionsX.Do_<type>` Sub.

The action log format is unchanged. Backup behavior is unchanged. Per-action `On Error Resume Next` containment is unchanged.

## 8. Smoke test plan

V1 has 10 tests covering 15 actions. V2 needs ~30 more tests. Each new action gets either (a) its own assertion block, or (b) is rolled into a multi-action test grouped by bucket. Final shape:

- 1 test per bucket exercising 3-6 actions in that bucket.
- 1 negative test per validation rule (missing field, wrong type, etc.) — sample only, not exhaustive.
- 1 final end-to-end test combining 20+ actions across buckets in a single instructions JSON.

Estimated total run time: ~3-4 minutes. Acceptable.

## 9. Risks (revised)

1. **`set_series_color` may open workbook silently.** If during implementation testing this triggers Excel, demote to V3 and remove from V2 surface.
2. **`set_shape_kind` constraint.** PowerPoint allows `AutoShapeType` reassignment only if both the source and target are auto-shapes. If source is a placeholder or freeform, action errors with `not_an_autoshape`. Document and let LLM handle.
3. **`group_shapes` shape_id stability.** After grouping, child shape IDs change in some PowerPoint builds. Snapshot must be regenerated after a group action; don't assume stale IDs.
4. **Per-paragraph font reads.** A paragraph with mixed-format runs returns `Size = -1`, `Color = mixed`. `BuildParagraphDict` must handle this — emit `null` for any mixed property and emit per-run details, so LLM can drill down.
5. **Smoke test deck regeneration.** `tests/make_test_decks.py` will need a new fixture deck `phase2.pptx` with a native chart, grouped shapes, merged table cells, multi-paragraph bullet body, and speaker notes. Generated by python-pptx.
6. **Carrier compile errors.** Each new module is a fresh compile-error risk. Sanity check after every sync (call `BuildSnapshotJson`); if it errors with "Sub or function not defined" the new module didn't compile.

## 10. Out of scope (V3+)

- Chart data CRUD, add/remove series, axis-tick formatting (require Excel COM workbook hook)
- URL image fetch
- Per-run text writes (read is fine in V2; write is V3)
- Theme / master / layout edits
- Animations, transitions
- Comments / review markup
- Hyperlinks
- Rotation, gradients, pattern fills
- Multi-monitor / multi-deck simultaneous edits

## 11. Success criteria

- All 70 actions individually exercise-able from a hand-written instructions JSON.
- A reviewer can paste a deck snapshot + a multi-bullet comment list ("slide 2 too whitespace-heavy, slide 4 add divider, slide 7 swap blue→red, slide 9 add speaker notes") into the LLM aggregator, get back a single instructions JSON, and have all edits land on first try in 80%+ of cases.
- `frmImportSlides` UserForm picks a file, parses range, imports — no LLM needed.
- Snapshot v2 always serializes cleanly even on edge-case decks (mixed-format paragraphs, charts, groups, merged cells).

# Decko.ai

VBA-based PowerPoint editor driven by natural language via an external LLM.

The user copies a JSON snapshot of the active deck into their LLM tool,
describes a desired change, copies the LLM's instructions JSON back, and
clicks Apply. A handful of Alt+F8 macros, three UserForms, no API calls from PowerPoint.

After every Apply, an automatic 32-check quality-verification loop sweeps the
deck and flags problems (off-slide shapes, unreadable contrast, tiny fonts,
broken hyperlinks, etc.). Two one-click buttons — **Fix Errors** (pre-Apply)
and **Fix This** (post-Apply) — copy LLM-ready repair prompts to the clipboard
so the user never reads raw error text or opens a JSON file by hand.

**Fix Errors covers every action type** (246 canonical types) with
canonical signature + working example, so even a weak LLM that produced
malformed JSON gets enough context to self-correct without the user typing
anything. When the LLM invents an action name that doesn't exist, Fix Errors
also emits **"DID YOU MEAN: ..."** suggestions (word-stem similarity over the
master action list). The same canonical guidance is mirrored into the public
`docs/ACTIONS_REFERENCE.md` auto-appendix via
`python tools/sync_actions_guidance.py` and policed by a drift test, so the
in-app and on-disk references can never disagree. See
[`docs/VERIFICATION.md`](docs/VERIFICATION.md) for the full check list and the
button mechanics.

## Documentation

- **[`SETUP.md`](SETUP.md)** — fresh-clone bulletproof setup. **Read first if
  you just cloned this repo.**
- **[`docs/PROMPTING_GUIDE.md`](docs/PROMPTING_GUIDE.md)** — how to phrase
  requests, worked examples, and a section for AI assistants on turning a VP
  request into the `actions` JSON.
- **[`docs/ACTIONS_REFERENCE.md`](docs/ACTIONS_REFERENCE.md)** — the complete,
  machine-precise schema for all 246 actions (required/optional fields, value
  vocabularies, examples). Read this literally; you don't need to have built
  Decko to use it. The auto-appendix covers every action including the
  high-level authoring layer (templates, decks-as-code, Deck DNA).
- **[`docs/VERIFICATION.md`](docs/VERIFICATION.md)** — quality-check loop and
  the two Fix buttons explained end-to-end (what runs, when, what each warning
  looks like, opt-out flags, performance limits).
- **[`docs/EXAMPLES.md`](docs/EXAMPLES.md)** — a corpus of paired examples
  (VP prompt → exact `actions` JSON), ~45 worked cases across every action
  category. Best place for an agent to learn the request→actions mapping.
- **[`docs/LAYOUT_RECIPES.md`](docs/LAYOUT_RECIPES.md)** — redesigning a slide's
  whole layout: a catalog of region presets (67/33, 50/50, quad, 3-/4-column,
  2-stacked-left + 1-right, etc., with exact pt coordinates for 960×540).
- **[`docs/USER_GUIDE.md`](docs/USER_GUIDE.md)** — practical guide for VPs/MDs.
- Design specs in `docs/specs/`.

## Desktop app (SP1)

`Decko-Setup.exe` (built via `python packaging/build.py` then
`packaging/installer.iss`) is a Windows installer requiring only
Microsoft PowerPoint — no Python. It wraps the same VBA/COM engine in a
chat + side-panel UI; users supply their own LLM key (Anthropic / OpenAI
/ generic OpenAI-compatible), stored in Windows Credential Manager.
Design: `docs/superpowers/specs/2026-05-17-decko-desktop-app-design.md`.
Plan: `docs/superpowers/plans/2026-05-17-decko-desktop-sp1.md`.
SP1 gate: `python tests/run_smoke_app.py` (UI visuals are manually
verified, not in the deterministic gate, per the design's honest scope).

### Templates panel (SP2)

The desktop app has a slide-over **Templates ▸** panel: apply the 7
built-in layouts or your captured "Deck DNA" templates with instant
placeholders (optional Fill-with-AI), capture the active slide,
rename/delete captured templates, generate layout variants, and
extract / build a deck spec (`{"deck":[…]}`) — all without an LLM
round-trip for the deterministic operations; each is logged as a
session turn. Gate: the `templates` gate inside
`python tests/run_smoke_app.py`; the panel itself is verified manually.
Design: `docs/superpowers/specs/2026-05-18-sp2-templates-ui-design.md`.
Plan: `docs/superpowers/plans/2026-05-18-sp2-templates-ui.md`.

## One-time developer setup

See [`SETUP.md`](SETUP.md) for the bulletproof checklist. Short version:

```bash
pip install -r requirements.txt        # pywin32 + python-pptx
python update_macros.py                # sync src/*.bas + *.frm into carrier
python tools/add_fix_button.py         # one-time install of Fix Errors / Fix This buttons
```

Plus: enable **Trust access to the VBA project object model** in PowerPoint
Trust Center.

## Daily use

1. Open the deck you want to edit, plus `PPT_AI_Editor.pptm`.
2. Press Alt+F8 → run `ExportSnapshot`. The form opens with a JSON snapshot.
3. Click **Copy snapshot + prompt template**. Your clipboard now holds a
   prompt block ending with `[REPLACE THIS LINE WITH YOUR REQUEST]`.
4. Paste into your LLM tool. Replace the placeholder line with what you want
   changed in the deck. Submit.
5. The LLM returns an instructions JSON. Copy it.
6. Press Alt+F8 → run `ExecuteInstructions`. Paste the JSON and click
   **Parse**, or — for a large batch — click **"Load from file..."** and pick
   the `.json` file (the text box corrupts big pastes; the file path doesn't).
   Review the action list — invalid rows are tagged, and a
   **`WILL DO (preview)`** block lists, in plain language, exactly what
   each action will do (read-only; nothing has changed yet). This is your
   chance to catch intent drift *before* Apply, which has no undo.
7. **If any action is INVALID** — click **Fix Errors**. The clipboard now
   holds a prompt with each failing action + canonical signature + example.
   Paste into your LLM chat, get a corrected batch back, go to step 6.
   *(If all actions are valid but the preview is not what you meant, click
   **Fix Errors** anyway — it copies a "re-steer" prompt: the plan plus a
   "this is not what I meant, revise the actions" template.)*
8. **All OK** — click **Apply**. There is **no auto-backup and no undo** —
   Apply mutates the currently open deck in place, so save (Ctrl+S) before
   closing or lose the changes. Valid actions execute in order; the
   returned summary shows applied/skipped counts, and on any failure a
   `FAILURES (N):` block naming each failed action's exact batch index,
   type, and reason (no silent swallow). The verify loop then runs
   automatically (typically <500 ms): `verification: N warning(s), M info`.
9. **If verify found warnings** — click **Fix This**. Clipboard holds a prompt
   with each quality issue + a suggested fix action. Paste into LLM, get the
   fix batch, go to step 6.

## Action types

246 dispatched actions across 17 `modActions*` modules. The tables below are a
summary; the **complete per-action schema** is in
[`docs/ACTIONS_REFERENCE.md`](docs/ACTIONS_REFERENCE.md) (auto-generated
appendix kept in sync via `tools/sync_actions_guidance.py`).

### High-level authoring: templates, decks-as-code, Deck DNA

These actions build whole slides/decks in one shot instead of placing
individual shapes — use them first when the request is "make a slide that
does X" rather than "move this box".

**Built-in templates (`modActionsTemplate.bas`, 7 layouts)**

| Action | Effect |
|---|---|
| `apply_template` | Build a slide from a named layout in ONE action. `template` ∈ `title` (`title,subtitle`), `section` (`section_number,section_title`), `bullets` (`heading,bullets[]`), `two_col` (`heading,left_body,right_body`), `comparison` (`heading,left_label,left_body,right_label,right_body`), `kpi_dashboard` (`heading,tiles[]` of `{stat,label}`), `quote` (`quote_text,attribution`) — or any captured-template name. `content` holds those literal slots. Optional `slide:N` to replace, else appends. |

**Decks-as-code (`modActionsSpec.bas`)**

| Action | Effect |
|---|---|
| `build_deck_from_spec` | Build a multi-slide deck from a compact `spec` (array of `{template, content}` slide objects). One action → full deck. |
| `extract_spec` | Reverse: read the deck back out as a `spec` JSON (round-trips with `build_deck_from_spec`). Use to clone/restyle an existing deck. |
| `generate_variants` | Re-render given content into N distinct **principled layout archetypes** (Hero / Split / Stack / Quote / Tiles, cycling) — NOT a cosmetic position shuffle. Form A: `template`+`n` (heading pulled from `title`/`heading`/`section_title`/`quote_text`, rest → body). Form B: `templates:[names]` renders the same content across each named template. |

**Deck DNA — user-captured templates (`modActionsCapture.bas`)**

A captured template is a real slide the user liked, saved as a reusable
stamp. The registry is external JSON data at
`%APPDATA%\Decko\templates.json` (NOT code) and a live manifest of captured
names is appended to the snapshot prompt so the LLM can target them by name.

| Action | Effect |
|---|---|
| `capture_template` | Save a slide's layout+style as a named reusable template. `name` + `slide:N` (default active). Auto-derives content slots. |
| `list_templates` | Return the captured-template registry (names + slot summary). |
| `delete_template` | Remove a captured template by `name`. Slides already built from it are unchanged. |
| `rename_template` | Rename a captured template — params `from` → `to` (not `name`/`new_name`). |

Once captured, a template name is valid in `apply_template` /
`build_deck_from_spec` exactly like a built-in.

**DECK DESIGN PRINCIPLES** — the snapshot prompt now ends with an injected
design-principles block (hierarchy, contrast, alignment, restraint) so the
LLM produces well-composed slides without the user re-explaining taste.

**Icons** — the prompt no longer ships an exhaustive Fluent allow-list. It
gives concise CDN-sourcing guidance + a short curated name list; the LLM
sources icon SVGs by semantic name from the CDN. Works on locked-down work
machines where the old allow-list path failed.

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

### Granular text (`modActionsText.bas`, 14)

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
| `set_paragraph_alignment` | Left / center / right / justify per paragraph. |
| `set_paragraph_line_spacing` | Line spacing as multiple (1.0, 1.5, 2.0...). |
| `set_text_vertical_align` | Whole-shape vertical anchor (top / middle / bottom). |
| `set_text_margin` | Whole-shape internal margins (left / right / top / bottom in pt). |
| `set_text_autofit` | Per-shape autofit mode (none / shrink / resize). |
| `enable_text_shrink_for_overflow` | Sweep deck/slide; turn on shrink-on-overflow for every text frame (skips titles). |

### Run-level formatting (`modActionsRun.bas`, 10)

| Action | Effect |
|---|---|
| `set_run_bold` | Toggle bold on a single run. |
| `set_run_italic` | Toggle italic. |
| `set_run_underline` | Toggle underline. |
| `set_run_subscript` | Subscript on/off (clears superscript). |
| `set_run_superscript` | Superscript on/off (clears subscript). |
| `set_run_font_color` | `#RRGGBB` on a single run. |
| `set_run_font_size` | Run size in pt. |
| `set_run_font_name` | Run font name. |
| `set_run_text` | Replace this run's chars; siblings untouched. |
| `set_run_hyperlink` | Set or clear (empty string) hyperlink on a run. |

### Layout + alignment (`modActionsLayout.bas`, 24)

| Action | Effect |
|---|---|
| `align_shapes` | Align multiple shapes (left/right/top/bottom/hcenter/vcenter). |
| `distribute_horizontal` | Even horizontal spacing across selected shapes. |
| `distribute_vertical` | Even vertical spacing across selected shapes. |
| `tile_grid` | Arrange shapes into N-column grid with gap. |
| `fit_to_slide_margins` | Fit a shape inside slide minus margin. |
| `add_line` | Add a line connector between (x1,y1) and (x2,y2). |
| `add_shape` | Add an `msoAutoShape` (rect, oval, etc.) at given pos. |
| `snap_to_grid` | Round shape `left`/`top` to nearest multiple of `grid_pt`. |
| `align_to_slide_center` | Center shape on slide horizontally / vertically / both. |
| `nudge` | Shift shape by `amount_pt` in direction `l`/`r`/`u`/`d`. |
| `fit_to_content` | Auto-resize shape to its text bounding box. |
| `match_size` | Copy reference shape's `width`/`height` to target shapes. |
| `uniform_size` | Set all listed shapes to identical width and height. |
| `smart_spacing` | After sorting, place each shape `gap_pt` from previous's far edge. |
| `equalize_spacing` | Distribute shapes with equal gaps along axis. |
| `match_position` | Align target's edge to reference's same edge. |
| `swap_positions` | Swap two shapes' positions and sizes. |
| `group_by_overlap` | Group only shapes whose bounding boxes intersect. |
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

### Tables — build & cell formatting (`modActionsTable.bas`, more)

| Action | Effect |
|---|---|
| `add_table` | Create a new table (`rows`×`cols`) at `pos`. |
| `set_cell_text` | Set text of cell `(row, col)`. |
| `set_table_col_width` / `set_table_row_height` | Resize a column / row (pt). |
| `set_cell_border` | Border on one cell side (color/weight/visible). |
| `set_cell_text_align` | h/v align inside a cell. |
| `set_cell_fill` | Cell background color. |
| `apply_table_style` | Apply a named/GUID Office table style. |
| `build_image_grid_table` | 2-col image+caption table from a row spec. |

### Charts — native chart objects (`modActionsChart.bas`, more)

Decko creates **real native PowerPoint charts** (editable, with an embedded data
sheet), not images. **All 39 PowerPoint chart types are supported** (2-D/3-D
column & bar, line, area, pie/doughnut, scatter, radar, surface, and the modern
types: waterfall, pareto, funnel, histogram, box-and-whisker, treemap, sunburst).
The 7 modern types are created with the correct type but PowerPoint's placeholder
data (a host automation limitation) — the user edits their data manually.

| Action | Effect |
|---|---|
| `add_chart` | Insert a new chart: `chart_type`, `pos`, `categories`, `series` (`[{name,values,color?}]`), optional `title`/`show_legend`/`show_values`/`clean_style`/`value_format`/`ref_name`. |
| `set_chart_type` | Change an existing chart's type. |
| `set_chart_title` / `set_chart_axis_title` | Title / axis title text. |
| `set_chart_legend_position` / `set_chart_legend` | Legend position / props. |
| `set_chart_categories` / `set_series_values` / `set_series_name` | Replace category labels / a series' values / a series' name. |
| `set_series_color` | Color a series by 1-based index. |
| `set_chart_axis` / `set_chart_format` / `set_chart_series` | Fine axis / chart-group / per-series props (min/max/units, gap width, bar shape, markers, plot-area pinning, per-point fills/markers/line-visibility for waterfalls and clipped line series, etc.). |
| `set_chart_gridlines` | Show / hide / style major & minor gridlines per axis. |
| `add_chart_trendline` / `set_chart_error_bars` | Add a trendline / error bars to a series. |

### Images & web (`modActionsImage.bas`, `modActionsWeb.bas`)

| Action | Effect |
|---|---|
| `insert_picture` | Insert a local image at position. |
| `replace_picture` | Replace existing picture, preserving frame. |
| `insert_icon` | Insert a Microsoft Fluent UI SVG icon (concept name → CDN fetch + recolor). |
| `fetch_page_images` | Scrape all images from a URL into a folder. |
| `download_image` | Download one image URL to a local path. |
| `open_image_picker` / `build_image_picker_slide` | Visual thumbnail-grid picker / build a grid slide from a folder. |
| `bulk_insert_image` | Same image, same box, across multiple slides. |

### Connectors + groups (`modActionsConnector.bas`, `modActionsGroup.bas`, 3)

| Action | Effect |
|---|---|
| `add_connector` | Add elbow/straight connector between two shapes. |
| `group_shapes` | Group shapes into one. |
| `ungroup` | Ungroup a group shape. |

### Visual polish + effects (`modActionsEffects.bas`, 16)

| Action | Effect |
|---|---|
| `rotate_shape` | Set rotation in degrees. |
| `flip_shape` | Flip horizontal (`h`) or vertical (`v`). |
| `set_line_color` | Outline color `#RRGGBB`. |
| `set_line_weight` | Outline weight in pt. |
| `set_line_style` | Outline style: `solid`/`dash`/`dot`/`dashdot`. |
| `set_shadow` | Shadow with offsetX/Y, blur, color, transparency. |
| `set_glow` | Outer glow with color, radius, transparency. |
| `set_reflection` | Reflection size, transparency, distance. |
| `set_transparency` | Fill transparency `0.0..1.0`. |
| `set_gradient_fill` | Two-color horizontal gradient with angle. |
| `set_3d_bevel` | 3D bevel: `circle`/`slope`/`cross`/`angle`/`softround` + depth. |
| `apply_preset_effect` | Office preset texture index `1..24`. |
| `crop_picture` | Crop edges (left/right/top/bottom in pt). |
| `recolor_picture` | `grayscale`/`sepia`/`washout`/`bw`/`auto`. |
| `set_brightness` | Picture brightness `-1.0..1.0`. |
| `set_contrast` | Picture contrast `-1.0..1.0`. |

### Deck-wide ops (`modActionsDeck.bas`, 9)

| Action | Effect |
|---|---|
| `find_replace_regex` | Regex find/replace across `deck` or `slide:N` scope. |
| `swap_font_deck_wide` | Replace one font name with another across all text. |
| `recolor_palette_deck_wide` | Replace one color with another (target = fill / font / both). |
| `apply_theme` | Apply a `.thmx` or `.potx` theme via `ApplyTemplate`. |
| `set_slide_size` | Set slide dims (`width_pt`+`height_pt`) or preset (`16:9` / `4:3`). |
| `set_theme_font` | Set major (heading) and/or minor (body) theme fonts. |
| `bulk_insert_image` | Insert same image at same position across listed slides. |
| `bulk_insert_text_box` | Insert same text box across listed slides. |
| `apply_layout_to_slides` | Force layout index N on listed slides. |

### Slide structure (`modActionsSlide.bas`)

| Action | Effect |
|---|---|
| `move_slide` | Reorder slide from index A to B. |
| `extract_slides` | Export selected slides to a new .pptx file. |
| `import_slides_from_deck` | Import slides from another deck at position. |
| `set_slide_background_color` | Solid background color on a slide. |
| `insert_slide_number` | Add a slide-number text placeholder (pos/font/color). |

### Misc shape/text ops

| Action | Effect |
|---|---|
| `add_text_box` | New plain text box at `pos` (text/font/align/fill/stroke optional). |
| `add_shape` | New autoshape (`kind` + `pos`; fill/stroke/text/ref_name optional). |
| `add_line` | Straight line/divider between two points (arrows/dash optional). |
| `z_order` | Bring to front / send to back / forward / backward. |
| `duplicate_shape` | Clone a shape at a new position. |
| `copy_formatting` | Copy fill/line/font/effects from one shape to another. |
| `set_shape_adjustment` | Drag a shape's yellow adjustment handle. |
| `flip_shape` | Flip horizontal / vertical. |
| `set_run_strikethrough` | Strikethrough on a single run. |
| `set_speaker_notes` / `append_speaker_notes` | Replace / append slide speaker notes. |

> The tables in this README are a curated summary. The **complete list with
> every field, default, and value vocabulary** is
> [`docs/ACTIONS_REFERENCE.md`](docs/ACTIONS_REFERENCE.md).

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

UserForms: `src/*.frm`/`.frx` is the source of truth. `tools/build_forms.py`
rebuilds the control **layout** of all three forms and **preserves** the
VBA code imported from `src/*.frm` (it no longer re-stamps code from a
constant — that previously caused silent regressions). Edit form *code* in
`src/*.frm` and run `update_macros.py`. To add a **button** to an existing
form, do NOT hand-author `.frx`; clone the idempotent installer pattern in
`tools/add_fix_button.py` / `tools/add_export_buttons.py` (adds the control
via the VBE Designer, appends the handler via the CodeModule, re-exports
`.frm`/`.frx`). Only re-run `build_forms.py` for a full layout rebuild.

## Tests

```bash
python tests/make_test_decks.py        # regenerate test decks
python update_macros.py                # ensure carrier matches src/
python tests/run_smoke.py              # end-to-end COM-driven smoke
```

Targeted, deterministic COM-driven harnesses (each exits non-zero unless its
metric is met; all are resilient to transient PowerPoint COM errors — they
retry bring-up / `com_error` only, never an assertion):

```bash
python tests/run_smoke_sanitizer.py    # SanitizeJsonInput corpus (49/49)
python tests/run_smoke_verify.py       # modVerify precision/recall = 1.0 vs frozen contract
python tests/run_smoke_preview.py      # BuildActionPlanSummary: 246 coverage + exact corpus
python tests/run_smoke_validate.py     # ValidateBatchJson: recognition/rejection/no-false-reject
python tests/run_smoke_guidance.py     # GetActionGuidance: 246 coverage + schema-valid EXAMPLEs
python tests/run_smoke_failcontract.py # ExecuteFromString partial-failure contract
python tests/run_smoke_schema_audit.py # cross-surface key-consistency lint
python tests/run_smoke_template.py     # apply_template: 7 builtin layouts
python tests/run_smoke_spec.py         # build_deck_from_spec / extract_spec round-trip
python tests/run_smoke_variants.py     # generate_variants: distinct principled archetypes
python tests/run_smoke_capture.py      # capture/list/delete/rename_template registry
python tests/run_smoke_dialogs.py      # Capture/Manage dialogs + icon-trim
python tests/run_smoke_icon_prompt.py  # prompt ships CDN guidance, not the allow-list
```

## Macros (Alt+F8)

| Macro | UserForm | Purpose |
|---|---|---|
| `ExportSnapshot` | `frmExport` | Build deck snapshot JSON; copy snapshot + prompt to clipboard. |
| `ExecuteInstructions` | `frmExecute` | Paste instructions JSON, parse, review, apply. |
| `ImportSlides` | `frmImportSlides` | Import slides from another deck at a given position. |
| `CaptureTemplate` | InputBox | Save the active slide as a named Deck DNA template (registry JSON). |
| `ManageTemplates` | InputBox | View captured templates (numbered) and delete one by name. |

`frmExport` also has two one-click action shortcuts (no JSON typing):
**Copy deck spec** (`extract_spec` → clipboard) and **Scan palette**
(`scan_palette` → clipboard). Both operate on the whole deck and work in
the ACTIVE-SLIDE and ALL-SLIDES export modes alike.

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
  modActionsText.bas              ← granular text actions (12)
  modActionsRun.bas               ← run-level formatting + hyperlink (10)
  modActionsLayout.bas            ← layout/align/distribute/recolor (24)
  modActionsTable.bas             ← table row/col/merge (5)
  modActionsChart.bas             ← chart type/title/axis/legend/series (5)
  modActionsImage.bas             ← insert/replace picture (2)
  modActionsConnector.bas         ← connector (1)
  modActionsGroup.bas             ← group/ungroup (2)
  modActionsSlide.bas             ← move/extract/import slides (3)
  modActionsDeck.bas              ← deck-wide ops (regex/font swap/recolor/theme/size/bulk) (9)
  modActionsEffects.bas           ← visual polish + picture effects (16)
  modActionsTemplate.bas          ← apply_template: 7 builtin slide layouts
  modActionsSpec.bas              ← decks-as-code: build_deck_from_spec/extract_spec/generate_variants
  modActionsCapture.bas           ← Deck DNA: capture/list/delete/rename_template + registry JSON
  frmExport.frm/.frx              ← snapshot UserForm
  frmExecute.frm/.frx             ← instructions UserForm
  frmImportSlides.frm/.frx        ← import UserForm
update_macros.py                  ← sync src/ → carrier
tools/build_carrier.py            ← bootstrap empty carrier
tools/build_forms.py              ← rebuild UserForm layout; preserves src/*.frm code
tools/add_fix_button.py           ← idempotent installer: Fix Errors/Fix This on frmExecute
tools/add_export_buttons.py       ← idempotent installer: Copy deck spec / Scan palette on frmExport
tools/inspect_form.py             ← inspect UserForm controls via COM
tools/screenshot_forms.py         ← DPI-aware screenshot of each UserForm
tools/rebuild_import_slides.py    ← escape hatch for frmImportSlides rebuild
tools/precheck_carrier.py         ← carrier sanity check
tests/                            ← smoke harness + deck generator
test_decks/                       ← deterministic test inputs
docs/specs/                       ← Phase 1 + Phase 2 design specs
docs/superpowers/plans/           ← implementation plans
docs/superpowers/specs/           ← post-Phase-2 design specs (granular text, ...)
```

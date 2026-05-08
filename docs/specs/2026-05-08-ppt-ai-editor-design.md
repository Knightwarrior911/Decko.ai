# PPT AI Editor — Design Spec

**Date:** 2026-05-08
**Status:** Approved (pending user review of this document)
**Location:** `C:\Users\vinit\Documents\PPT_AI_Editor\`

## 1. Overview

A VBA-based system that edits PowerPoint presentations using natural language, by bridging an LLM (accessed via the user's employer-provided LLM aggregator) and PowerPoint VBA. The LLM is the brain (understands requests, generates structured action instructions). VBA is the hands (executes actions on the active deck).

No Python at runtime, no server, no API connection from PowerPoint. Two human steps: copy a snapshot to the LLM, paste the LLM's instructions back. Python is used only at build time to sync `.bas`/`.frm` source into the carrier `.pptm`.

## 2. Goals & non-goals

**Goals:**
- Edit any open PowerPoint deck via natural-language requests routed through any external LLM.
- Capture deck state with enough fidelity that an LLM can reason about layout, color, fonts, tables, and pictures — not just text.
- Apply edits with a backup, dry-run preview, and append-only action log.
- Cover 15 action types in V1 (text, format, geometry, slides, tables).

**Non-goals:**
- No direct API calls from VBA. The user pastes through their employer's LLM aggregator UI.
- No automatic round-trip / agent loop. Human is in the loop on each turn.
- No chart-internal editing in V1 (charts are inspected as opaque shapes; future work).
- No theme recolor or master-slide editing in V1.

## 3. Architecture

```
Documents/PPT_AI_Editor/
├── PPT_AI_Editor.pptm           ← carrier; macros loaded from this file
├── src/
│   ├── modJSON.bas              ← VBA-JSON (Tim Hall, vendored, MIT)
│   ├── modExportSnapshot.bas    ← snapshot builder
│   ├── modExecuteInstructions.bas ← action dispatcher + validation
│   ├── modActions.bas           ← one Sub per action type
│   ├── modBackup.bas            ← auto-backup + JSONL log writer
│   ├── modUI.bas                ← entry-point Subs (registered in macro list)
│   ├── frmExport.frm            ← export UserForm
│   ├── frmExport.frx            ← form binary
│   ├── frmExecute.frm           ← execute UserForm (with dry-run preview)
│   └── frmExecute.frx           ← form binary
├── update_macros.py             ← sync .bas/.frm → .pptm via pywin32 COM
├── tests/
│   └── run_smoke.py             ← Python smoke harness using COM
├── test_decks/
│   ├── smoke_3slide.pptx        ← minimal text-only deck
│   └── full_visual.pptx         ← deck with table + picture + colors
├── docs/
│   └── specs/
│       └── 2026-05-08-ppt-ai-editor-design.md   ← this file
└── README.md
```

### 3.1 Data flow

1. User opens any target deck plus the `PPT_AI_Editor.pptm` carrier.
2. Alt+F8 → `ExportSnapshot` → `frmExport` opens with the snapshot JSON visible.
3. User clicks **Copy snapshot + prompt template** → clipboard now holds a ready-to-paste block (snapshot JSON + prompt scaffold).
4. User pastes into the LLM aggregator UI, fills in the natural-language request, submits.
5. LLM returns instructions JSON.
6. Alt+F8 → `ExecuteInstructions` → `frmExecute` opens.
7. User pastes instructions JSON → clicks **Parse**. The form lists every parsed action with a status (`ok` / `invalid: <reason>`).
8. User clicks **Apply** → backup runs first; valid actions execute in order; log lines append per action.
9. Form shows summary: `N applied, M skipped. Log: <path>. Backup: <path>.`

## 4. JSON contracts

### 4.1 Snapshot JSON (full visual fidelity)

```json
{
  "deck": {
    "path": "C:\\Users\\vinit\\Desktop\\X.pptx",
    "slide_width_pt": 960,
    "slide_height_pt": 540,
    "theme": {
      "accent1": "#1F4E79", "accent2": "#2E75B6",
      "accent3": "#9DC3E6", "accent4": "#BDD7EE",
      "accent5": "#DEEBF7", "accent6": "#A9D18E",
      "dk1": "#000000", "lt1": "#FFFFFF"
    }
  },
  "slides": [
    {
      "slide_number": 1,
      "layout_name": "Title Slide",
      "shapes": [
        {
          "shape_id": 3,
          "shape_name": "Title 1",
          "type": "title",
          "pos": {"left": 50, "top": 40, "width": 860, "height": 80},
          "text": "Q3 Results",
          "font": {"name": "Calibri", "size": 44, "bold": true,
                   "italic": false, "color": "#1F4E79"},
          "fill": null
        },
        {
          "shape_id": 5,
          "shape_name": "Table 4",
          "type": "table",
          "pos": {"left": 50, "top": 150, "width": 860, "height": 300},
          "table": {
            "rows": 4,
            "cols": 3,
            "cells": [
              [
                {"text": "Metric", "font": {"size": 14, "bold": true,
                                              "color": "#FFFFFF"}, "fill": "#1F4E79"},
                {"text": "Q2", "font": {"size": 14, "bold": true,
                                          "color": "#FFFFFF"}, "fill": "#1F4E79"},
                {"text": "Q3", "font": {"size": 14, "bold": true,
                                          "color": "#FFFFFF"}, "fill": "#1F4E79"}
              ]
            ]
          }
        },
        {
          "shape_id": 7,
          "shape_name": "Picture 6",
          "type": "picture",
          "pos": {"left": 50, "top": 480, "width": 200, "height": 40},
          "picture": {"filename": "logo.png"}
        }
      ]
    }
  ]
}
```

**Type domain:** `title`, `body`, `textbox`, `table`, `picture`, `chart`, `other`.

**Conventions:**
- All positions in points (1 inch = 72 pt). Computed via `Shape.Left / Shape.Top / Shape.Width / Shape.Height`.
- Colors as `#RRGGBB` hex. If the color is theme-bound, value is `null` and an optional `font.theme_color` / `fill_theme_color` slot name is added (`accent1` etc.).
- Multi-paragraph text joined with `\n`.
- Only top-level shapes captured. Grouped shapes' children are not exposed in V1 — the group is reported as a single `other`-type shape with bounding pos.

### 4.2 Instructions JSON

```json
{
  "actions": [
    {"type": "set_text", "slide": 1, "shape_id": 3, "value": "Q3 2026 Board Update"},
    {"type": "set_font_size", "slide": 1, "shape_id": 3, "value": 36},
    {"type": "set_font_bold", "slide": 1, "shape_id": 3, "value": true},
    {"type": "set_font_italic", "slide": 1, "shape_id": 3, "value": false},
    {"type": "set_font_color", "slide": 1, "shape_id": 3, "value": "#FF0000"},
    {"type": "set_fill_color", "slide": 1, "shape_id": 5, "value": "#1F4E79"},
    {"type": "move_shape", "slide": 1, "shape_id": 5, "left": 100, "top": 200},
    {"type": "resize_shape", "slide": 1, "shape_id": 5, "width": 700, "height": 250},
    {"type": "delete_shape", "slide": 1, "shape_id": 7},
    {"type": "add_slide", "position": 3, "layout_index": 1},
    {"type": "delete_slide", "slide": 4},
    {"type": "duplicate_slide", "slide": 2},
    {"type": "set_cell_text", "slide": 1, "shape_id": 5, "row": 2, "col": 1, "value": "Revenue"},
    {"type": "swap_table_columns", "slide": 1, "shape_id": 5, "col_a": 1, "col_b": 2},
    {"type": "swap_table_rows", "slide": 1, "shape_id": 5, "row_a": 1, "row_b": 2}
  ]
}
```

**Conventions:**
- Slide and shape numbers are 1-based. Table row/col are 1-based.
- `slide` + `shape_id` together identify a shape (`Shape.Id` is unique per slide, not per deck).
- Colors as `#RRGGBB`.
- Lengths in points.
- Boolean values are real JSON `true`/`false`.

## 5. V1 action set

| Action | Required keys | Notes |
|---|---|---|
| `set_text` | slide, shape_id, value | Replaces all text in shape's TextFrame. |
| `set_font_size` | slide, shape_id, value (int) | Applies to all runs in TextFrame. |
| `set_font_bold` | slide, shape_id, value (bool) | All runs. |
| `set_font_italic` | slide, shape_id, value (bool) | All runs. |
| `set_font_color` | slide, shape_id, value (hex) | All runs. |
| `set_fill_color` | slide, shape_id, value (hex) | Solid fill. |
| `move_shape` | slide, shape_id, left, top | Points. |
| `resize_shape` | slide, shape_id, width, height | Points. |
| `delete_shape` | slide, shape_id | |
| `add_slide` | position, layout_index | layout_index into `SlideMaster.CustomLayouts`. |
| `delete_slide` | slide | |
| `duplicate_slide` | slide | New slide placed immediately after source. |
| `set_cell_text` | slide, shape_id, row, col, value | Shape must be a table. |
| `swap_table_columns` | slide, shape_id, col_a, col_b | Swaps text + per-cell formatting. |
| `swap_table_rows` | slide, shape_id, row_a, row_b | Swaps text + per-cell formatting. |

## 6. Components

### 6.1 `modJSON`
Vendored VBA-JSON (https://github.com/VBA-tools/VBA-JSON). Public functions used: `JsonConverter.ParseJson(jsonString) As Object`, `JsonConverter.ConvertToJson(value, [indent]) As String`. MIT license; copyright header preserved.

### 6.2 `modExportSnapshot`
- `Public Function BuildSnapshotJson() As String` — returns JSON string for `ActivePresentation`.
- `Private Function BuildShapeDict(sh As Shape) As Object` — builds one shape's dictionary entry. Switches on `sh.Type` and `sh.HasTextFrame`/`sh.HasTable`.
- Helpers: `RGBToHex(rgbLong)`, `PtToPt` (no-op, sanity wrapper), `EscapeText`.

### 6.3 `modExecuteInstructions`
- `Public Sub ExecuteFromString(jsonText As String)` — parse, validate, dispatch.
- `Private Function ValidateAction(act As Object, presentation As Presentation) As String` — empty string if valid, else reason.
- `Private Sub DispatchAction(act As Object)` — `Select Case act("type")` → call `modActions.Do_<type>`.
- Maintains running counts: applied, skipped. Returns summary string.
- All execution wrapped in `On Error Resume Next` per action so one failure does not abort the batch.

### 6.4 `modActions`
One Sub per action type. Pure execution given already-validated input. Examples:
- `Sub Do_set_text(slide_num As Long, shape_id As Long, value As String)`
- `Sub Do_move_shape(slide_num As Long, shape_id As Long, left_pt As Single, top_pt As Single)`
- `Sub Do_swap_table_columns(slide_num As Long, shape_id As Long, col_a As Long, col_b As Long)`

Lookup helper: `Function FindShape(slide_num, shape_id) As Shape` — `Nothing` if not found.

### 6.5 `modBackup`
- `Function BackupActiveDeck() As String` — copies `ActivePresentation.FullName` to `<base>_backup_<yyyy-mm-dd_hhmmss>.<ext>` via `FileSystemObject.CopyFile`. Returns destination path. Raises clean error if save path is read-only or unsaved.
- `Sub LogAction(deckPath As String, action As Object, status As String, reason As String)` — appends one JSON line to `<deckPath>.action_log.jsonl`. UTF-8.

### 6.6 `modUI`
Entry points registered as macros (visible in Alt+F8):
- `Sub ExportSnapshot()` — `frmExport.Show vbModeless`.
- `Sub ExecuteInstructions()` — `frmExecute.Show vbModeless`.

### 6.7 `frmExport`
Controls:
- `txtSnapshot` (multi-line, read-only) — populated from `modExportSnapshot.BuildSnapshotJson` on `UserForm_Initialize`.
- `btnCopySnapshot` — copies raw JSON to clipboard via `MSForms.DataObject`.
- `btnCopyWithTemplate` — copies prompt template (see §7) with snapshot JSON inlined.
- `btnSaveTxt` — writes `<deck>_snapshot_<ts>.txt` next to deck.
- `btnClose`.

### 6.8 `frmExecute`
Controls:
- `txtInstructions` (multi-line, editable) — user pastes instructions JSON here.
- `btnParse` — runs `modJSON.ParseJson` + validation. Populates `lstActions` (slide / shape_id / type / status). Enables `btnApply` only if at least one action is valid.
- `lstActions` — listbox showing parsed action preview. Invalid rows shown red-tagged.
- `btnApply` — disabled until parse succeeds; on click runs backup → execute loop → updates `lblStatus`.
- `btnCancel` — closes form, no changes.
- `lblStatus` — final summary + paths.

## 7. Prompt template (Copy with template button)

The button copies the following to clipboard, with `{snapshot}` filled in:

```
You are editing a PowerPoint presentation. Below is the current state as JSON:

```json
{snapshot}
```

I want the following changes:

[REPLACE THIS LINE WITH YOUR REQUEST]

Return ONLY a valid instructions JSON in this exact format. No prose, no
explanation, no markdown fences:

{
  "actions": [
    {"type": "<action_type>", "slide": <int>, "shape_id": <int>, ...}
  ]
}

Rules:
- Use only shape_ids that exist in the snapshot. Do not invent ids.
- Slide numbers are 1-based.
- Colors as #RRGGBB hex.
- Lengths in points.
- Allowed action types: set_text, set_font_size, set_font_bold,
  set_font_italic, set_font_color, set_fill_color, move_shape,
  resize_shape, delete_shape, add_slide, delete_slide, duplicate_slide,
  set_cell_text, swap_table_columns, swap_table_rows.
```

## 8. Error handling

### 8.1 Validation gate (pre-execution)
Each action is validated before any edit runs. Reasons captured per action:
- `invalid_json` — top-level parse failure → no actions executed at all.
- `unknown_type` — action type not in V1 set.
- `missing_field: <name>` — required key absent.
- `slide_out_of_range` — `slide` > slide count or < 1.
- `shape_not_found` — `shape_id` absent on target slide.
- `shape_wrong_type` — table action on a non-table shape, etc.

User sees per-row status in the Execute form preview before any change is applied.

### 8.2 Auto-backup
Before any valid action runs, `BackupActiveDeck()` is invoked. If backup fails, the entire batch aborts and the form shows a clear error. No edits.

### 8.3 Per-action execution
Wrapped in `On Error Resume Next`. Outcome (`ok` or `error: <Err.Description>`) recorded. Loop continues. One bad action does not poison the batch.

### 8.4 Action log
Append-only `<deck>.action_log.jsonl` next to deck. Format:

```json
{"ts":"2026-05-08T14:30:15Z","op":"set_text","slide":1,"shape_id":3,"params":{"value":"Q3"},"status":"ok"}
{"ts":"2026-05-08T14:30:15Z","op":"move_shape","slide":1,"shape_id":99,"params":{"left":100,"top":200},"status":"skipped","reason":"shape_not_found"}
```

## 9. Build / sync (`update_macros.py`)

Mirrors the BrandRethemer pattern. Run from repo root:

```bash
python update_macros.py
```

Behavior:
1. Open `PPT_AI_Editor.pptm` headlessly via `win32com.client.DispatchEx("PowerPoint.Application")`.
2. For each `.bas` / `.frm` file under `src/`:
   - If a code module of the same name exists in the VBProject, remove it.
   - Import the file via `VBProject.VBComponents.Import(path)`.
3. Save and close the presentation.

Requires "Trust access to the VBA project object model" enabled in PowerPoint Trust Center (one-time setup; documented in README).

## 10. Testing

### 10.1 Automated smoke (`tests/run_smoke.py`)
Uses pywin32 COM to drive PowerPoint:
1. Open `test_decks/smoke_3slide.pptx` and `PPT_AI_Editor.pptm`.
2. Call `Application.Run "BuildSnapshotJson"` → assert valid JSON, 3 slides, expected shape_ids and types.
3. Construct an instructions JSON in Python exercising every V1 action type. Pass via `Application.Run "ExecuteFromString", json_text`.
4. Re-export snapshot → assert mutations visible (text changed, shape moved, table cell swapped).
5. Verify `<deck>_backup_*.pptm` exists, `<deck>.action_log.jsonl` non-empty with expected entries.
6. Negative cases: invalid JSON → log unchanged; unknown shape_id → action skipped, others succeed.

### 10.2 Manual checklist (in README)
- Carrier opens cleanly, macros visible in Alt+F8.
- ExportSnapshot form populates with valid JSON for each test deck.
- Copy-with-template produces paste-ready block.
- ExecuteInstructions: valid JSON parses cleanly, preview shows actions, Apply executes and updates summary.
- Invalid JSON: clear error, Apply disabled.
- Backup file appears with expected name.
- Log file appended.
- All 15 action types verified at least once on `full_visual.pptx`.

## 11. Constraints & assumptions

- Windows + PowerPoint (Office 2016 or newer recommended).
- 64-bit Office supported (VBA-JSON works in both bitnesses; `MSScriptControl` deliberately avoided).
- The active presentation must be the deck the user wants to edit (not the carrier itself). The macros target `ActivePresentation`.
- The user must enable "Trust access to the VBA project object model" once for `update_macros.py` to work (developer machine only; end users do not need this).
- LLM access is via the user's employer-provided aggregator. No assumption about which model — prompt template is provider-agnostic.

## 12. Out of scope (V2+)

- Theme recolor, master-slide editing.
- Group-shape internals.
- Chart-internal data edits.
- Per-run formatting within a TextFrame (V1 applies font ops to entire TextFrame).
- Animations, transitions.
- Speaker notes editing.
- Round-trip "do same on slide N" macro replay (action log enables this; building it is future work).

## 13. Success criteria

- User opens any deck + carrier, runs ExportSnapshot, copies snapshot+template, pastes into LLM, describes a change, gets back instructions JSON, pastes into Execute form, clicks Apply, and the deck reflects the requested change.
- Invalid LLM output is caught at the validation gate, never corrupts the deck.
- Every applied edit has a backup and a log entry.
- All 15 V1 action types work end-to-end on `test_decks/full_visual.pptx`.

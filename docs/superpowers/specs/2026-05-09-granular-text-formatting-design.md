# Granular Text Formatting — Design Spec

**Status:** Approved (2026-05-09). Ready for implementation plan.
**Owner:** Decko.ai backend.
**Touches:** `src/modExportSnapshot.bas`, `src/modExecuteInstructions.bas`, `src/modActionsText.bas`, new `src/modActionsRun.bas`, `tests/`, `tests/make_test_decks.py`, `README.md`.

## Goal

Eliminate the "can't format mid-paragraph" gap so a VP+ user has confidence the tool can express any reasonable text edit on a slide.

Today's 56 actions cover whole-shape and whole-paragraph text changes. They cannot:

- Bold one word inside a sentence while leaving siblings untouched.
- Set a hyperlink on a substring.
- Mix font sizes/colors/fonts within a paragraph.
- Set paragraph alignment, line spacing, vertical anchor, or text-frame margins.
- See where formatting changes inside a paragraph (snapshot v2 collapses runs).

This spec adds 15 actions and a Snapshot v3 that exposes per-run formatting to the LLM.

## Scope

In:

- 15 new actions across runs, paragraphs, and text frames.
- Snapshot v3: paragraphs gain `runs[]`; shapes gain `text_frame{}`; paragraphs gain `alignment`, `line_spacing`, `space_before`, `space_after`.
- Backwards compatibility: snapshot v3 is a superset of v2; `paragraphs[].text` (concatenated string) stays so v2 parsers still work.
- No UserForm changes.
- New fixture deck `text_v3.pptx` and ~22 new test assertions.

Out (deferred):

- Tab stops, custom underline colors/styles.
- Font themes beyond plain font-name swap.
- Bidirectional/RTL text.
- Comment annotations on runs.

## Snapshot v3 schema

Per text shape:

```json
{
  "shape_id": 5,
  "kind": "text",
  "pos": {"left": 0, "top": 0, "width": 0, "height": 0},
  "text_frame": {
    "vertical_align": "middle",
    "word_wrap": true,
    "auto_size": "none",
    "margin": {"left": 7.2, "right": 7.2, "top": 3.6, "bottom": 3.6}
  },
  "paragraphs": [
    {
      "text": "Revenue grew 23% in Q3",
      "bullet": "none",
      "indent": 0,
      "alignment": "left",
      "line_spacing": 1.0,
      "space_before": 0,
      "space_after": 0,
      "runs": [
        {"text": "Revenue ", "bold": false, "italic": false, "underline": false, "strike": false, "subscript": false, "superscript": false, "size": 18, "font": "Calibri", "color": "#000000", "hyperlink": null},
        {"text": "grew 23%", "bold": true,  "italic": false, "underline": false, "strike": false, "subscript": false, "superscript": false, "size": 18, "font": "Calibri", "color": "#000000", "hyperlink": null},
        {"text": " in Q3",   "bold": false, "italic": false, "underline": false, "strike": false, "subscript": false, "superscript": false, "size": 18, "font": "Calibri", "color": "#000000", "hyperlink": null}
      ]
    }
  ]
}
```

Rules:

- `paragraphs[].text` = concatenation of all `runs[].text` for that paragraph (kept for v2 parsers).
- `runs[]` is in display order, 0-indexed. The index is the `run_index` argument used by all run actions.
- `hyperlink` is `null` when no link, otherwise the URL string.
- `line_spacing` always emitted as a multiple. Single-line = `1.0`, 1.5 line = `1.5`. When PowerPoint stores spacing in points (LineRuleWithin = msoFalse), the snapshot converts to a multiple by dividing by the paragraph's first-run font size; a mixed-size paragraph therefore reports a slight approximation. The action `set_paragraph_line_spacing` always writes back as a multiple (LineRuleWithin = msoTrue).
- `margin.*` in points.
- Empty paragraphs serialize as `"runs": []` with `"text": ""`.
- Builder: extend `BuildShapeDict` in `modExportSnapshot.bas`. The `text_frame` block is only emitted when the shape has a `TextFrame`.

## Action surface (15 new)

### Run-level (10) — `modActionsRun.bas` (new)

Common args: `slide_num`, `shape_id`, `paragraph_index`, `run_index` (all 0-indexed; dispatcher converts to PowerPoint's 1-indexed API).

| Action | Trailing args | Effect |
|---|---|---|
| `set_run_bold` | `value: bool` | Toggle bold on run. |
| `set_run_italic` | `value: bool` | Toggle italic. |
| `set_run_underline` | `value: bool` | Toggle underline. |
| `set_run_strikethrough` | `value: bool` | Toggle strikethrough. |
| `set_run_subscript` | `value: bool` | Mutually exclusive with superscript; setting one auto-clears the other. |
| `set_run_superscript` | `value: bool` | Same. |
| `set_run_font_color` | `hex: "#RRGGBB"` | Set run color. |
| `set_run_font_size` | `pt: int` | Set run size. |
| `set_run_font_name` | `name: string` | Set run font name. |
| `set_run_text` | `value: string` | Replace this run's chars; surrounding runs untouched. |

### Run hyperlink (1) — `modActionsRun.bas`

| Action | Trailing args | Effect |
|---|---|---|
| `set_run_hyperlink` | `url: string \| null` | Set or clear hyperlink on run. |

### Paragraph + frame (4) — appended to `modActionsText.bas`

| Action | Args | Effect |
|---|---|---|
| `set_paragraph_alignment` | `slide, shape, paragraph_index, align: "left"\|"center"\|"right"\|"justify"` | Paragraph alignment. |
| `set_paragraph_line_spacing` | `slide, shape, paragraph_index, multiple: float` | Paragraph line spacing as multiple (1.0, 1.5, 2.0...). |
| `set_text_vertical_align` | `slide, shape, anchor: "top"\|"middle"\|"bottom"` | Whole-shape vertical anchor. |
| `set_text_margin` | `slide, shape, left, right, top, bottom` (all pt floats) | Whole-shape internal margins. |

Total carrier action count goes 56 → 71.

### Implementation notes

- Run resolution: `tf.Paragraphs(paragraph_index + 1).Runs(run_index + 1)`.
- `set_run_text`: TextRange replace. Doesn't fragment the run; if new text is shorter/longer, neighbor runs are unaffected because they live in different `Runs()` index slots.
- Hyperlink: `Run.ActionSettings(ppMouseClick).Hyperlink.Address = url`. `null` clears via `Address = ""`.
- Sub/superscript: `Run.Font.BaselineOffset` (-0.25 for sub, +0.30 for super, 0 for off).
- Line spacing: `Paragraph.ParagraphFormat.SpaceWithin` (multiple) with `LineRuleWithin = msoTrue`.
- Vertical align: `Shape.TextFrame.VerticalAnchor` (`msoAnchorTop` / `msoAnchorMiddle` / `msoAnchorBottom`).
- Alignment: `Paragraph.ParagraphFormat.Alignment` (`ppAlignLeft` / `ppAlignCenter` / `ppAlignRight` / `ppAlignJustify`).

## Validation, dispatch, error model

### Validation (extends `modExecuteInstructions.bas`)

Every new action validates BEFORE dispatch. On failure, the row is skipped, logged with reason, and the batch continues.

Common shape-target checks:

- `slide_num` resolves to a slide.
- `shape_id` resolves to a shape on that slide.
- Shape has a `TextFrame` (otherwise skip with `"shape has no text"`).

Run-target additional checks:

- `paragraph_index >= 0` and `< Paragraphs.Count`.
- `run_index >= 0` and `< Runs.Count` for that paragraph.
- Out of range → skip with `"paragraph/run index out of range"`.

Per-action arg checks:

| Action | Validation |
|---|---|
| `set_run_*_color` | hex matches `#RRGGBB`. |
| `set_run_font_size`, `set_paragraph_line_spacing` | numeric, > 0. |
| `set_run_font_name` | non-empty string. PowerPoint substitutes silently if font missing — no pre-flight check (too slow). |
| `set_run_hyperlink` | URL string or `null`. If string, must look like `http://`, `https://`, `mailto:`, or in-deck fragment `#slide:N`. |
| `set_paragraph_alignment` | one of 4 enums. |
| `set_text_vertical_align` | one of 3 enums. |
| `set_text_margin` | all 4 floats present, >= 0. |

### Dispatch

15 new `Case` arms in the existing dispatcher Select-Case. Each:

1. Pulls typed args from the action dict.
2. Calls the matching `Do_*` Sub.
3. On VBA error: log `"ERROR: <Err.Description>"`, increment `skipped`. Else increment `applied`.

### Error model

- Bad arg → row skipped, logged, batch continues. Same as today's actions.
- Out-of-range index → skipped, logged.
- Missing font → silently substituted by PowerPoint.
- Malformed hyperlink URL → skipped with `"invalid hyperlink URL"`.

### Reverse-order processing inside same paragraph

When multiple run-level actions target the same paragraph in one batch, the dispatcher pre-sorts them by `run_index DESCENDING` so earlier runs' index shifts (from `set_run_text` shrinking/growing) don't invalidate later actions.

Implementation: group incoming actions by `(slide_num, shape_id, paragraph_index)`, sort within group by descending `run_index`, then run.

## Tests

### New fixture deck

`tests/make_test_decks.py` gains `make_text_v3(path)`. Generated file: `test_decks/text_v3.pptx`.

- Slide 1: heading with a mixed-format paragraph (`"Revenue [bold]grew 23%[/bold] in Q3"`).
- Slide 2: bulleted list, mixed sizes per bullet, one bullet with hyperlink, one with strikethrough.
- Slide 3: empty paragraphs + paragraphs with sub/superscript (`"H₂O"`, `"E=mc²"`).
- Slide 4: shape with non-default text frame (vertical-middle anchor, custom margins, line spacing 1.5).

### Snapshot v3 assertions (extend `tests/run_smoke.py`)

For the `text_v3.pptx` snapshot:

- Paragraph 0 of slide 1 has 3 runs.
- Run 1 has `bold: true`; runs 0 and 2 have `bold: false`.
- Hyperlink run carries `hyperlink: "https://..."`.
- Sub/superscript runs report `subscript: true` / `superscript: true`.
- Slide 4 shape has `text_frame.vertical_align == "middle"`.

### Action smoke tests (new `tests/run_smoke_text.py`)

One test per action (15 total). Pattern matches Phase 2 smokes:

1. Open `text_v3.pptx`.
2. Run action via `app.Run("PPT_AI_Editor!Do_<action>", ...)`.
3. Re-snapshot.
4. Assert expected run/paragraph/frame state.

Plus 2 cross-action tests:

- Reverse-order: 3 `set_run_text` actions on same paragraph (run 0, 1, 2) in a single batch; assert all applied without index drift.
- Round-trip: bold → unbold → bold again on the same run; assert idempotent.

### Total

~22 new assertions on top of the current 81. Target post-merge: ~103 assertions all green.

## Migration

- Snapshot v3 is a superset of v2. The existing `paragraphs[].text` field stays; new fields are additive. Existing LLM prompt templates keep working without changes.
- README "Action types" section grows from 56 → 71. Update the table.
- No UserForm changes; no carrier rebuild beyond `update_macros.py` syncing the new `.bas` file and the modified ones.

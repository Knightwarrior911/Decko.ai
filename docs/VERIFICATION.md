# Decko.ai — Verification Loop & Fix Buttons

After every batch of actions you Apply, Decko sweeps the deck and reports
quality problems automatically. Two clipboard-ready buttons on the Execute
form let you hand findings to an LLM in one click — no file-opening, no
hand-typed feedback.

## What runs automatically

After every successful `ExecuteInstructions` Apply:

1. All actions in the batch dispatch top-to-bottom against the active deck.
2. `modVerify.RunVerificationLoop("deck", 100)` sweeps every slide.
3. Findings (up to 100 per batch) are written to two places:
   - The Execute form's status bar shows `N warning(s), M info`.
   - A sidecar JSON file is written at `<deck-path>.warnings.json` with
     the full payload.

Cost: ~250-400 ms on a typical 10-slide deck. Capped behavior makes large
decks safe (skip per-shape deep checks if a slide has >100 shapes, skip
chart deep-inspect if chart has >20 series, etc.).

## What gets checked (32 checks across 8 domains)

### Layout / geometry (8)

| Check | Triggers when |
|---|---|
| `off_slide_shape` | Shape extends past left/right/top/bottom slide edge |
| `duplicate_position` | Two shapes with exact same bounds |
| `zero_size_shape` | Width or height < 1 pt |
| `crowded_slide` | >40 shapes on one slide |
| `cramped_to_edge` | Shape within 12 pt of slide edge |
| `inconsistent_row_alignment` | 3+ shapes share Y-band but tops differ 1-5 pt |
| `inconsistent_col_alignment` | Same for X-band |
| `chart_covered_by_shape` | Chart fully enclosed by higher-Z shape |

### Text (8)

| Check | Triggers when |
|---|---|
| `text_overflow` | Text bounds exceed shape frame; skipped if shrink-to-fit on |
| `empty_shape` | No fill, no line, no text (invisible orphan) |
| `shape_text_contrast` | Font/fill contrast below WCAG AA (4.5) |
| `tiny_shape_font` | Font < 8 pt |
| `very_large_body_font` | Font > 40 pt on non-title shape |
| `mixed_font_families` | 3+ different fonts in one shape |
| `placeholder_text_present` | "Click to add title/text" never replaced |
| `trailing_whitespace` | Text ends with 3+ trailing spaces |

### Chart (7)

| Check | Triggers when |
|---|---|
| `chart_label_contrast` | Data labels < AA contrast vs series fill |
| `chart_no_title` | Chart has no title text |
| `chart_all_zero_values` | Every series value is 0 (missing data) |
| `pie_too_many_slices` | Pie/doughnut with >8 categories |
| `chart_default_series_name` | Series still named "Series 1" placeholder |
| `chart_pointless_legend` | 1 series + legend visible |
| `chart_axis_clips_data` / `chart_axis_excess_headroom` | Value-axis range mismatched with data magnitude |

### Table (3)

| Check | Triggers when |
|---|---|
| `tiny_table_font` | Cell font < 8 pt |
| `cell_text_contrast` | Cell font/fill contrast below AA |
| `table_column_overflow` | Sum of column widths > table width |

### Slide (4)

| Check | Triggers when |
|---|---|
| `slide_no_title` | Layout has no title placeholder (TOC/screen-reader gap) |
| `slide_empty_title` | Title placeholder present but empty or "Click to add..." |
| `too_many_colors` | >10 distinct fill colors on one slide |
| `too_many_fonts` | >3 distinct font families on one slide |
| `duplicate_text_content` | Two shapes on same slide with identical text |

### Picture (1)

| Check | Triggers when |
|---|---|
| `picture_no_alt_text` | Picture without alt text (accessibility gap) |

### Hyperlinks (1)

| Check | Triggers when |
|---|---|
| `broken_internal_hyperlink` | `#slide:N` where N > deck slide count |

### Connectors (1)

| Check | Triggers when |
|---|---|
| `orphan_connector` | Connector with BeginConnected=false or EndConnected=false |

## Warning payload format

Each warning is a JSON object with five fields:

```json
{
  "severity": "warn" | "info",
  "kind": "off_slide_shape",
  "slide": 3,
  "shape_id": 17,
  "message": "Rectangle 1 extends 90 pt past RIGHT edge of slide",
  "suggestion": "move or resize shape inside slide bounds (slide is 960 x 540 pt)"
}
```

The **`suggestion`** field is the killer feature: it carries either a
literal action signature the LLM can use as-is, or a one-line natural-
language hint. Examples:

- `set_cell_font_color slide=4 shape_id=2 row=2 col=2 value=#FFFFFF`
- `set_chart_series slide=3 shape_id=2 series_index=1 props.label_color=#FFFFFF`
- `recolor_palette_deck_wide for too-many-colors slide`

## The two "Fix" buttons

The Execute form has four bottom-row buttons: `Fix Errors`, `Fix This`,
`Cancel`, `Apply`. The two Fix buttons close the two halves of the
LLM-to-Decko feedback loop.

### Fix Errors (pre-Apply)

When `Parse` shows red INVALID lines (missing fields, bad enum values,
out-of-range slides), click **Fix Errors**. Clipboard receives a prompt
with:

```
--- ACTION 1 (type: set_paragraph_text) ---
YOU SENT: {"type":"set_paragraph_text","slide":1,"shape_id":5,
           "paragraph_index":0,"text":"Hello"}
ERROR: missing_field: value
CORRECT SHAPE:
  REQUIRED: slide, shape_id, paragraph_index (0-based int), value(string)
  EXAMPLE:  {"type":"set_paragraph_text","slide":1,"shape_id":3,
             "paragraph_index":0,"value":"Hello"}
```

Paste into your LLM chat. The LLM rewrites the failing actions and you
loop back to Parse.

Behind the scenes: `modExecuteInstructions.BuildErrorFixPrompt(jsonText)`
runs `PreviewValidate` on each action, then looks up
`GetActionGuidance(actionType)` for the canonical signature of any failing
type.

**Coverage:** every one of the ~165 known action types has a canonical
guidance entry (REQUIRED fields + working example) in `GetActionGuidance`.
The coverage test `tests/test_guidance_coverage.py` enforces this — it
parses every Case label in `DispatchAction` and asserts each one returns
non-fallback guidance from `GetActionGuidance`. CI-grade safety net so a
new action added to the dispatcher without matching guidance trips the
test, not the user's broken batch.

**Doc-sync:** the public `docs/ACTIONS_REFERENCE.md` includes an
auto-generated appendix containing the canonical signature for every
action. The appendix is regenerated by
`python tools/sync_actions_guidance.py`, which invokes `GetActionGuidance`
in VBA for each known type and writes the result into the markdown between
`<!-- BEGIN AUTO-GUIDANCE -->` / `<!-- END AUTO-GUIDANCE -->` markers.
Drift detector: `tests/test_guidance_doc_sync.py` fails if the committed
markdown disagrees with what VBA currently returns. Keeps the in-app Fix
Errors button and the public reference docs from ever drifting apart.

**"Did you mean" suggestions:** when the LLM emits an action with an
unknown type, the Fix Errors prompt now scores all 234 known types by
word-stem overlap and suggests the closest 3. Example: emit
`set_chart_axis_gridlines` (doesn't exist) → suggestion line reads
`DID YOU MEAN: set_chart_axis_title, set_chart_gridlines, set_chart_axis`.
Implementation: `FindSimilarActions(badType)` + `GetAllActionTypes()` in
modExecuteInstructions.bas.

### Fix This (post-Apply)

After Apply runs, the form status bar shows `N warning(s)`. Click
**Fix This** to copy a prompt like:

```
[warn] chart_label_contrast (slide 3, shape_id 12)
  ISSUE: series 1 (fill #0F285A) has data labels in #000000 — contrast 1.5
  SUGGESTION: set_chart_series slide=3 shape_id=12 series_index=1 props.label_color=#FFFFFF
```

Paste into your LLM chat. The LLM emits a fix batch using the `SUGGESTION`
lines almost verbatim. Apply again, verify re-runs, ideally 0 warnings.

Behind the scenes: `modVerify.CopyWarningsPromptToClipboard()` reads the
sidecar `<deck>.warnings.json`, runs `BuildLLMPromptFromWarnings()` to
format, drops the result on the system clipboard via `MSForms.DataObject`.

## Programmatic / scripted use

```vba
' Sweep deck programmatically (no UI):
Dim warnings As Collection
Set warnings = modVerify.RunVerificationLoop("deck", 100)
' warnings.Count, warnings(1)("kind"), warnings(1)("suggestion"), ...

' Or one-shot to clipboard:
Dim n As Long
n = modVerify.CopyWarningsPromptToClipboard()

' Or just the prompt string (no clipboard side-effect):
Dim prompt As String
prompt = modVerify.BuildLLMPromptFromWarnings("C:\path\to\deck.pptm.warnings.json")
```

From Python via COM:

```python
import json, win32com.client
app = win32com.client.Dispatch("PowerPoint.Application")
result = app.Run(
    "PPT_AI_Editor.pptm!modExecuteInstructions.ExecuteFromString",
    json.dumps({"actions": [], "verify_after": True, "verify_scope": "deck"})
)
# Result string ends with "verification: N warning(s), M info"
# Full payload at <deck>.warnings.json
```

## Opt-out / scope

Top-level fields on the instructions JSON:

| Field | Default | Meaning |
|---|---|---|
| `verify_after` | `true` | Set `false` to skip verification on this batch |
| `verify_scope` | `"deck"` | Set `"slide:N"` to verify only one slide |

Or trigger an explicit mid-batch sweep with the standalone action:

```json
{"type": "run_verification", "scope": "deck", "max_warnings": 100}
```

## Performance limits (hard-coded in modVerify.bas)

| Constant | Value | Effect |
|---|---|---|
| `MAX_WARNINGS_DEFAULT` | 100 | Hard ceiling per batch |
| `MAX_SHAPES_PER_SLIDE_FOR_DEEP` | 100 | Skip per-shape deep checks above this |
| `MAX_CHART_SERIES_FOR_DEEP` | 20 | Skip chart deep-inspection above this |
| `MAX_TABLE_CELLS_FOR_DEEP` | 200 | Skip table deep-inspection above this |
| `WCAG_CONTRAST_THRESHOLD` | 4.5 | AA body-text floor |
| `SLIDE_BOUNDS_TOL_PT` | 2.0 | Tolerance for off-slide detection |
| `TINY_TABLE_FONT_PT` | 8 | Below this = tiny-font warning |

Every per-shape check is wrapped in `On Error Resume Next`, so a single
malformed shape never aborts the sweep — it's just skipped silently.

## Limits / what verify does NOT catch

- **Subjective judgment.** "This color is ugly" or "this layout is
  boring" — not the loop's job.
- **Visual rendering quirks.** Verify reads shape properties, not pixels.
  A font that just happens to look bad at this size will pass.
- **Domain validation.** "Is $1.36B revenue correct for FY25?" — analyst
  work, not Decko's.
- **Animations / transitions / slide-show settings.** Decko doesn't manage
  these at all (per Hard Rule 16 in ACTIONS_REFERENCE.md).

## Tests

| Script | Verifies |
|---|---|
| `tests/test_verify_loop.py` | Build deck with 14+ planted problems, run sweep, all caught in <500 ms |
| `tests/test_fix_button.py` | Verify writes sidecar, button copies LLM prompt to clipboard, prompt parseable back |
| `tests/test_fix_errors_button.py` | Bad-actions batch → BuildErrorFixPrompt produces canonical-guidance prompt |

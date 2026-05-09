# Decko.ai

VBA-based PowerPoint editor driven by natural language via an external LLM.

The user copies a JSON snapshot of the active deck into their LLM tool,
describes a desired change, copies the LLM's instructions JSON back, and
clicks Apply. Two macros, two UserForms, no API calls from PowerPoint.

See `docs/specs/2026-05-08-ppt-ai-editor-design.md` for full design.

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

## Action types (V1)

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

## Files

```
PPT_AI_Editor.pptm           ← carrier (regenerated from src/)
src/                         ← VBA source-of-truth
update_macros.py             ← sync src/ → carrier
tools/build_carrier.py       ← bootstrap empty carrier
tools/build_forms.py         ← rebuild UserForms in carrier and export to src/
tests/                       ← smoke harness + deck generator
test_decks/                  ← deterministic test inputs
docs/specs/                  ← design spec
docs/superpowers/plans/      ← implementation plan
```

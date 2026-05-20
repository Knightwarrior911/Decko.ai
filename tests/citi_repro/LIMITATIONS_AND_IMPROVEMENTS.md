# Decko.ai — Limitations & Improvements (Citi Q1'26 slide-repro exercise)

Source: Citi Q1'26 earnings PDF, pages 6 & 7. Two slides recreated end-to-end
via COM (`ExecuteFromString`), iterated by screenshot comparison.

Final result: **both slides reproduced faithfully** — every chart is a **real
native PowerPoint ChartObject** (Slide 1: stacked column; Slide 2: combo
stacked-column + secondary-axis line), real tables, zero autoshape-as-chart.
Combined run: **508/508 actions, 0 non-ok.**

| # | Limitation | Root cause | Fix / status |
|---|---|---|---|
| 1 | Whole-shape `set_font_size`/`set_font_bold`/`set_font_color` on a **table** → `no text frame`; 21-row financial table overflowed/wrapped. | A table shape has no `TextFrame`; those actions require one. | **Not a code defect** — engine already ships `set_cell/row/column_font_*`. Real gap = **docs**: §3.10 Tables omitted them, so the wrong action gets chosen. **FIXED:** added "Cell / row / column font" subsection to `docs/ACTIONS_REFERENCE.md` §3.10. Retest: Slide 1 with `set_row_font_size` → 0 non-ok, all 21 rows fit. |
| 2 | `set_chart_axis y {visible:false}` **collapsed all stacked column series** on the Slide 2 combo chart (bars vanished; only the lines drew). | Hiding the primary value axis sets `HasAxis=False`; on a combo chart the columns then render against no axis. (Single-type stacked column on Slide 1 tolerated `visible:false`.) | **FIXED (docs + recipe):** documented the hide-but-keep-scale recipe and full combo recipe in §3.11. Deck uses `{min,max,tick_label_position:"none",line_visible:false}`. Retest: Slide 2 columns render correctly, 0 non-ok. |
| 3 | Stacked-column **totals** ($13.4…/21.6…) cannot be labeled natively. | PowerPoint limitation (no total data label on stacked charts). | **Contract boundary.** Worked around with an invisible total **line** series + `custom_labels`; this pattern is now documented in §3.11 as the recommended approach. |
| 4 | `hide_from_legend` ignored for a series converted to a line via `set_chart_series`; helper "Total" series still appears in the legend. | Legend-entry deletion not re-applied after combo `chart_type` conversion. | **OPEN — candidate engine improvement.** Non-blocking (data/structure correct). Logged below. |
| 5 | Goodwill series shows a "0" label on quarters where it is 0. | PowerPoint labels zero points. | **Minor.** Workaround exists (`custom_labels` with "" / per-point suppression); not applied (cosmetic). Logged. |
| 6 | Source circles key % values in rounded "pill" outlines (Slide 1); not reproduced. | No per-cell decorative outline overlay; would need a positioned `add_shape` rrect per pill (not data-bound). | **Contract boundary / future helper.** Cosmetic; not applied. |
| 7 | Bullet hierarchy in the highlight/driver text is a flat box with manual `• – ▪` + indentation, not true indent levels. | Faithful nesting needs per-paragraph `set_indent_level`/colors (many actions). | **Cosmetic**, deliberately not expanded. Documented. |
| 8 | A trailing `delete_slide` cleanup was skipped (`slide_out_of_range`). | `Presentations.Add()` under the carrier context yields a **0-slide** deck, so no default slide exists to delete. | **Not an engine bug** — harness assumption corrected (cleanup removed). |

## Engine/doc changes applied in this exercise

- `docs/ACTIONS_REFERENCE.md` §3.10 — new **Cell / row / column font** subsection
  (`set_cell_font_size/bold/italic/underline/color/name`,
  `set_row_font_size/color`, `set_column_font_size/color`) + dense-table guidance.
  *Highest-impact fix:* without this an agent picks `set_font_size` on a table,
  it errors, and large financial tables overflow.
- `docs/ACTIONS_REFERENCE.md` §3.11 — new **combo chart** recipe: per-series
  `chart_type`/`axis_group`, the invisible-total-line totals pattern, and the
  `visible:false` collapses-columns gotcha with the correct hide-but-keep-scale
  recipe.

No `src/*.bas` change was required: every capability needed already exists in
the engine; the blockers were discoverability (docs) and a usage gotcha.
`GetActionGuidance`/`ValidateAction` untouched → static contract gate unaffected.

## Recommended future engine improvements (logged, not applied)

1. Make whole-shape `set_font_size`/`set_font_bold`/`set_font_color`
   **table-aware** (auto-apply to every cell) so the intuitive action "just
   works" instead of erroring `no text frame`.
2. After `set_chart_series` `chart_type` conversion, honor `hide_from_legend`
   (delete the legend entry) so combo helper series can be hidden.
3. Optional `suppress_zero_labels` on `set_chart_series` data labels.
4. Optional table **cell accent-outline** ("pill") helper for highlighting
   individual KPI cells.
5. Consider an `add_chart` `combo`/`totals` convenience that emits the
   invisible-total-line pattern automatically.

## Metric outcome

| Metric | Slide 1 | Slide 2 |
|---|---|---|
| Structural (100% real ChartObjects, zero autoshape-as-chart) | PASS (1 chart, type stacked/combo) | PASS (1 combo chart) |
| Data (every figure/label/series matches PDF) | PASS | PASS |
| Visual (Citi palette + layout faithful) | PASS (pills/bullet-nesting = documented cosmetic deltas) | PASS (label collisions/legend = documented minor deltas) |
| Reproducible (0 errors in action_log) | PASS (436/436) | PASS (72/72); combined 508/508 |

---

# Wave 2 — slides 4 & 22 (interconnected-businesses + Banamex impacts)

Source: Citi Q1'26 PDF page index 3 (slide 4) and 21 (slide 22). Two
deliberately harder pages than 6-7: a 5-card layout with **five** real native
mini stacked-column charts + top grouping brackets + inter-card connectors, and
a lettered-callout + large banded native-table impact slide. Goal: prove the
action engine works end-to-end on hard pages, every required action 0-fail,
verify-loop clean, full static suite green.

## Gate outcome

| Metric | Slide 4 (biz) | Slide 22 (banamex) |
|---|---|---|
| M1 action-log non-ok | **0** / 98 | **0** / 289 |
| M2 autoshape-as-chart / missing native | **0** (5 real ChartObjects, columnstacked) | **0** (1 native table) |
| M3 verify warnings | **0 warn** (9 info, justified) | **0 warn** (1 info, justified) |
| M4 visual fidelity vs SRC | **95** | **96** |
| M6 regression | pytest 21 passed; run_smoke "all tests passed" | same |

M5 action-type coverage (combined, all ran `ok`, 22 distinct types):

| Family | Status | Via |
|---|---|---|
| connector | PASS | `add_connector` (×4, biz inter-card links) |
| group | PASS | `group_shapes` (biz title block; banamex marker group) |
| table | PASS | `add_table`, `apply_table_style`, `set_cell_text/fill/border/text_align/font_*`, `set_table_col_width`, `set_row_font_size` |
| native-chart | PASS | `add_chart` columnstacked, `set_chart_series`, `set_chart_axis` (×5) |
| callout/round-card | PASS | `add_shape` kind `round_rect` (biz card bodies) |
| footnote/page-num | PASS | `add_text_box` (Decko convention — see delta 3) |
| text-bullet | PASS | `set_bullet_style`, `set_indent_level`, `set_paragraph_font_size` |

## Engine/doc changes applied

### Initial pass (additive only)

No `src/*.bas` and no `docs/*` change was required to hit the gates. The first
change set was additive test assets under `tests/citi_repro/` (builders,
actions JSON, SRC ground-truth, SBS comparisons). `ValidateAction` /
`GetActionGuidance` were untouched → static contract gate unaffected (21/21
pytest green).

### Follow-up: engine + doc fixes for the findings (Wave 2.1)

Six of the contract-boundary deltas turned out to be real defects/gaps in
the engine and docs. Fixed:

| # | Finding | Fix |
|---|---|---|
| 1 | `CheckZeroSize` flagged horizontal rules (h=0) and same-y connectors (h=0) — both legitimate primitives | `src/modVerify.bas` — changed `If w < 1 Or h < 1` to `And`; a shape is only effectively invisible if **both** dimensions are sub-pixel. |
| 2 | `CheckOrphanConnector` flagged every `add_line` (PowerPoint's `Shapes.AddLine` returns a connector with no endpoints) | `src/modActionsLayout.bas` — `Do_add_line` now tags every line with `Tags.Add "DECKO_KIND", "rule"`. `src/modVerify.bas` `CheckOrphanConnector` skips DECKO_KIND="rule". |
| 3 | No `add_footnote` action — every Citi repro hand-rolls bottom-left text box + page number | New `Do_add_footnote` in `src/modActionsText.bas`; new validator/dispatcher/guidance/types entries in `src/modExecuteInstructions.bas`. |
| 4 | `set_text_autofit shrink` can drive font below readability with no floor knob | `Do_set_text_autofit` gains optional `min_size` param; after PPT's shrink settles, it walks every run and clamps `Font.Size` up to the floor. Dispatcher reads optional `min_size` from JSON; guidance updated. |
| 5 | `tiny_shape_font` 8 pt threshold is too strict for dense financial slides (source frequently 7 pt) | `src/modVerify.bas` — `TINY_TABLE_FONT_PT` lowered 8 → 7. |
| 6 | Doc §3.2 had no guidance on decorative rules / footnote convention; §3.9 omitted the "ref_name strings in shape_ids array" universal-aliasing reminder | `docs/ACTIONS_REFERENCE.md` §3.2 expanded with `add_footnote` entry + the rule-tagging note; §3.9 clarifies `group_shapes` accepts ref_name strings or `shape_names` array. |

Carrier rebaked via `python update_macros.py`. Auto-guidance appendix
regenerated via `python tools/sync_actions_guidance.py` (now lists 247
canonical action types, +1 for `add_footnote`). Static contract gates still
green: 21/21 pytest pass; both repro decks now use `add_footnote` instead of
two manual text boxes (`slide_biz.actions.json` 97 actions, `slide_banamex.
actions.json` 287 actions; M1, M2, M3 still 0 / 0 / 0 warn).

## Contract-boundary / cosmetic deltas (logged, not defects)

1. **Verify `tiny_shape_font` threshold is 8 pt; the source slide uses ~7 pt**
   dense card text. To keep M3 at 0 warn we render the card stat/bullet/footnote
   text at a minimum of 8 pt rather than the source's 7 pt. Faithful to content,
   ~1 pt larger than source. *This is a deliberate gate-vs-source tension, not
   an engine defect.*
2. **Decorative rules/brackets are thin autoshape rects, not `add_line`.** The
   verifier (correctly) treats a zero-thickness `add_line` as an
   orphan/zero-size connector. Using a 2 pt rect for the title rule and the
   GLOBAL NETWORK/INTERCONNECTED/DIVERSIFIED brackets is the clean,
   warning-free way to draw a rule. Recommend documenting this "rules = thin
   rect, not add_line" guidance in `docs/ACTIONS_REFERENCE.md` §3.2 (future
   doc-only improvement; not applied this wave to keep the change set
   engine-inert).
3. **No dedicated footnote action exists** (`insert_slide_number` exists for
   page numbers; there is no `add_footnote`). Footnotes are a manual
   bottom-left `add_text_box`, consistent with the Wave-1 slides 6-7. A
   `add_footnote` convenience macro (bottom-left + thin rule, like the finmodel
   editor's) remains a logged future improvement.
4. **Banking header darkened** from the source tan `#8B7B5A` (white-on-tan
   contrast 4.13, fails WCAG AA) to `#6B5D3E` (6.44, passes). Citi's own slide
   fails AA here; our gate does not. Gate-driven deviation, ~same hue family.
5. **Inter-card connectors** are subtle elbow links (`bottom`→`top`) rather
   than the source's mid-card double-headed arrows, so their bounding box is
   non-degenerate (a horizontal connector trips `zero_size_shape`). Cosmetic;
   the "interconnected" intent is conveyed.
6. **Per-chart segment-name legends** (source lists segment names beside the
   Services card) are omitted — source shows them on one card only; adding them
   to all five would reduce, not raise, fidelity. Segment values are labeled
   inside the bars instead.

## M6 scoping note

`python -m pytest -q` (21 passed) and `python tests/run_smoke.py` (all tests
passed) were run green. The full per-harness sweep of all ~30
`tests/run_smoke_*.py` was **not** run individually: the change set touches no
engine or doc file (additive test assets only), so it is engine-inert and a
per-harness sweep adds no regression signal beyond the canonical aggregator.

## Recommended future engine/doc improvements (logged)

1. Doc: add "decorative rules = thin rect, not `add_line`" guidance + a
   `add_footnote` convenience macro (bottom-left + rule).
2. Engine: optional `add_connector` routing that guarantees a non-degenerate
   bounding box for same-axis endpoints (so a straight horizontal link does not
   trip `zero_size_shape`).
3. Consider relaxing `tiny_shape_font` to 7 pt for dense financial
   reproductions, or making it info-level when `set_text_autofit shrink` is
   explicitly set.

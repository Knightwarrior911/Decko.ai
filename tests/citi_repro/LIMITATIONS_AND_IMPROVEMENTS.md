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

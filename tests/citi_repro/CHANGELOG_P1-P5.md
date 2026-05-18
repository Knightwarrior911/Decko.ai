# Decko engine fixes P1–P5 — CHANGELOG

All implemented in `src/*.bas`, synced to carrier via `update_macros.py`,
guidance appendix regenerated via `tools/sync_actions_guidance.py`. Contract
gate GREEN (246 actions, 0 drift), coverage + doc-sync OK. NOT pushed.

## P1 — table-aware whole-shape font
`src/modActions.bas`: `Do_set_font_size/bold/italic/color` now check
`sh.HasTable` first and route to `modActionsTable.ApplyFontToWholeTable`
(new Public helper that loops every cell via `SetCellTextProp`) instead of
raising `no text frame`. Non-table behavior unchanged.
**Verified:** `set_font_size` on the Slide-1 table → cell(3,1) 8→7 and
cell(10,3)→7, 0 errors. Dense 21-row table sizable without the per-row
workaround.

## P2 — hide_from_legend persists after combo conversion
`src/modActionsChart.bas`: new `HideSeriesFromLegend(ch, idx)` —
`DoEvents` forces a legend layout pass so the entry exists, deletes it,
re-checks the count and retries once (a combo `chart_type` switch re-lays-out
and resurrects a naively-deleted entry). `Do_set_chart_series`
`hide_from_legend` now calls it.
**Verified:** Slide-2 chart = 7 series, legend = 6 entries (Total excluded).

## P3 — suppress_zero_labels on set_chart_series
`src/modActionsChart.bas`: new optional prop `suppress_zero_labels`(bool) —
final pass over `ser.Points`, sets `HasDataLabel=False` where the value is 0
(overrides show_labels/custom_labels).
**Verified:** Goodwill series zero-points [1,2,4,5] → 0 still-labeled.

## P4 — add_chart combo/totals convenience
`src/modActionsChart.bas` `Do_add_chart`: two new optional params
`comboSpec`(array of `{series_index|name, chart_type, axis_group}`) and
`totalsLabel`(bool). Combo retypes/repositions series; totals_label sums every
primary-axis column/bar series, appends a "Total" line series (no line/marker),
labels above with `value_format`, excluded from the legend via
`HideSeriesFromLegend`. `src/modExecuteInstructions.bas`: dispatcher passes
`act("combo")`/`act("totals_label")`; `add_chart` GetActionGuidance NOTE
updated → appendix regenerated.
**Verified:** Slide-2 combo built by ONE add_chart — efficiency series
ChartType=65 (line) AxisGroup=2 (secondary), auto "Total" series present,
columns ChartType=52 primary. Replaced 4 `set_chart_series` + a manual Total
series.

## P5 — real PowerPoint bullets (no fake glyph proxies)
Engine already supported real bullets (`Do_set_bullet_style` →
`ParagraphFormat.Bullet`, `Do_set_indent_level` → `IndentLevel` 0–4); the
defect was in the Citi builders using literal `• – ▪` + leading spaces.
`tests/citi_repro/build_slide1.py` (highlights) and `build_slide2.py` (driver
cards) rewritten: plain-text paragraphs + per-paragraph `set_indent_level` +
`set_bullet_style` (disc/dash/square by level) + level color/size.
**Verified (structural, not pixels):** Slide-1 highlights 8 paragraphs,
`Bullet.Type=1`, `IndentLevel` 1/2/3, first chars are real words (R,N,N,N,N,
E,P,R) — no glyph chars. Slide-2 card paragraphs likewise `Bullet.Type=1`.

## Combined regression
`citi_final.actions.json` (569 actions) → **0 non-ok**, slide 1 + slide 2 each
= 1 real ChartObject + 1 real table, **0 autoshape-as-chart**. Side-by-side
PNGs (`SBS_slide1.png`, `SBS_slide2.png`) confirm fidelity retained with real
bullets and the single-action combo chart.

## Docs
- `ACTIONS_REFERENCE.md` §3.11: `suppress_zero_labels` added to set_chart_series
  data-labels; combo recipe + `visible:false`-collapses-columns gotcha (prior).
- `ACTIONS_REFERENCE.md` §3.10: cell/row/column font subsection (prior).
- `EXAMPLES.md` §17.9: one-action combo + totals + suppress-zero example.
- Appendix auto-regenerated; contract/coverage/doc-sync gates GREEN.

## Files changed (staged, NOT pushed)
`src/modActions.bas`, `src/modActionsChart.bas`, `src/modActionsTable.bas`,
`src/modExecuteInstructions.bas`, `PPT_AI_Editor.pptm` (rebaked),
`docs/ACTIONS_REFERENCE.md`, `docs/EXAMPLES.md`, and `tests/citi_repro/*`.

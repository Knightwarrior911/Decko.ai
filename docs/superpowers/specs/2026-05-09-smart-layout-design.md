# Smart Layout + Alignment — Design Spec

**Status:** SHIPPED (2026-05-09). 11 actions live; 12 smoke tests green (9 verified + 3 fixture-limited skips). Carrier action count: 70 → 81.
**Owner:** Decko.ai backend.
**Touches:** `src/modActionsLayout.bas`, `src/modExecuteInstructions.bas`, `tests/run_smoke_layout.py` (new), `README.md`.

## Goal

Close the "smart layout" gap so a VP+ user can express any reasonable spatial change without nudging shapes by hand: snap to a grid, fit a box to its content, equalize spacing, match a reference shape, swap two shapes, auto-group what visually overlaps.

Current carrier has the basic six (`align_shapes`, `distribute_horizontal`, `distribute_vertical`, `tile_grid`, `fit_to_slide_margins`, `move_shape_relative`) — they cover bulk arrangement but not "I want this one fixed relative to that one" or "snap everything to clean coordinates."

This spec adds 11 actions. Carrier action count grows 70 → 81.

## Scope

In:
- 11 new actions in `modActionsLayout.bas`.
- 11 new validation arms + 11 dispatch arms in `modExecuteInstructions.bas`.
- New smoke harness `tests/run_smoke_layout.py` mirroring the `run_smoke_text.py` shared-app pattern.

Out (deferred):
- Smart guides / connector reroute on shape moves.
- Z-order ops (existing? unclear; out of scope here regardless).
- Multi-slide layout cascades.
- Animation timing.

No snapshot schema changes — all actions target existing shapes by id.

## Actions

All actions resolve targets via existing `modActions.FindShape(slide, id)`.

| Action | Args | Effect |
|---|---|---|
| `snap_to_grid` | `slide, shape_id, grid_pt: float` | Round shape's `Left` and `Top` to nearest multiple of `grid_pt`. |
| `align_to_slide_center` | `slide, shape_id, axis: "h" \| "v" \| "both"` | Center shape on slide. `h` aligns horizontally, `v` vertically, `both` does both. |
| `nudge` | `slide, shape_id, direction: "l" \| "r" \| "u" \| "d", amount_pt: float` | Shift by amount in given direction. Sugar over `move_shape_relative` for LLM ergonomics. |
| `fit_to_content` | `slide, shape_id` | Auto-resize shape to its text bounding box, then re-lock auto-size to none. |
| `match_size` | `slide, ref_shape_id, target_shape_ids: [int]` | Copy reference shape's `Width` and `Height` to each target. |
| `uniform_size` | `slide, shape_ids: [int], width_pt: float, height_pt: float` | Set all shapes to identical width and height. |
| `smart_spacing` | `slide, shape_ids: [int], gap_pt: float, axis: "h" \| "v"` | After sorting along axis, place each subsequent shape exactly `gap_pt` away from the previous one's far edge. |
| `equalize_spacing` | `slide, shape_ids: [int], axis: "h" \| "v"` | Detect existing total span, divide by gap count, write equal gaps. |
| `match_position` | `slide, ref_shape_id, target_shape_id, edge: "left" \| "right" \| "top" \| "bottom" \| "hcenter" \| "vcenter"` | Align target's edge to reference's same edge. |
| `swap_positions` | `slide, shape_a_id, shape_b_id` | Swap `Left`/`Top`/`Width`/`Height` of two shapes. |
| `group_by_overlap` | `slide, shape_ids: [int]` | Group only the subset of input shapes whose bounding boxes intersect at least one other in the set. Non-overlapping shapes left alone. If <2 overlaps, no-op (logged). |

## Validation

Each action must pass:
- `slide_num` in range.
- All `shape_id`s resolve to shapes on that slide.
- Numeric fields (`grid_pt`, `amount_pt`, `gap_pt`, `width_pt`, `height_pt`) are positive numbers.
- Enum fields (`axis`, `direction`, `edge`) are one of the documented values.
- For multi-shape actions (`match_size`, `uniform_size`, `smart_spacing`, `equalize_spacing`, `group_by_overlap`), `shape_ids` is a non-empty array.
- `equalize_spacing` requires `shape_ids.Count >= 3` (need at least 2 gaps to equalize).

Validation failure → row skipped, logged, batch continues.

## Implementation notes

- `snap_to_grid`: `sh.Left = Round(sh.Left / grid) * grid`; same for Top. Negative or zero `grid_pt` rejected at validation.
- `align_to_slide_center`: read slide width/height from `pres.PageSetup.SlideWidth/Height`; compute target so shape center matches slide center.
- `nudge`: dx/dy from direction enum, then `sh.Left += dx`, `sh.Top += dy`.
- `fit_to_content`: temporarily set `sh.TextFrame.AutoSize = ppAutoSizeShapeToFitText`, force a layout via reading any size, then set back to `ppAutoSizeNone`. If shape has no text frame, raise.
- `match_size`: trivial — loop targets.
- `uniform_size`: trivial — loop and assign.
- `smart_spacing`: sort by `Left` (axis="h") or `Top` (axis="v"); for i = 1..N-1, set the i-th shape's start = prev.start + prev.size + gap.
- `equalize_spacing`: sort, compute first.start, last.start+size = total span; equal_gap = (span - sum_widths) / (N-1); assign.
- `match_position`: read ref edge, write target's same coordinate. For `hcenter`/`vcenter`, account for half-width/half-height.
- `swap_positions`: simple temp swap of `Left`, `Top`, `Width`, `Height`.
- `group_by_overlap`: build pairwise overlap matrix; keep shapes that have ≥1 overlap with any other in the input set; if ≥2 such shapes, call `Slide.Shapes.Range(...).Group`. Else log "no overlaps to group" and exit cleanly (counted as applied = false; skipped + reason logged).

## Tests

`tests/run_smoke_layout.py` (new) — single shared PowerPoint instance pattern. 11 tests, one per action. Reuses `test_decks/phase2.pptx` (already has multiple shapes for layout exercise) for most; uses `text_v3.pptx` for `fit_to_content`.

Each test:
1. Reload deck.
2. Snapshot → grab shape ids.
3. Run action via `app.Run("PPT_AI_Editor!Do_<name>", ...)`.
4. Re-snapshot.
5. Assert geometry.

Plus 1 cross-action test: `snap_to_grid` then `equalize_spacing` chained — confirms sequential shape mutations don't interfere.

Total: 12 new assertions on top of current text-smoke 16 and main-smoke ~38.

## Migration

- Snapshot v3 unchanged — no schema bump.
- `modActionsLayout.bas` grows from 13 to 24 actions.
- `README.md` action table grows 70 → 81.
- No new modules. No UserForm changes.

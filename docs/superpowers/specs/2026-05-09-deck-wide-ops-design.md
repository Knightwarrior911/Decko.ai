# Deck-Wide Operations — Design Spec

**Status:** SHIPPED (2026-05-09). 9 actions live; 8 smoke tests green (1 skipped without logo.png fixture). Carrier action count: 81 → 90.
**Owner:** Decko.ai backend.
**Touches:** new `src/modActionsDeck.bas`, `src/modExecuteInstructions.bas`, new `tests/run_smoke_deck.py`, `README.md`.

## Goal

Close the "I want this change everywhere" gap. Today's actions edit one slide or one shape at a time; today's `find_replace_text` is literal-only and per-shape. A VP+ user expects to swap a font, recolor a palette, apply a theme, change aspect ratio, or paste a logo on every slide in one shot.

This spec adds 9 actions in a new `modActionsDeck.bas`. Carrier action count grows 81 → 90.

## Scope

In:
- 9 new actions in a new module `modActionsDeck.bas`.
- 9 new validation arms + 9 dispatch arms in `modExecuteInstructions.bas`.
- New smoke harness `tests/run_smoke_deck.py` reusing the shared-PowerPoint pattern.

Out (deferred):
- Slide master / custom layout authoring (creating new layouts, not just applying).
- Animation timing rewrites.
- Comment / annotation propagation.
- Header/footer/page-numbers — moved to a future spec.

## Actions

| Action | Args | Effect |
|---|---|---|
| `find_replace_regex` | `scope: "deck" \| "slide:N", pattern: string, replacement: string` | Regex find/replace across text in scope. Uses VBScript.RegExp late-bound. Case-sensitive by default. Replaces *all* matches per shape. |
| `swap_font_deck_wide` | `from_name: string, to_name: string` | Walk every slide → shape → text run; if `Font.Name == from_name`, set to `to_name`. |
| `recolor_palette_deck_wide` | `from_hex: "#RRGGBB", to_hex: "#RRGGBB", target: "fill" \| "font" \| "both"` | Replace `from_hex` with `to_hex` across deck. `fill` = shape solid fills; `font` = run font colors; `both` = both. |
| `apply_theme` | `thmx_path: string` | Apply theme via `Presentation.ApplyTemplate(thmx_path)`. Path must be absolute and exist. Accepts `.thmx` or `.potx`. |
| `set_slide_size` | `width_pt: float, height_pt: float` OR `preset: "16:9" \| "4:3"` | Set `pres.PageSetup.SlideWidth/Height`. Preset values: 16:9 = 960×540, 4:3 = 720×540. Either explicit dims OR preset, not both. |
| `set_theme_font` | `major: string, minor: string` | Write both `Theme.ThemeFontScheme.MajorFont.Latin.Name` and `MinorFont.Latin.Name` on the slide master. Empty string = leave that slot unchanged. |
| `bulk_insert_image` | `slide_indices: [int], picture_path: string, left: float, top: float, width: float, height: float` | Insert the same image at the same position on each listed slide. Slides outside range are skipped and logged. |
| `bulk_insert_text_box` | `slide_indices: [int], text: string, left, top, width, height: float` | Same shape: AddTextbox on each listed slide with the same text + dims. |
| `apply_layout_to_slides` | `slide_indices: [int], layout_index: int` | For each listed slide, set `slide.CustomLayout = master.CustomLayouts(layout_index + 1)`. 0-indexed input, 1-indexed PowerPoint API. |

## Validation

Common:
- `slide_indices` (when present) must be a non-empty array of positive integers.
- `from_hex` / `to_hex` not pre-validated (HexToRgb will raise on bad input → caught by dispatcher).

Per-action:
- `find_replace_regex`: scope must match `deck` or `slide:N` (N positive); `pattern` non-empty.
- `swap_font_deck_wide`: both `from_name` and `to_name` non-empty strings.
- `recolor_palette_deck_wide`: `target` ∈ {`fill`, `font`, `both`}.
- `apply_theme`: `thmx_path` non-empty; existence check is at runtime (raises if missing).
- `set_slide_size`: either both `width_pt`+`height_pt` (positive numbers) OR a `preset` string. If both, error "specify dims OR preset, not both."
- `set_theme_font`: at least one of `major`/`minor` non-empty.
- `bulk_insert_image`/`bulk_insert_text_box`: dimension floats present and ≥ 0.
- `apply_layout_to_slides`: `layout_index` ≥ 0.

Validation failure → row skipped, batch continues.

## Implementation notes

### `find_replace_regex`

Late-bound `CreateObject("VBScript.RegExp")` with `.Global = True`. Per shape with text frame, get `tf.TextRange.Text`, run `re.Replace(text, replacement)`, write back if changed. Counters track matches replaced.

### `swap_font_deck_wide`

Walk: `For each Slide → For each Shape → If HasTextFrame → For each Paragraph → For each Run → If Run.Font.Name = from_name Then Run.Font.Name = to_name`.

### `recolor_palette_deck_wide`

For `fill` target: walk shapes; if `sh.Fill.Type = msoFillSolid` and color matches `from_hex`, set to `to_hex`.

For `font` target: walk shapes → text runs; if `Run.Font.Color.RGB` matches, set to new RGB.

For `both`: do both passes.

Color comparison via `RgbToHex(actual) = from_hex`.

### `apply_theme`

```vb
ActivePresentation.ApplyTemplate themePath
```

Wrapped in `On Error` to convert "file not found" / "invalid theme" to clear errors.

### `set_slide_size`

```vb
If preset = "16:9" Then
    PageSetup.SlideWidth = 960
    PageSetup.SlideHeight = 540
ElseIf preset = "4:3" Then
    PageSetup.SlideWidth = 720
    PageSetup.SlideHeight = 540
Else
    PageSetup.SlideWidth = widthPt
    PageSetup.SlideHeight = heightPt
End If
```

### `set_theme_font`

```vb
With ActivePresentation.SlideMaster.Theme.ThemeFontScheme
    If Len(majorName) > 0 Then .MajorFont.Latin.Name = majorName
    If Len(minorName) > 0 Then .MinorFont.Latin.Name = minorName
End With
```

### `bulk_insert_image` / `bulk_insert_text_box`

Iterate `slide_indices`. For each that is in range, call `Slides(i).Shapes.AddPicture(path, msoFalse, msoTrue, left, top, width, height)` or `AddTextbox(msoTextOrientationHorizontal, ...).TextFrame.TextRange.Text = text`. Out-of-range skipped (logged).

### `apply_layout_to_slides`

```vb
Dim master As Master: Set master = ActivePresentation.SlideMaster
For Each idx In indices
    Slides(idx).CustomLayout = master.CustomLayouts(layoutIndex + 1)
Next
```

## Tests

`tests/run_smoke_deck.py` — single shared PowerPoint instance pattern. ~9 tests, one per action. Reuses `phase2.pptx` and `text_v3.pptx` fixtures.

Each test:
1. Open carrier + fresh deck copy.
2. Run action via `app.Run("PPT_AI_Editor!Do_<name>", ...)`.
3. Re-snapshot or COM-inspect.
4. Assert state matches expectation.

`apply_theme` skipped if no `.thmx` available — try `set_theme_font` instead as a comparable theme-touching test.

## Migration

- No snapshot schema changes — all actions either don't emit fields or modify existing ones (text/color/font name) that snapshot already captures.
- New module `modActionsDeck.bas` auto-imported by `update_macros.py`.
- README action table grows 81 → 90; new module entry added.
- No UserForm changes.

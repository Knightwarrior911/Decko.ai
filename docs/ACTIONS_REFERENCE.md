# Decko.ai — Complete Action Reference

This is the **machine-precise** spec of every action Decko can execute. It is
written for AI assistants (Hermes, OpenClaw, GPT-class, Claude-class — any model)
so that the assistant can (a) help a VP write a request and (b) emit the exact
JSON `actions` array a model should return.

**If you only read one thing, read [§0 Hard Rules](#0-hard-rules).** Breaking
any of them makes the whole batch fail.

> **Counts:** ~130 action types across 14 VBA modules. The authoritative source
> of truth is `src/modExecuteInstructions.bas` (`ValidateAction` = required
> fields, `DispatchAction` = how each field is read). This document mirrors it.

---

## 0. Hard Rules

1. **The output is one JSON object: `{"actions": [ ... ]}`.** A bare array
   `[ ... ]` is **invalid** and will crash with runtime error 438. Always wrap.
2. **Each item in `actions` is an object with a `"type"` field** naming the
   action, plus that action's fields. Actions run **top to bottom, in order**.
3. **`slide` is the 1-based slide number** (slide 1 = first slide).
4. **`shape_id` is the numeric ID shown in the snapshot** for that shape — *not*
   the shape's name and *not* its position in a list. Exception: a shape you
   create earlier in the *same* batch (via `ref_name`) can be referenced by that
   `ref_name` string anywhere a `shape_id` is expected.
5. **All distances/sizes are in points (pt).** 1 inch = 72 pt. A 16:9 slide is
   **960 × 540 pt** (PowerPoint's "Widescreen"); a 4:3 slide is 720 × 540 pt.
   The snapshot reports the actual slide size — use it.
6. **Colors are `"#RRGGBB"` strings** (uppercase or lowercase hex, leading `#`).
7. **Booleans are JSON `true` / `false`.** (Strings `"yes"`/`"no"`/`"on"`/`"off"`/`"1"`/`"0"` are also accepted, but prefer real booleans.)
8. **`pos` is an object: `{"left": N, "top": N, "width": N, "height": N}`** —
   used by `add_shape`, `add_text_box`, `add_chart`, `add_table`,
   `insert_picture`, `insert_slide_number`.
9. **Scope strings are `"deck"` or `"slide:N"`** (e.g. `"slide:3"`) for the
   deck-wide find/replace and recolor actions.
10. **You must be given a snapshot before you can reference existing shapes.**
    The snapshot lists every slide, every shape, its numeric `shape_id`, kind,
    position, and current text. Without it you cannot write correct actions for
    existing content (you *can* still add brand-new content blind, but you
    should ask for the snapshot first whenever the request touches what's
    already there).
11. **Decko never invents data.** If a request needs facts (financials, names,
    dates), they must come from the VP or from public sources the VP approves —
    not from the model's imagination.
12. **Unknown action types, missing required fields, or out-of-range
    slide/shape IDs cause that single action to be `skipped` (logged, batch
    continues).** A malformed JSON document fails the whole batch. So: when in
    doubt, omit the action rather than guess.
13. **Large batches: don't paste into the Execute textbox — use the form's
    "Load from file..." button.** MSForms textboxes corrupt big/long-line
    pastes (they inject whitespace into numbers and keys). If you must paste,
    emit **one action per line** so no single line is more than ~1 KB.

---

## 1. The snapshot (input you receive)

A snapshot is a plain-text dump the VP exports from Decko (Alt+F8 →
`ExportSnapshot`). It contains, per slide:

- Slide number, layout name/index, slide size.
- Every shape: `shape_id` (number), `name`, `kind` (textbox / autoshape /
  table / chart / picture / group / placeholder…), bounding box (left/top/
  width/height in pt), and the current text broken into **paragraphs** and,
  within mixed-format paragraphs, **runs** — each paragraph has a
  `paragraph_index` (0-based) and each run a `run_index` (0-based).
- For tables: rows × cols and cell text.
- For charts: chart type, series names, categories.
- Speaker notes per slide.

**Indices you will reference:**
- `slide` — 1-based slide number.
- `shape_id` — the number in the snapshot.
- `paragraph_index` — 0-based paragraph within a shape's text frame.
- `run_index` — 0-based run within a paragraph (only needed for `set_run_*`).
- table `row` / `col` — 1-based.
- chart `series_index` — 1-based.

---

## 2. Universal value vocabularies

| Concept | Allowed values |
|---|---|
| Color | `"#RRGGBB"` |
| Boolean | `true` / `false` |
| Scope | `"deck"` or `"slide:N"` |
| Horizontal align (`h_align`, paragraph `value`) | `left` `center` `right` `justify` |
| Vertical align (`v_align`, `set_text_vertical_align`) | `top` `middle` `bottom` |
| Bullet style (`set_bullet_style`) | `none` `disc` (a.k.a. `bullet`) `square` `dash` `number` `letter` |
| Autofit mode (`set_text_autofit`) | `none` `shrink` `resize` |
| Line/dash style | `solid` `dash` `dot` `round_dot` `dash_dot` (also `dashdot`) |
| Arrowhead (`arrow_end`/`arrow_start`) | `none` `filled` (`triangle`) `open` `stealth` `diamond` `oval` |
| Arrow size (`arrow_size`) | `small` `medium` `large` |
| Connector kind (`add_connector` `kind`) | `straight` `elbow` `curved` |
| Connector anchor point (`from_point`/`to_point`) | `auto` `top` `bottom` `left` `right` |
| Z-order (`z_order` `order`) | `front` `back` `forward` `backward` |
| Recolor target (`recolor_palette_deck_wide` `target`) | `fill` `font` `both` |
| Picture recolor (`recolor_picture` `color_type`) | `grayscale` `sepia` `washout` `bw` `auto` |
| 3D bevel (`set_3d_bevel` `type`) | `circle` `slope` `cross` `angle` `softround` |
| Nudge direction (`nudge` `direction`) | `l` `r` `u` `d` |
| Slide-size preset (`set_slide_size` `preset`) | `16:9` `4:3` |
| Axis (`set_chart_axis_title` `axis`, `set_chart_axis` `axis`) | `category` (`x`) `value` (`y`) |
| Legend position | `top` `right` `bottom` `left` `none` |

### Layout indices (`layout_index` for `add_slide`, `apply_layout_to_slides`)

0-based index into the deck's slide-master custom layouts. In a default Office
theme:

| Index | Layout |
|---|---|
| 0 | Title Slide |
| 1 | Title and Content |
| 2 | Section Header |
| 3 | Two Content |
| 4 | Comparison |
| 5 | Title Only |
| 6 | Blank |
| 7 | Content with Caption |
| 8 | Picture with Caption |
| 9, 10, … | theme-specific extras |

If the deck uses a custom template, the snapshot reports the actual layout
names — prefer those. `6` (Blank) is the safe default for "add a slide I'll
fill with shapes".

### `add_shape` / `set_shape_kind` shape kinds

Pass the `kind` string:

`rect`/`rectangle`, `rrect`/`round_rect`, `capsule`, `oval`/`ellipse`/`circle`,
`parallelogram`, `trapezoid`, `diamond`, `octagon`, `hexagon`, `pentagon`,
`triangle`/`isosceles`, `right_triangle`, `cross`/`plus`,
`arrow`/`right_arrow`, `left_arrow`, `up_arrow`, `down_arrow`,
`double_arrow`/`left_right_arrow`, `up_down_arrow`, `quad_arrow`,
`curved_right_arrow`, `curved_left_arrow`, `curved_up_arrow`,
`curved_down_arrow`, `striped_arrow`, `notched_arrow`, `bent_arrow`,
`u_turn_arrow`, `chevron`, `chevron_pentagon`,
`callout_rect`/`rectangular_callout`, `callout_rrect`/`rounded_callout`,
`callout_oval`, `callout_cloud`, `callout_line1`, `callout_line2`,
`star4`, `star5`/`star`, `star8`, `star10`, `star12`, `star16`, `star24`, `star32`,
`ribbon_up`, `ribbon_down`, `donut`/`ring`, `block_arc`, `arc`,
`brace_left`, `brace_right`, `bracket_left`, `bracket_right`, `plaque`,
`no_symbol`, `cloud`.

### Chart types (`add_chart` `chart_type`, `set_chart_type` `value`)

All of these are supported and create a **real native PowerPoint chart object**
(not an image):

**2-D column/bar:** `columnclustered`, `columnstacked`, `columnstackedpercent`
(a.k.a. `column_100pct`), `barclustered`, `barstacked`, `barstackedpercent`
(`bar_100pct`)
**3-D column/bar:** `column3d`, `columnclustered3d`, `columnstacked3d`,
`bar3d`, `barclustered3d`, `barstacked3d`
**Line:** `line`, `linemarkers`, `linestacked`, `linestackedmarkers`, `line3d`
**Area:** `area`, `areastacked`, `areapercent` (`area_100pct`), `area3d`, `areastacked3d`
**Pie/doughnut:** `pie`, `pie3d`, `pieexploded3d`, `doughnut`
**XY:** `scatter`
**Radar:** `radar`, `radarmarkers`, `radarfilled`
**Surface:** `surface`, `surfacewireframe`
**Modern (Office 2016+):** `waterfall`, `pareto`, `funnel`, `histogram`, `boxwhisker`, `treemap`, `sunburst`

> **Limitation for the 7 "modern" types** (`waterfall`, `pareto`, `funnel`,
> `histogram`, `boxwhisker`, `treemap`, `sunburst`): PowerPoint's automation
> surface for these is broken — Decko creates the chart shape with the correct
> type and PowerPoint's **default placeholder data**, but cannot write your
> `categories`/`series`/`title` into them. The VP edits the data manually after
> insertion. For all other 32 types, `categories`, `series`, and `title` are
> applied normally.

---

## 3. Action catalogue

Format: **`action_name`** — what it does.
`req:` required fields (type). `opt:` optional fields (type) = default.
`ex:` a minimal example object.

### 3.1 Shape text & basic formatting (whole shape)

- **`set_text`** — replace ALL text in a shape with one plain string (collapses
  to a single run; loses mixed formatting — use `set_run_text` to preserve it).
  `req:` `slide`(int), `shape_id`(int|ref), `value`(string).
  `ex:` `{"type":"set_text","slide":2,"shape_id":5,"value":"Q3 Revenue"}`
- **`set_font_size`** — font size (pt) on the whole text frame.
  `req:` `slide`, `shape_id`, `value`(int).
- **`set_font_bold`** / **`set_font_italic`** — toggle on whole shape.
  `req:` `slide`, `shape_id`, `value`(bool).
- **`set_font_color`** — recolor all text in a shape. `req:` `slide`, `shape_id`, `value`(`#RRGGBB`).
- **`set_fill_color`** — solid background fill. `req:` `slide`, `shape_id`, `value`(`#RRGGBB`).
- **`move_shape`** — absolute position. `req:` `slide`, `shape_id`, `left`(num), `top`(num).
- **`resize_shape`** — `req:` `slide`, `shape_id`, `width`(num), `height`(num).
- **`delete_shape`** — `req:` `slide`, `shape_id`.
- **`duplicate_shape`** — clone a shape at a new position.
  `req:` `slide`, `shape_id`, `left`(num), `top`(num). `opt:` `ref_name`(string).
- **`rotate_shape`** — `req:` `slide`, `shape_id`, `degrees`(num).
- **`flip_shape`** — `req:` `slide`, `shape_id`, `axis`(`h`|`v`).
- **`set_shape_adjustment`** — drag a shape's yellow adjustment handle.
  `req:` `slide`, `shape_id`, `index`(int, 0-based handle), `value`(num, 0.0–1.0-ish).
- **`z_order`** — `req:` `slide`, `shape_id`, `order`(`front`|`back`|`forward`|`backward`).
- **`copy_formatting`** — copy fill/line/font/effects from one shape to another.
  `req:` `slide`, `source_shape_id`, `target_shape_id`.

### 3.2 Add new shapes / text boxes / lines

- **`add_shape`** — new autoshape.
  `req:` `slide`, `kind`(string, see vocab), `pos`({left,top,width,height}).
  `opt:` `fill`(`#RRGGBB`|null), `stroke`(`#RRGGBB`|null), `stroke_weight_pt`(num)=1.0,
  `ref_name`(string), `text`(string), `font_color`(`#RRGGBB`), `font_size`(int),
  `font_bold`(bool)=false, `h_align`(string)=`center`, `v_align`(string)=`middle`,
  `super_suffix`(string), `sub_suffix`(string).
  `ex:` `{"type":"add_shape","slide":3,"kind":"rrect","pos":{"left":60,"top":120,"width":200,"height":80},"fill":"#15283C","text":"Phase 1","font_color":"#FFFFFF","font_size":18,"ref_name":"box_p1"}`
- **`add_text_box`** — plain text box (no fill/stroke by default).
  `req:` `slide`, `text`(string), `pos`.
  `opt:` `ref_name`, `font_color`, `font_size`(int), `font_bold`(bool)=false,
  `font_italic`(bool)=false, `h_align`, `fill`(`#RRGGBB`|null), `stroke`(`#RRGGBB`|null),
  `stroke_weight_pt`(num)=1.0, `super_suffix`, `sub_suffix`.
- **`add_line`** — straight line/divider between two points.
  `req:` `slide`, `x1`,`y1`,`x2`,`y2`(num), `color`(`#RRGGBB`), `weight_pt`(num).
  `opt:` `arrow_end`(string)=`none`, `arrow_start`(string)=`none`, `dash_style`(string)=`solid`.
- **`set_shape_kind`** — morph an existing shape to a different `kind` (keeps pos/text).
  `req:` `slide`, `shape_id`, `kind`(string).

### 3.3 Paragraph-level text

All take `slide`, `shape_id`, `paragraph_index`(0-based).
- **`set_paragraph_text`** — replace one paragraph's text (single-run; loses mixed formatting). `req:` … `value`(string).
- **`add_paragraph`** — insert a paragraph. `req:` `slide`, `shape_id`, `after_paragraph_index`(int, use `-1` to prepend), `value`(string).
- **`delete_paragraph`** — `req:` `slide`, `shape_id`, `paragraph_index`.
- **`set_bullet_style`** — `req:` … `value`(`none`|`disc`|`square`|`dash`|`number`|`letter`).
- **`set_indent_level`** — `req:` … `value`(int 0–4).
- **`set_paragraph_font_size`** — `req:` … `value`(int).
- **`set_paragraph_font_color`** — `req:` … `value`(`#RRGGBB`).
- **`set_paragraph_alignment`** — `req:` … `value`(`left`|`center`|`right`|`justify`).
- **`set_paragraph_line_spacing`** — `req:` … `value`(num, multiple e.g. 1.0, 1.5).

### 3.4 Run-level formatting (sub-paragraph precision)

All take `slide`, `shape_id`, `paragraph_index`(0-based), `run_index`(0-based).
Use these when a paragraph mixes formats (e.g. **bold drug name** + plain
description) — they touch only the named run.
- **`set_run_text`** — `req:` … `value`(string).
- **`set_run_bold`** / **`set_run_italic`** / **`set_run_underline`** / **`set_run_subscript`** / **`set_run_superscript`** / **`set_run_strikethrough`** — `req:` … `value`(bool).
- **`set_run_font_color`** — `req:` … `value`(`#RRGGBB`).
- **`set_run_font_size`** — `req:` … `value`(int).
- **`set_run_font_name`** — `req:` … `value`(string font name).
- **`set_run_hyperlink`** — `req:` … `value`(string URL; empty string `""` clears the link).

### 3.5 Text-frame behaviour

- **`set_text_vertical_align`** — `req:` `slide`, `shape_id`, `value`(`top`|`middle`|`bottom`).
- **`set_text_autofit`** — `req:` `slide`, `shape_id`, `mode`(`none`|`shrink`|`resize`).
- **`set_text_margin`** — internal padding. `req:` `slide`, `shape_id`, `left`,`right`,`top`,`bottom`(num pt).
- **`enable_text_shrink_for_overflow`** — sweep a scope and turn on shrink-on-overflow for every text frame.
  `req:` `scope`(`deck`|`slide:N`). `opt:` `include_titles`(bool)=false.
- **`fit_to_content`** — auto-resize a shape to fit its text. `req:` `slide`, `shape_id`.

### 3.6 Find / replace

- **`find_replace_text`** — literal, safe find/replace.
  `req:` `scope`(`deck`|`slide:N`), `find`(string), `replace`(string).
- **`find_replace_regex`** — regex find/replace.
  `req:` `scope`, `pattern`(regex string), `replacement`(string).

### 3.7 Layout, alignment, distribution

Several take `shape_ids` — a JSON **array of shape_id numbers** (a `ref_name`
string is also accepted per element).
- **`align_shapes`** — `req:` `slide`, `shape_ids`(array), `anchor`(`left`|`right`|`top`|`bottom`|`hcenter`|`vcenter`).
- **`distribute_horizontal`** / **`distribute_vertical`** — even gaps. `req:` `slide`, `shape_ids`(array).
- **`tile_grid`** — N-column grid. `req:` `slide`, `shape_ids`(array), `cols`(int), `gap_pt`(num).
- **`smart_spacing`** — sort by axis, place each `gap_pt` from previous edge. `req:` `slide`, `shape_ids`(array), `gap_pt`(num), `axis`(`h`|`v`).
- **`equalize_spacing`** — equal gaps along axis. `req:` `slide`, `shape_ids`(array), `axis`(`h`|`v`).
- **`uniform_size`** — set all to same size. `req:` `slide`, `shape_ids`(array), `width_pt`(num), `height_pt`(num).
- **`match_size`** — copy one shape's size to others. `req:` `slide`, `ref_shape_id`, `target_shape_ids`(array).
- **`match_position`** — align target edge to reference edge. `req:` `slide`, `ref_shape_id`, `target_shape_id`, `edge`(`left`|`right`|`top`|`bottom`|`hcenter`|`vcenter`).
- **`swap_positions`** — swap two shapes' pos+size. `req:` `slide`, `shape_a_id`, `shape_b_id`.
- **`group_by_overlap`** — group shapes whose boxes intersect. `req:` `slide`, `shape_ids`(array).
- **`fit_to_slide_margins`** — shrink one shape to fit inside slide minus margin. `req:` `slide`, `shape_id`. `opt:` `margin_pt`(num)=36.
- **`move_shape_relative`** — `req:` `slide`, `shape_id`, `dx_pt`(num), `dy_pt`(num).
- **`nudge`** — `req:` `slide`, `shape_id`, `direction`(`l`|`r`|`u`|`d`), `amount_pt`(num).
- **`snap_to_grid`** — round pos to a grid. `req:` `slide`, `shape_id`, `grid_pt`(num).
- **`align_to_slide_center`** — `req:` `slide`, `shape_id`, `axis`(`h`|`v`|`both`).
- **`clear_slide`** — delete all shapes on a slide. `req:` `slide`. `opt:` `keep_shape_ids`(array) — IDs/ref_names to spare.

### 3.8 Match / batch recolor / delete by filter

- **`recolor_fill_match`** — `req:` `scope`, `from`(`#RRGGBB`), `to`(`#RRGGBB`).
- **`recolor_font_match`** — `req:` `scope`, `from`, `to`.
- **`delete_shapes_match`** — delete shapes matching filters. `req:` `scope`. `opt:` (at least one) `kind`(string), `fill`(`#RRGGBB`), `text_contains`(string).

### 3.9 Connectors & groups

- **`add_connector`** — line/elbow/curved connector between two shapes.
  `req:` `slide`, `kind`(`straight`|`elbow`|`curved`), and a from/to pair: either
  `from_shape_id` & `to_shape_id` (numbers), or `from_shape_name` & `to_shape_name`
  (or `ref_name`s of shapes created in this batch).
  `opt:` `arrow_end`(string)=`filled`, `arrow_start`(string)=`none`, `arrow_size`(`small`|`medium`|`large`)=`medium`,
  `color`(`#RRGGBB`)=`#000000`, `weight_pt`(num)=1.0, `from_point`/`to_point`(`auto`|`top`|`bottom`|`left`|`right`)=`auto`, `dash_style`(string)=`solid`.
  `ex:` `{"type":"add_connector","slide":3,"kind":"elbow","from_shape_name":"box_p1","to_shape_name":"box_p2","arrow_end":"filled"}`
- **`group_shapes`** — `req:` `slide`, `shape_ids`(array). `opt:` `ref_name`(string).
- **`ungroup`** — `req:` `slide`, `shape_id`(the group).

### 3.10 Tables

- **`add_table`** — new table. `req:` `slide`, `rows`(int), `cols`(int), `pos`. `opt:` `ref_name`(string).
- **`set_cell_text`** — `req:` `slide`, `shape_id`(the table), `row`(1-based), `col`(1-based), `value`(string).
- **`add_table_row`** — `req:` `slide`, `shape_id`, `after_row`(int, 0 = before first).
- **`delete_table_row`** — `req:` `slide`, `shape_id`, `row`(int).
- **`add_table_col`** — `req:` `slide`, `shape_id`, `after_col`(int).
- **`delete_table_col`** — `req:` `slide`, `shape_id`, `col`(int).
- **`swap_table_columns`** — `req:` `slide`, `shape_id`, `col_a`(int), `col_b`(int).
- **`swap_table_rows`** — `req:` `slide`, `shape_id`, `row_a`(int), `row_b`(int).
- **`merge_cells`** — `req:` `slide`, `shape_id`, `row_a`,`col_a`,`row_b`,`col_b`(int).
- **`set_table_col_width`** — `req:` `slide`, `shape_id`, `col`(int), `width_pt`(num).
- **`set_table_row_height`** — `req:` `slide`, `shape_id`, `row`(int), `height_pt`(num).
- **`set_cell_border`** — `req:` `slide`, `shape_id`, `row`, `col`, `side`(`left`|`right`|`top`|`bottom`). `opt:` `color`(`#RRGGBB`), `weight_pt`(num), `visible`(bool)=true.
- **`set_cell_text_align`** — `req:` `slide`, `shape_id`, `row`, `col`. `opt:` `h_align`(string), `v_align`(string) — at least one.
- **`set_cell_fill`** — `req:` `slide`, `shape_id`, `row`, `col`, `color`(`#RRGGBB`).
- **`apply_table_style`** — `req:` `slide`, `shape_id`, `style_id`(string — an Office table style GUID or name; common: `"NoStyleNoGrid"`, `"MediumStyle2Accent1"`).
- **`build_image_grid_table`** — build a 2-column image+caption table from a row spec (see §3.12).

### 3.11 Charts

> See [§2 chart types](#chart-types-add_chart-chart_type-set_chart_type-value) for the full type list and the modern-type limitation.

- **`add_chart`** — insert a new native chart.
  `req:` `slide`, `chart_type`(string, see list), `pos`({left,top,width,height}),
  `categories`(array of strings — the x-axis labels),
  `series`(array of `{ "name": string, "values": [numbers], "color"?: "#RRGGBB" }`).
  `opt:` `ref_name`(string), `title`(string), `show_legend`(bool)=true,
  `show_values`(bool)=false (data labels), `clean_style`(bool)=false
  (hides y-axis labels/gridlines/borders for a minimalist look),
  `value_format`(string — Excel number format, e.g. `"$#,##0\"M\""` or `"0.0%"`).
  `ex:`
  ```json
  {"type":"add_chart","slide":4,"chart_type":"columnclustered",
   "pos":{"left":60,"top":120,"width":560,"height":340},
   "categories":["FY21","FY22","FY23","FY24"],
   "series":[{"name":"Revenue ($M)","values":[120,138,151,170]},
             {"name":"EBITDA ($M)","values":[22,28,33,41]}],
   "title":"Revenue & EBITDA","show_legend":true,"value_format":"$#,##0\"M\""}
  ```
- **`set_chart_type`** — `req:` `slide`, `shape_id`(the chart), `value`(chart type string).
- **`set_chart_title`** — `req:` `slide`, `shape_id`, `value`(string). `opt:` `enabled`(bool)=true.
- **`set_chart_axis_title`** — `req:` `slide`, `shape_id`, `axis`(`category`|`value` / `x`|`y`), `value`(string).
- **`set_chart_legend_position`** — `req:` `slide`, `shape_id`, `value`(`top`|`right`|`bottom`|`left`|`none`).
- **`set_series_color`** — `req:` `slide`, `shape_id`, `series_index`(1-based), `value`(`#RRGGBB`).
- **`set_series_values`** — `req:` `slide`, `shape_id`, `series_index`, `values`(array of numbers).
- **`set_chart_categories`** — `req:` `slide`, `shape_id`, `categories`(array of strings).
- **`set_series_name`** — `req:` `slide`, `shape_id`, `series_index`, `value`(string).
- **`set_chart_axis`** — fine axis control. `req:` `slide`, `shape_id`, `axis`(`x`|`y`|`y2`|`x2`), `props`(object — any of: `visible`(bool — show/hide the whole axis; hiding the value axis also removes its gridlines), `line_visible`(bool — hide just the axis line), `min`(num), `max`(num), `major_unit`(num), `title`(string), `number_format`(string), `tick_label_position`(`low`|`high`|`next_to_axis`|`none`), `scale_type`(`linear`|`logarithmic`), `major_tick_mark`(`outside`|`inside`|`cross`|`none`)).
- **`set_chart_gridlines`** — show / hide / style chart gridlines. `req:` `slide`, `shape_id`, `props`. `opt:` `axis`(`x`|`category`|`y`|`value`|`both`)=`y`. `props`(object — any of: `major`(bool — show/hide major gridlines), `minor`(bool), `major_color`(`#RRGGBB`), `major_weight`(num pt), `major_dash`(`solid`|`dash`|`dot`|`round_dot`|`dash_dot`|`long_dash`|`long_dash_dot`), `minor_color`, `minor_weight`, `minor_dash`).
  `ex:` remove horizontal gridlines → `{"type":"set_chart_gridlines","slide":1,"shape_id":2,"props":{"major":false}}`
  `ex:` faint dotted gridlines on both axes → `{"type":"set_chart_gridlines","slide":1,"shape_id":2,"axis":"both","props":{"major":true,"major_color":"#E0E0E0","major_dash":"dot","major_weight":0.75}}`
- **`set_chart_format`** — chart-group props. `req:` `slide`, `shape_id`, `props`(object — any of: `gap_width`(0–500), `overlap`(-100–100), `bar_shape`(`box`|`cone`|`cone_to_max`|`cylinder`|`pyramid`|`pyramid_to_max`), `vary_by_categories`(bool), `reverse_categories`(bool), `reverse_series`(bool), `scale_type`(`linear`|`logarithmic`), `doughnut_hole_size`(10–90), **`plot_area_left`/`plot_area_top`/`plot_area_width`/`plot_area_height`**(num pt, pinning the plot rectangle inside the chart frame so caller-side overlays land on bars deterministically), `chart_area_fill`/`chart_area_border`/`plot_area_fill`/`plot_area_border`(`#RRGGBB`)).
- **`set_chart_series`** — per-series props. `req:` `slide`, `shape_id`, `series_index`, `props`(object — any of: `chart_type`(string, for combo charts), `axis_group`(`primary`|`secondary`), `marker_style`(`circle`|`square`|`triangle`|`diamond`|`x`|`none`), `marker_size`(num), `marker_fill`/`marker_line`(`#RRGGBB`), `line_color`/`line_weight`/`line_dash`(`solid`|`dash`|`dot`|`round_dot`|`dash_dot`|`long_dash`|`long_dash_dot`), `fill`/`fill_color`(`#RRGGBB`), `fill_visible`(bool — set false to hide a series visually while keeping it in the data, e.g. waterfall base), `show_labels`(bool), `label_format`(string Excel format), `custom_labels`(array — per-point label text override, supersedes `label_format`), `label_position`(`outside_end`/`above`|`inside_end`|`inside_base`|`center`|`below`|`left`|`right`), `label_color`/`label_size`/`label_bold`/`label_italic`/`label_fill`/`label_fill_visible`/`label_line_visible`, `point_fills`(array of `#RRGGBB` — recolor each bar individually in a single-series chart), `point_marker_fills`(array — per-point marker fill on line/scatter), **`point_marker_styles`**(array — per-point marker style; pass `"none"` to hide a single point's marker, e.g. to clip the last point of a line series), **`point_line_visible`**(array of bools — per-point segment visibility on a line series; `false` hides the line segment ending at point i, useful to break the line before a sentinel last point), `hide_from_legend`(bool), `pattern`(`dotted_5`…`zig_zag`), `gradient_fill`(object), `gradient_direction`(`horizontal`|`vertical`|`diagonal_up`|`diagonal_down`|`from_corner`|`from_center`)).
  > Tip: setting `set_chart_axis` `props.visible:false` removes the axis entirely (`HasAxis = False`) and breaks any series rendered against it (line series on a hidden secondary axis disappear). To hide an axis visually but keep its scale active, use `tick_label_position:"none"` + `line_visible:false` instead — the same recipe `add_chart` `clean_style:true` uses internally.
- **`set_chart_legend`** — `req:` `slide`, `shape_id`, `props`(object — `position`(`top`|`right`|`bottom`|`left`|`corner`), `visible`(bool), `font_size`(int)).
- **`add_chart_trendline`** — `req:` `slide`, `shape_id`, `series_index`, `props`(object — `kind`(`linear`|`log`|`polynomial`|`power`|`exponential`|`moving_avg`), `order`(int, for polynomial), `period`(int, for moving_avg), `display_equation`(bool), `display_r2`(bool), `dash`(`solid`|`dash`|`dot`|`round_dot`|`dash_dot`|`long_dash`|`long_dash_dot`), `color`(`#RRGGBB`))).
- **`set_chart_error_bars`** — `req:` `slide`, `shape_id`, `series_index`, `props`(object — `direction`(`x`|`y`|`both`), `include`(`both`|`plus`|`minus`), `type`(`fixed`|`percent`|`stdev`|`stderr`|`custom`), `amount`(num), `end_style`(`cap`|`no_cap`))).

### 3.12 Images & web

- **`insert_picture`** — insert a local image file. `req:` `slide`, `pos`, and one of `path` or `picture_path`(string — absolute local file path).
- **`replace_picture`** — swap an existing picture, keeping its frame. `req:` `slide`, `shape_id`, `path`(string).
- **`insert_icon`** — insert a Microsoft Fluent UI SVG icon.
  `req:` `slide`, `icon`(string — a lowercase_underscore name from the Fluent UI set, e.g. `building_factory`, `people`, `globe`, `chart_multiple`, `arrow_trending`, `money`, `shield`). `opt:` `style`(`filled`|`regular`)=`filled`, `size`(16|20|24|28|32|48)=48, `color`(`#RRGGBB`)=`#000000`, `left`,`top`,`width`,`height`(num pt). Icons are fetched from the unpkg CDN and cached.
  > If unsure of the exact icon name, pick the nearest semantic match. When Decko's Export prompt template injects an allow-list of icon names, use ONLY names from that list.
- **`fetch_page_images`** — scrape all images from a URL into a local folder. `req:` `url`(string). `opt:` `dest_folder`(string), `ref_name`(string).
- **`download_image`** — download one image URL to a local path. `req:` `url`(string), `dest_path`(string).
- **`open_image_picker`** — open Decko's visual image-picker UI on a folder. `req:` (none); `opt:` `folder`(string).
- **`build_image_picker_slide`** — build a thumbnail-grid slide from a folder of images. `opt:` `folder`(string), `cols`(int)=4, `insert_at`(int)=0 (0 = append), `max_per_slide`(int)=24.
- **`build_image_grid_table`** — build a 2-column (image | caption) table from a spec; pass `rows` as an array of `{ "image_path": string, "text": string }` plus `slide`, `pos`. (See `modActionsTable.Do_build_image_grid_table_act` for exact keys.)
- **`bulk_insert_image`** — same image, same box, on multiple slides. `req:` `slide_indices`(array of ints), `picture_path`(string), `left`,`top`,`width`,`height`(num).

### 3.13 Slides & deck-wide

- **`add_slide`** — `req:` `position`(int, 1-based; out-of-range → appended), `layout_index`(int, 0-based — see table).
- **`delete_slide`** — `req:` `slide`.
- **`duplicate_slide`** — clone (copy lands right after the source). `req:` `slide`.
- **`move_slide`** — reorder. `req:` `from_slide` (or `from`)(int), `to_slide` (or `to`)(int).
- **`extract_slides`** — export selected slides to a new `.pptx`. `req:` `slide_indices`(array of ints), `output_path`(string).
- **`import_slides_from_deck`** — pull slides from another deck. `req:` `source_path`(string), `slide_indices`(array of ints), `target_position`(int).
- **`apply_layout_to_slides`** — force a layout on listed slides. `req:` `slide_indices`(array), `layout_index`(int).
- **`apply_theme`** — apply a `.thmx`/`.potx`. `req:` `theme_path`(string).
- **`set_theme_font`** — set theme fonts. `opt:` `major`(string — heading font), `minor`(string — body font) (at least one).
- **`swap_font_deck_wide`** — replace one font name everywhere. `req:` `from_name`(string), `to_name`(string).
- **`recolor_palette_deck_wide`** — replace a color deck-wide. `req:` `from_hex`(`#RRGGBB`), `to_hex`(`#RRGGBB`), `target`(`fill`|`font`|`both`).
- **`set_slide_size`** — `req:` either `preset`(`16:9`|`4:3`) OR `width_pt`(num) & `height_pt`(num).
- **`bulk_insert_text_box`** — same text box on multiple slides. `req:` `slide_indices`(array), `text`(string), `left`,`top`,`width`,`height`(num).
- **`set_slide_background_color`** — solid background on a slide. `req:` `slide`, `color`(`#RRGGBB`).
- **`insert_slide_number`** — add a slide-number text placeholder. `req:` `slide`, `pos`. `opt:` `ref_name`(string), `font_color`(`#RRGGBB`), `font_size`(int).

### 3.14 Speaker notes

- **`set_speaker_notes`** — replace a slide's notes. `req:` `slide`, `value`(string).
- **`append_speaker_notes`** — append to existing notes. `req:` `slide`, `value`(string).

### 3.15 Visual effects (shapes & pictures)

- **`set_line_color`** — outline color. `req:` `slide`, `shape_id`, `value`(`#RRGGBB`).
- **`set_line_weight`** — `req:` `slide`, `shape_id`, `weight_pt`(num).
- **`set_line_style`** — `req:` `slide`, `shape_id`, `style`(`solid`|`dash`|`dot`|`dashdot`).
- **`set_shadow`** — drop shadow. `req:` `slide`, `shape_id`, `offset_x`(num), `offset_y`(num), `blur`(num), `color`(`#RRGGBB`), `transparency`(num 0.0–1.0).
- **`set_glow`** — outer glow. `req:` `slide`, `shape_id`, `color`(`#RRGGBB`), `radius`(num), `transparency`(num 0.0–1.0).
- **`set_reflection`** — `req:` `slide`, `shape_id`, `size`(num 0.0–1.0), `transparency`(num 0.0–1.0), `distance`(num pt).
- **`set_transparency`** — fill transparency. `req:` `slide`, `shape_id`, `value`(num 0.0–1.0).
- **`set_gradient_fill`** — two-color gradient. `req:` `slide`, `shape_id`, `color1`(`#RRGGBB`), `color2`(`#RRGGBB`), `angle`(num degrees).
- **`set_3d_bevel`** — `req:` `slide`, `shape_id`, `type`(`circle`|`slope`|`cross`|`angle`|`softround`), `depth_pt`(num).
- **`apply_preset_effect`** — Office texture/preset. `req:` `slide`, `shape_id`, `preset_index`(int 1–24).
- **`crop_picture`** — crop edges. `req:` `slide`, `shape_id`, `left`,`right`,`top`,`bottom`(num pt).
- **`recolor_picture`** — `req:` `slide`, `shape_id`, `color_type`(`grayscale`|`sepia`|`washout`|`bw`|`auto`).
- **`set_brightness`** — `req:` `slide`, `shape_id`, `value`(num -1.0–1.0).
- **`set_contrast`** — `req:` `slide`, `shape_id`, `value`(num -1.0–1.0).

---

## 4. Patterns the assistant should know

- **Building a new diagram (org chart, process flow):** `add_shape` for each box
  with a unique `ref_name`, then `add_connector` referencing those `ref_name`s
  with `from_shape_name`/`to_shape_name`. You don't need a snapshot for the new
  shapes — the `ref_name`s resolve within the batch.
- **Rebuilding a slide:** `clear_slide` (optionally `keep_shape_ids` for the
  title) → then a series of `add_text_box` / `add_shape` / `add_table` /
  `add_chart`. Lay out with explicit `pos` coordinates; a 16:9 slide is
  960 × 540 pt, leave ~40 pt margins.
- **Surgical text edits:** prefer `find_replace_text` (whole-deck literal swap,
  e.g. company name) or `set_run_text` (one run inside a mixed-format paragraph)
  over `set_text` (which flattens formatting).
- **Recoloring a deck to a new brand:** `recolor_palette_deck_wide` for each old
  → new color pair, and `swap_font_deck_wide` for the typeface.
- **Charts:** always supply `categories` and `series` for the 32 standard types.
  For the 7 modern types, the chart will appear with placeholder data; tell the
  VP to edit it. Use `clean_style: true` for a minimalist banker look.
- **Overflowing text:** `enable_text_shrink_for_overflow` on the slide or deck.
- **Ordering matters:** create shapes/slides/tables BEFORE the actions that
  reference them. `add_slide` before `add_chart` on that new slide. `add_table`
  before `set_cell_text` on it.

See **`docs/PROMPTING_GUIDE.md`** for worked end-to-end examples (VP English →
snapshot excerpt → exact `actions` JSON) and a step-by-step recipe.

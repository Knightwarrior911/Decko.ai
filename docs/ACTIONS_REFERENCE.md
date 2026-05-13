# Decko.ai — Complete Action Reference

This is the **machine-precise** spec of every action Decko can execute. It is
written for AI assistants (Hermes, OpenClaw, GPT-class, Claude-class — any model)
so that the assistant can (a) help a VP write a request and (b) emit the exact
JSON `actions` array a model should return.

**If you only read one thing, read [§0 Hard Rules](#0-hard-rules).** Breaking
any of them makes the whole batch fail.

> **Counts:** ~135 action types across 14 VBA modules. The authoritative source
> of truth is `src/modExecuteInstructions.bas` (`ValidateAction` = required
> fields, `DispatchAction` = how each field is read). This document mirrors it.
> When in doubt, the dispatcher wins — file a doc fix.

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
   - **Universal aliasing:** ANY field whose name ends in `_shape_id` or `shape_id`
     also accepts the parallel `_shape_name` / `shape_name` form (string ref_name
     or actual shape name). Same rule applies to plural `_shape_ids` ↔ `_shape_names`
     arrays. So: `from_shape_id`↔`from_shape_name`, `to_shape_id`↔`to_shape_name`,
     `ref_shape_id`↔`ref_shape_name`, `target_shape_id`↔`target_shape_name`,
     `target_shape_ids`↔`target_shape_names`, `source_shape_id`↔`source_shape_name`,
     `shape_a_id`↔`shape_a_name`, `shape_b_id`↔`shape_b_name`,
     `keep_shape_ids`↔`keep_shape_names`. Pick whichever you have; do NOT pass
     both for the same shape.
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
14. **Chart data shape:** in `add_chart`, every series's `values` array MUST have
    the same length as `categories`. Mismatched lengths cause a runtime error.
    `series_index` is **1-based** (first series = 1). `point_fills`, `custom_labels`,
    `point_marker_styles`, `point_line_visible` arrays inside `set_chart_series`
    `props` are also positionally aligned with the categories — pass one element
    per category.
15. **Hyperlink URL strict prefix:** `set_run_hyperlink` `value` MUST start with
    one of `http://`, `https://`, `mailto:`, or `#slide:N` (internal jump to slide
    N), OR be the empty string `""` (which clears the link). Anything else is
    rejected with `value: invalid hyperlink URL`.
16. **What Decko CANNOT do (don't ask):** no SmartArt creation/edit, no
    animations or slide transitions, no comments/review tracking, no embedded
    OLE objects (Excel/Word/audio/video insertion), no slide-show or kiosk
    settings, no master-slide editing (only layout-level via `apply_theme` /
    `apply_layout_to_slides`), no password/encryption, no slide-show pen/laser
    actions, no morph transitions. If the VP asks for any of these, say so
    instead of emitting fake actions.

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
| Shape line style (`set_line_style` `style`) | `solid` `dash` `dot` `dashdot` (strict — `round_dot`/`dash_dot` rejected here) |
| Chart line/dash (`*_dash` props on charts) | `solid` `dash` `dot` `round_dot` `dash_dot` `long_dash` `long_dash_dot` |
| Add-line `dash_style` (`add_line`/`add_connector`) | `solid` `dash` `dot` `round_dot` `dash_dot` `long_dash` `long_dash_dot` |
| Arrowhead (`arrow_end`/`arrow_start`) | `none` `filled` (`triangle`) `open` `stealth` `diamond` `oval` |
| Arrow size (`arrow_size`) | `small` `medium` `large` |
| Connector kind (`add_connector` `kind`) | `straight` `elbow` `curved` |
| Connector anchor point (`from_point`/`to_point`) | `auto` `top` `bottom` `left` `right` |
| Z-order (`z_order` `order`) | `front` `back` `forward` `backward` |
| Recolor target (`recolor_palette_deck_wide` `target`) | `fill` `font` `both` |
| Picture recolor (`recolor_picture` `color_type`) | `grayscale` `sepia` `washout` `bw` `auto` |
| 3D bevel (`set_3d_bevel` `type`) | `circle` `slope` `cross` `angle` `softround` |
| Nudge direction (`nudge` `direction`) | `l` `r` `u` `d` |
| Flip / center axis | `h` `v` (`align_to_slide_center` also accepts `both`) |
| Slide-size preset (`set_slide_size` `preset`) | `16:9` `4:3` |
| Chart axis name (`set_chart_axis` `axis`, `set_chart_gridlines` `axis`) | `x` (or `category`) · `y` (or `value`) · `y2` (or `secondary`) · `x2` · `both` (gridlines only) |
| Chart axis title axis (`set_chart_axis_title` `axis`) | `category` (`x`) · `value` (`y`) |
| Chart legend position (`set_chart_legend_position` `value`) | `top` `right` `bottom` `left` `none` |
| Chart legend position (`set_chart_legend` `props.position`) | `top` `right` `bottom` `left` `corner` (5 values — adds `corner`, drops `none`) |
| Label position (`set_chart_series` `props.label_position`) | `outside_end` `above` `inside_end` `inside_base` `center` `below` `left` `right` |
| Bar shape (`set_chart_format` `props.bar_shape`) | `box` `cone` `cone_to_max` `cylinder` `pyramid` `pyramid_to_max` |
| Marker style (`set_chart_series` `props.marker_style`) | `circle` `square` `triangle` `diamond` `x` `none` |
| Trendline kind | `linear` `log` `polynomial` `power` `exponential` `moving_avg` |
| Error-bar direction | `x` `y` `both` |
| Error-bar include | `both` `plus` `minus` |
| Error-bar type | `fixed` `percent` `stdev` `stderr` `custom` |
| Cell border side (`set_cell_border` `side`) | `top` `left` `bottom` `right` `diag_down` `diag_up` `all` |
| Edge (`match_position` `edge`) | `left` `right` `top` `bottom` `hcenter` `vcenter` |
| Anchor (`align_shapes` `anchor`) | `left` `right` `top` `bottom` `hcenter` `vcenter` |

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
  `req:` `slide`, `shape_id`, `index`(int, **1-based** — first handle = 1), `value`(num, 0.0–1.0-ish).
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
- **`add_text_box`** — plain text box (no fill/stroke by default). Field for content is **`text`**, not `value`.
  `req:` `slide`, `text`(string), `pos`.
  `opt:` `ref_name`, `font_color`, `font_size`(int), `font_bold`(bool)=false,
  `font_italic`(bool)=false, `h_align`(string), `fill`(`#RRGGBB`|null), `stroke`(`#RRGGBB`|null),
  `stroke_weight_pt`(num)=1.0, `super_suffix`, `sub_suffix`.
  > Note: `add_text_box` does **NOT** support `v_align` (use `set_text_vertical_align` afterwards via `ref_name`). Compare with `add_shape`, which DOES support both `h_align` and `v_align`.
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
- **`set_run_hyperlink`** — `req:` … `value`(string — MUST start with `http://`, `https://`, `mailto:`, or `#slide:N`; OR pass `""` to clear an existing link). Any other value is rejected at validation.
  `ex:` external → `{"type":"set_run_hyperlink","slide":1,"shape_id":3,"paragraph_index":0,"run_index":1,"value":"https://example.com"}`
  `ex:` internal jump → `{"type":"set_run_hyperlink","slide":1,"shape_id":3,"paragraph_index":0,"run_index":1,"value":"#slide:5"}`

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
- **`set_cell_border`** — `req:` `slide`, `shape_id`, `row`, `col`, `side`(`top`|`left`|`bottom`|`right`|`diag_down`|`diag_up`|`all`). `opt:` `color`(`#RRGGBB`), `weight_pt`(num), `visible`(bool)=true. (`all` sets all 4 outer edges; `diag_down`/`diag_up` add diagonal slash lines through the cell.)
- **`set_cell_text_align`** — `req:` `slide`, `shape_id`, `row`, `col`. `opt:` `h_align`(string), `v_align`(string) — at least one.
- **`set_cell_fill`** — `req:` `slide`, `shape_id`, `row`, `col`, `color`(`#RRGGBB`).
- **`apply_table_style`** — `req:` `slide`, `shape_id`, `style_id`(string — an Office table style GUID `{...}` **or** one of the named style keys below; keys are `lowercase_underscore`):
  `no_style_no_grid`, `no_style_with_grid`, `themed_style_1`, `themed_style_1_accent1`, `themed_style_1_accent2`, `themed_style_2`, `themed_style_2_accent1`, `medium_style_2`, `medium_style_2_accent1`, `medium_style_2_accent2`, `dark_style_2`, `dark_style_2_accent1`, `light_style_1`, `light_style_1_accent1`, `light_style_2`, `light_style_2_accent1`.
  `ex:` `{"type":"apply_table_style","slide":2,"shape_id":7,"style_id":"medium_style_2_accent1"}`
- **`build_image_grid_table`** — build a 2-column image+caption table from a row spec. See full schema in §3.12.

### 3.11 Charts

> See [§2 chart types](#chart-types-add_chart-chart_type-set_chart_type-value) for the full type list and the modern-type limitation.

- **`add_chart`** — insert a new native chart.
  `req:` `slide`, `chart_type`(string, see list), `pos`({left,top,width,height}),
  `categories`(array of strings — the x-axis labels),
  `series`(array of `{ "name": string, "values": [numbers], "color"?: "#RRGGBB" }`).
  > **Constraint:** every `series[i].values` array must have **exactly `categories.length` elements**. Mismatched lengths raise a runtime error. Use `0` for missing data points if you need to align; do not shorten/pad arbitrarily.
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
- **`set_chart_axis`** — fine axis control. `req:` `slide`, `shape_id`, `axis`(`x`|`category` / `y`|`value` / `y2`|`secondary` / `x2`|`category_secondary`), `props`(object — any of):
  - **Visibility:** `visible`(bool — show/hide whole axis; hiding the value axis also removes its gridlines), `line_visible`(bool — hide just the axis line)
  - **Scale:** `min`(num), `max`(num), `major_unit`(num — gap between major ticks), `minor_unit`(num), `scale_type`(`linear`|`logarithmic`)
  - **Tick labels:** `tick_label_position`(`low`|`high`|`next_to_axis`|`none`), `number_format`(string — Excel format e.g. `"$#,##0"` or `"0.0%"`), `label_color`(`#RRGGBB`), `label_size`(int), `label_bold`(bool), `label_italic`(bool), `label_rotation`(int, -90..90)
  - **Tick marks:** `major_tick_mark`(`outside`|`inside`|`cross`|`none`)
  - **Axis title:** `title`(string — turning on `HasTitle`), `title_size`(int), `title_color`(`#RRGGBB`), `title_bold`(bool), `title_italic`(bool)
- **`set_chart_gridlines`** — show / hide / style chart gridlines. `req:` `slide`, `shape_id`, `props`. `opt:` `axis`(`x`|`category`|`y`|`value`|`both`)=`y`. `props`(object — any of: `major`(bool — show/hide major gridlines), `minor`(bool), `major_color`(`#RRGGBB`), `major_weight`(num pt), `major_dash`(`solid`|`dash`|`dot`|`round_dot`|`dash_dot`|`long_dash`|`long_dash_dot`), `minor_color`, `minor_weight`, `minor_dash`).
  `ex:` remove horizontal gridlines → `{"type":"set_chart_gridlines","slide":1,"shape_id":2,"props":{"major":false}}`
  `ex:` faint dotted gridlines on both axes → `{"type":"set_chart_gridlines","slide":1,"shape_id":2,"axis":"both","props":{"major":true,"major_color":"#E0E0E0","major_dash":"dot","major_weight":0.75}}`
- **`set_chart_format`** — chart-group / chart-area / plot-area props. `req:` `slide`, `shape_id`, `props`(object — any of):
  - **Bars/columns:** `gap_width`(0–500), `overlap`(-100–100), `bar_shape`(`box`|`cone`|`cone_to_max`|`cylinder`|`pyramid`|`pyramid_to_max`), `vary_by_categories`(bool)
  - **Series order:** `reverse_categories`(bool), `reverse_series`(bool)
  - **Scale:** `scale_type`(`linear`|`logarithmic`)
  - **Doughnut:** `doughnut_hole_size`(10–90)
  - **Line-chart extras:** `drop_lines`(bool — vertical lines from each marker to the category axis), `hi_lo_lines`(bool — vertical line between max/min of multiple series at each category), `up_down_bars`(bool — diff boxes between two series)
  - **Chart area (outer frame):** `chart_area_fill`(`#RRGGBB`), `chart_area_fill_visible`(bool), `chart_area_border`(`#RRGGBB`), `chart_area_border_visible`(bool), `chart_area_image`(string — local file path; sets the chart-area background to a picture fill)
  - **Plot area (inner data region):** `plot_area_fill`(`#RRGGBB`), `plot_area_fill_visible`(bool), `plot_area_border`(`#RRGGBB`), `plot_area_image`(string — local file path)
  - **Plot-area pinning:** `plot_area_left`/`plot_area_top`/`plot_area_width`/`plot_area_height`(num pt, relative to chart frame top-left — pins the plot rectangle so caller-side overlays land on bars deterministically)
  - **3D-only:** `rotation`(deg 0–360), `elevation`(deg -90–90), `perspective`(0–100), `right_angle_axes`(bool), `height_percent`(int — Z-axis height as % of base), `gap_depth`(int — depth between series for 3D bar/column)
- **`set_chart_series`** — per-series props. `req:` `slide`, `shape_id`, `series_index`(1-based), `props`(object — any of):
  - **Identity / combo:** `name`(string — rename series), `chart_type`(string — for combo charts; use any `add_chart` `chart_type` value), `axis_group`(`primary`|`secondary`)
  - **Fill:** `fill`/`fill_color`(`#RRGGBB`), `fill_visible`(bool — hide a series visually while keeping its data, e.g. waterfall base), `border_visible`(bool — for filled series like columns/areas)
  - **Pattern fill (object):** `pattern_fill` = `{ "fore": "#RRGGBB", "back": "#RRGGBB", "type": <pattern-name> }` where pattern-name is one of `dotted_5/10/20/25/30/40/50/60/70/75/80/90`, `dark_horizontal`, `dark_vertical`, `dark_diagonal_down`, `dark_diagonal_up`, `light_horizontal`, `light_vertical`, `light_diagonal_down`, `light_diagonal_up`, `small_checker`, `small_grid`, `small_confetti`, `large_checker`, `large_grid`, `large_confetti`, `horizontal_brick`, `diagonal_brick`, `weave`, `plaid`, `divot`, `dotted_diamond`, `shingle`, `wave`, `zig_zag`, `trellis`
  - **Gradient fill (object):** `gradient_fill` = `{ "from": "#RRGGBB", "to": "#RRGGBB", "direction": "horizontal"|"vertical"|"diagonal_up"|"diagonal_down" }` (direction is nested **inside** `gradient_fill`)
  - **Line series:** `line_color`(`#RRGGBB`), `line_weight`(num pt), `line_dash`(`solid`|`dash`|`dot`|`round_dot`|`dash_dot`|`long_dash`|`long_dash_dot`)
  - **Markers:** `marker_style`(`circle`|`square`|`triangle`|`diamond`|`x`|`none`), `marker_size`(num), `marker_fill`(`#RRGGBB`), `marker_line`(`#RRGGBB`)
  - **Data labels:** `show_labels`(bool), `label_format`(string Excel format), `label_position`(`outside_end`|`above`|`inside_end`|`inside_base`|`center`|`below`|`left`|`right`), `label_color`/`label_size`/`label_bold`/`label_italic`(font props), `label_fill`(`#RRGGBB`), `label_fill_visible`(bool), `label_line_visible`(bool)
  - **Custom per-point label text:** `custom_labels`(array — per-point label text override; supersedes `label_format`; one element per category)
  - **Per-point overrides (arrays, one element per category):** `point_fills`(array of `#RRGGBB`), `point_marker_fills`(array of `#RRGGBB`), `point_marker_styles`(array — per-point marker style; pass `"none"` to hide one point, e.g. to clip last point of a line series), `point_line_visible`(array of bools — `false` hides the line segment ending at point i, useful to break the line before a sentinel last point), `point_label_colors`(array of `#RRGGBB`), `point_label_positions`(array of label-position strings)
  - **Pie/doughnut:** `show_leader_lines`(bool), `leader_line_color`(`#RRGGBB`)
  - **Legend:** `hide_from_legend`(bool)
  > Tip: setting `set_chart_axis` `props.visible:false` removes the axis entirely (`HasAxis = False`) and breaks any series rendered against it (line series on a hidden secondary axis disappear). To hide an axis visually but keep its scale active, use `tick_label_position:"none"` + `line_visible:false` instead — the same recipe `add_chart` `clean_style:true` uses internally.
- **`set_chart_legend`** — `req:` `slide`, `shape_id`, `props`(object — `visible`(bool), `position`(`top`|`right`|`bottom`|`left`|`corner`), `font_size`(int), `font_color`(`#RRGGBB`)).
  > Compared to `set_chart_legend_position`, this action's `position` vocab drops `none` but adds `corner`. To hide the legend, pass `visible:false` here OR pass `value:"none"` to `set_chart_legend_position`.
- **`add_chart_trendline`** — `req:` `slide`, `shape_id`, `series_index`, `props`(object — `kind`(`linear`|`log`|`polynomial`|`power`|`exponential`|`moving_avg`), `order`(int, for polynomial), `period`(int, for moving_avg), `display_equation`(bool), `display_r2`(bool), `dash`(`solid`|`dash`|`dot`|`round_dot`|`dash_dot`|`long_dash`|`long_dash_dot`), `color`(`#RRGGBB`))).
- **`set_chart_error_bars`** — `req:` `slide`, `shape_id`, `series_index`, `props`(object — `direction`(`x`|`y`|`both`), `include`(`both`|`plus`|`minus`), `type`(`fixed`|`percent`|`stdev`|`stderr`|`custom`), `amount`(num), `end_style`(`cap`|`no_cap`))).

### 3.12 Images & web

- **`insert_picture`** — insert a local image file. `req:` `slide`, `pos`, and one of `path` or `picture_path`(string — absolute local file path).
- **`replace_picture`** — swap an existing picture, keeping its frame. `req:` `slide`, `shape_id`, `path`(string).
- **`insert_icon`** — insert a Microsoft Fluent UI SVG icon.
  `req:` `slide`, `icon`(string — a lowercase_underscore name from the Fluent UI set, e.g. `building_factory`, `people`, `globe`, `chart_multiple`, `arrow_trending`, `money`, `shield`), **`left`, `top`, `width`, `height`** (num pt — all four REQUIRED, no `pos` object). `opt:` `style`(`filled`|`regular`)=`filled`, `size`(16|20|24|28|32|48)=48, `color`(`#RRGGBB`)=`#000000`, `ref_name`(string). Icons fetched from unpkg CDN and cached.
  `ex:` `{"type":"insert_icon","slide":2,"icon":"building_factory","left":60,"top":120,"width":48,"height":48,"color":"#15283C"}`
  > If unsure of the exact icon name, pick the nearest semantic match. When Decko's Export prompt template injects an allow-list of icon names, use ONLY names from that list.
- **`fetch_page_images`** — scrape all images from a URL into a local folder. `req:` `url`(string). `opt:` `dest_folder`(string), `ref_name`(string).
- **`download_image`** — download one image URL to a local path. `req:` `url`(string), `dest_path`(string).
- **`open_image_picker`** — open Decko's visual image-picker UI on a folder. `req:` (none); `opt:` `folder`(string).
- **`build_image_picker_slide`** — build a thumbnail-grid slide from a folder of images. `opt:` `folder`(string), `cols`(int)=4, `insert_at`(int)=0 (0 = append), `max_per_slide`(int)=24.
- **`build_image_grid_table`** — build a 2-column (image | description) table from a row-spec array. Full schema:
  `req:` `slide`, `pos`({left,top,width,height}), `rows`(array — see below).
  `opt:`
  - `ref_name`(string)
  - `image_col`(int)=1, `desc_col`(int)=2 — which column holds images vs captions
  - `name_position`(`top`|`bottom`)=`bottom` — where the caption sits inside the image cell
  - `name_strip_pt`(num)=30 — height of the caption strip in the image cell
  - `image_pad_pt`(num)=6 — padding around the image inside the cell
  - `image_fit`(`contain`|`stretch`)=`contain` — `contain` preserves aspect ratio (letterbox); `stretch` fills exactly
  - `header_row`(bool)=false — add a header row at the top
  - `header_text_image`(string)=`""`, `header_text_desc`(string)=`""` — header cell text
  - `col1_width_pt`(num), `col2_width_pt`(num) — explicit column widths
  - `name_font`(object: `size`, `bold`, `color`) — caption strip font
  - `desc_font`(object: `size`, `color`) — description column font

  Each element of `rows` is an object:
  - `name`(string) — caption text shown in the image cell's label strip
  - `image_path`(string) — absolute local file path  *(use this OR `image_url`)*
  - `image_url`(string) — URL; Decko downloads it automatically
  - `bullets`(array of strings) — text lines in the description cell

  `ex:`
  ```json
  {"type":"build_image_grid_table","slide":3,
   "pos":{"left":30,"top":80,"width":900,"height":420},
   "col1_width_pt":280,"col2_width_pt":620,
   "image_fit":"contain","name_position":"bottom","name_strip_pt":28,
   "name_font":{"size":11,"bold":true,"color":"#15283C"},
   "desc_font":{"size":10,"color":"#333333"},
   "rows":[
     {"name":"Aerospace","image_path":"C:\\images\\aero.jpg","bullets":["Leader in propulsion","3 platforms"]},
     {"name":"Defence","image_url":"https://example.com/def.jpg","bullets":["NATO-certified","$2B backlog"]}
   ]}
  ```
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


---

## Common Mistakes (read before writing any actions)

These are the most frequent errors from LLMs — each one causes silent skips or
formatting destruction that is hard to debug.

### Mistake 1 — Wrong payload field name for text content

Most text-setting actions use `"value"` — never `"text"`, `"content"`, `"color"`, `"size"`, or `"fill"`.

**Exception — `add_text_box` and `bulk_insert_text_box` use `"text"` (not `"value"`):**
```json
// RIGHT for add_text_box / bulk_insert_text_box
{"type":"add_text_box","slide":1,"text":"Label","pos":{"left":60,"top":80,"width":200,"height":40}}
{"type":"bulk_insert_text_box","slide_indices":[1,2,3],"text":"Confidential","left":800,"top":510,"width":120,"height":20}
```

For all other text actions (`set_text`, `set_paragraph_text`, `add_paragraph`, `set_run_text`, `set_font_color`, `set_font_size`, …) use `"value"`:

```json
// WRONG
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":0,"text":"Hello"}
{"type":"add_paragraph","slide":1,"shape_id":3,"after_paragraph_index":2,"text":"Hello"}
{"type":"set_font_color","slide":1,"shape_id":3,"color":"#FF0000"}

// RIGHT
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":0,"value":"Hello"}
{"type":"add_paragraph","slide":1,"shape_id":3,"after_paragraph_index":2,"value":"Hello"}
{"type":"set_font_color","slide":1,"shape_id":3,"value":"#FF0000"}
```

### Mistake 2 — `scope` on shape-level actions

`scope` is only valid on 8 deck/slide-sweep actions:
`find_replace_text`, `find_replace_regex`, `recolor_palette_deck_wide`,
`recolor_fill_match`, `recolor_font_match`, `delete_shapes_match`,
`enable_text_shrink_for_overflow`, `swap_font_deck_wide`.

Any action that has a `slide` field must NOT have `scope`. The parser reads
`scope` as the slide specifier and then fails to find `slide`, causing
"missing_field: slide".

```json
// WRONG — scope on a shape-level action
{"type":"set_paragraph_text","scope":"slide:1","slide":1,"shape_id":3,"paragraph_index":0,"value":"Hello"}

// RIGHT
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":0,"value":"Hello"}
```

### Mistake 3 — `set_text` on a multi-paragraph / multi-level bullet box

`set_text` replaces the entire text frame as a single string. It destroys all
per-paragraph formatting — every paragraph inherits the first paragraph's color,
size, and bold, making all bullets look like the heading.

Use `set_paragraph_text` (per paragraph_index) when the shape has multiple
paragraphs with different formatting. Only use `set_text` for plain single-style
shapes or when the user explicitly says "replace everything".

```json
// WRONG — destroys all per-paragraph colors/sizes
{"type":"set_text","slide":1,"shape_id":3,"value":"Heading\nBullet 1\nBullet 2"}

// RIGHT — text changes only; existing per-paragraph formatting preserved
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":0,"value":"Heading"}
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":1,"value":"Bullet 1"}
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":2,"value":"Bullet 2"}
```

### Mistake 4 — Setting font on existing bullet paragraphs when only text changes

If the snapshot already shows the right colors/sizes per paragraph level, only
emit `set_paragraph_text`. Adding `set_font_color`/`set_font_size` resets the
formatting and overwrites the per-level styling.

**Exception:** NEW paragraphs added with `add_paragraph` have no inherited
formatting and MUST have `set_paragraph_font_color`, `set_paragraph_font_size`,
`set_indent_level` set explicitly.

```json
// WRONG — existing para already has correct color; setting it again may override
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":1,"value":"New text"}
{"type":"set_font_color","slide":1,"shape_id":3,"value":"#000000"}   // <-- resets ALL paragraphs

// RIGHT — existing para: text only
{"type":"set_paragraph_text","slide":1,"shape_id":3,"paragraph_index":1,"value":"New text"}

// RIGHT — new added para: needs explicit font
{"type":"add_paragraph","slide":1,"shape_id":3,"after_paragraph_index":3,"value":"New bullet"}
{"type":"set_indent_level","slide":1,"shape_id":3,"paragraph_index":4,"value":1}
{"type":"set_paragraph_font_color","slide":1,"shape_id":3,"paragraph_index":4,"value":"#000000"}
{"type":"set_paragraph_font_size","slide":1,"shape_id":3,"paragraph_index":4,"value":10}
```

### Mistake 5 — Shape-level font actions on multi-paragraph shapes

`set_font_size` and `set_font_color` without `paragraph_index` apply to the
**entire shape** — all paragraphs get the same size/color, erasing the
per-level differences (heading navy, L1 black, L2 blue, L3 gray).

Use `set_paragraph_font_size` / `set_paragraph_font_color` with a specific
`paragraph_index` to change only one paragraph's font.

```json
// WRONG — flattens all paragraph colors to one value
{"type":"set_font_color","slide":1,"shape_id":3,"value":"#1F3864"}

// RIGHT — changes only paragraph 0 (the heading)
{"type":"set_paragraph_font_color","slide":1,"shape_id":3,"paragraph_index":0,"value":"#1F3864"}
```

### Mistake 6 — `v_align` on `add_text_box` (it's silently ignored)

`add_text_box` only reads `h_align`. If you pass `v_align`, it does nothing — the textbox keeps the default top alignment. To vertically center, create the box first with a `ref_name`, then use `set_text_vertical_align` referencing that name:

```json
// WRONG — v_align ignored
{"type":"add_text_box","slide":1,"text":"Centered","pos":{"left":100,"top":200,"width":300,"height":60},"v_align":"middle"}

// RIGHT
{"type":"add_text_box","slide":1,"text":"Centered","pos":{"left":100,"top":200,"width":300,"height":60},"ref_name":"tb1"}
{"type":"set_text_vertical_align","slide":1,"shape_id":"tb1","value":"middle"}
```

`add_shape` DOES accept both `h_align` and `v_align` directly — the asymmetry is unfortunate but real.

### Mistake 7 — `insert_icon` uses flat `left/top/width/height` (NOT `pos`)

Unlike `add_shape`/`add_text_box`/`add_chart`/`add_table`, the icon action does NOT take a `pos` object. All four position fields are required at the action's top level:

```json
// WRONG
{"type":"insert_icon","slide":2,"icon":"people","pos":{"left":60,"top":120,"width":48,"height":48}}

// RIGHT
{"type":"insert_icon","slide":2,"icon":"people","left":60,"top":120,"width":48,"height":48}
```

Same flat-coordinate convention applies to `bulk_insert_image` and `bulk_insert_text_box`.

### Mistake 8 — Mixing `scope` semantics across actions

Three different `scope`-using action families read it slightly differently:
- `find_replace_text`, `find_replace_regex`: `deck` or `slide:N` — searches text frames only.
- `recolor_fill_match`, `recolor_font_match`, `recolor_palette_deck_wide`, `delete_shapes_match`: `deck` or `slide:N` — sweeps shapes for color/kind/text filters.
- `swap_font_deck_wide`: **no `scope`** — always deck-wide (`from_name`/`to_name` only).
- `enable_text_shrink_for_overflow`: `deck` or `slide:N`, plus optional `include_titles`(bool).

If you mean "this slide only", use `slide:N`. If you mean "the whole deck", use `deck`. Don't pass numeric values (the parser will reject them).

### Validator error glossary (when an action is skipped, the log shows one of these)

| Error string | Cause | Fix |
|---|---|---|
| `unknown_type: X` | `"type"` field not recognised | Spelling — see catalogue |
| `missing_field: X` | Required field absent | Re-read `req:` for the action |
| `missing_field: X or Y` | Either-or field pair missing both | Provide one of the two |
| `slide must be numeric (1-based)` | `slide` was a string | Pass an integer, not `"1"` |
| `slide_out_of_range: N (deck has M)` | `slide > slide count` (or < 1) | Add a slide first or fix the index |
| `shape_id 'X': not found as Id or ref_name` | Shape doesn't exist on that slide | Re-check snapshot; remember `ref_name`s only resolve in the same batch |
| `value: must be ...` | String/number outside the allowed vocab | Match the vocab tables in §2 |
| `value: invalid hyperlink URL` | URL doesn't start with `http://`/`https://`/`mailto:`/`#slide:` | Use one of the four prefixes |
| `width_pt/height_pt: must be > 0` | Zero or negative size | Use a positive number |
| `preset: must be 16:9 or 4:3` | `set_slide_size` preset typo | Lowercase, exact `16:9` or `4:3` |
| `from_name/to_name: empty` | `swap_font_deck_wide` got blank strings | Both must be non-empty font names |
| `axis: must be h or v` | `flip_shape`/`smart_spacing` got something else | Use single-letter `h`/`v` |

When the batch finishes, a JSON report appears with `actions_executed` / `actions_skipped` counts and per-action `error` strings. **Read those before assuming success.**

### Pattern: populate bullet box with variable-length refined text

Use this sequence regardless of how many bullets the refinement produces:

1. `set_paragraph_text` for p0 (heading) — text only, no font actions
2. `set_paragraph_text` for each existing paragraph p1…pN — text only
3. For refined text that has MORE points than existing paragraphs:
   - `add_paragraph` with `after_paragraph_index` = last written index
   - `set_indent_level` on the new paragraph
   - `set_paragraph_font_color` matching the level (L1=#000000, L2=#264F82, L3=#707070)
   - `set_paragraph_font_size` = same pt as existing bullets (read from snapshot)
4. Last action: `enable_text_shrink_for_overflow` on the slide

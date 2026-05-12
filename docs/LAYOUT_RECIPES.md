# Decko.ai — Layout Recipes

How to take an existing slide and **redesign its whole layout** — split it into
columns, stack regions, build a quad, or explode a bullet list into a visual
arrangement of cards. Written for AI assistants (Hermes, OpenClaw, GPT, Claude)
that drive Decko by emitting `{"actions":[...]}` from a snapshot.

There is **no special "layout" action**. A layout redesign is just a batch of
ordinary actions:

- **Re-arrange existing objects** → a batch of `move_shape` + `resize_shape`
  (one pair per object). Target each object by `shape_name` if it has a
  `ref_name`, otherwise by numeric `shape_id` (icons / connectors / pictures
  usually need the id).
- **Transform an object** (table → chart, chart → table, paragraph box → cards)
  → `delete_shape` the old one, then `add_chart` / `add_table` / `add_shape`
  in the target region.
- **Add structure** → `add_line` rules between regions, `add_text_box` region
  headers.
- Always end the batch with `enable_text_shrink_for_overflow` (`scope:"slide:N"`
  or `"deck"`) so anything that got narrower auto-shrinks to fit.

## Hard rules

1. **Read the snapshot for the slide size.** It is in `deck.slide_width_pt` /
   `deck.slide_height_pt`. The presets below assume the default **960 × 540 pt**
   (16:9). For 4:3 (720 × 540) or any other size, recompute from the formulas.
2. **Coordinates are points, origin top-left.** `left`/`top` = position,
   `width`/`height` = size.
3. **Don't change text or colors** unless the request asks for it. A layout
   redesign moves and resizes; it does not rewrite content.
4. **Pick the slide by content, not by number** — find the slide whose objects
   match the request.
5. **Tables auto-grow their row heights.** A table will often end a few pt
   *taller* than the `height` you set. Leave slack below tables; never butt a
   table right up against the slide edge.
6. **Charts, tables, pictures fill their region** (minus a small inset, ~0–8 pt).
   Several small objects in one region → sub-stack them vertically inside it.
7. **Order:** do the `delete_shape` first (frees the space), then `add_*`, then
   the `move_shape`/`resize_shape` re-arrangements, then `add_line` rules, then
   `enable_text_shrink_for_overflow` last.

## The working area (960 × 540 pt)

| Band | Range | Use |
|------|-------|-----|
| Outer margin | 40 pt all sides | content lives in **x 40 → 920** (w 880), **y 40 → 500** |
| Title band | y **14 → 52** (h ~38) | the slide title text box — usually left untouched |
| Content band | y **64 → 516** (h ~452) | everything else |
| Footer band | y **520 → 536** (h ~16) | source line / footnote — full width |
| Gutter | **16 pt** between regions (use 14 for 3–4 columns) | |

Generic formula for an N-way split of a span `[a, b]` with gutter `g`:
`size = (b - a - (N-1)*g) / N`; region *i* (0-based) starts at `a + i*(size + g)`.

---

## Region presets — 960 × 540 pt

Coordinates are `left, top, width, height`. Content band is y 64 → 516 (h 452);
divider lines sit in the gutter.

### Full content (single region)
`40, 64, 880, 452`

### Left / Right vertical splits (gutter 16 in the middle)
| Preset | Left region | Right region | Divider |
|--------|-------------|--------------|---------|
| **L50 / R50** | `40, 64, 432, 452` | `488, 64, 432, 452` | vertical line at x **480** |
| **L67 / R33** | `40, 64, 580, 452` | `636, 64, 284, 452` | vertical line at x **624** |
| **L33 / R67** | `40, 64, 284, 452` | `340, 64, 580, 452` | vertical line at x **324** |
| **L60 / R40** | `40, 64, 520, 452` | `576, 64, 344, 452` | vertical line at x **564** |

### Top / Bottom horizontal splits (gutter 16)
| Preset | Top region | Bottom region | Divider |
|--------|------------|---------------|---------|
| **Top50 / Bot50** | `40, 64, 880, 218` | `40, 298, 880, 218` | horizontal line at y **290** |
| **Top33 / Bot67** | `40, 64, 880, 138` | `40, 218, 880, 298` | horizontal line at y **210** |
| **Top67 / Bot33** | `40, 64, 880, 298` | `40, 378, 880, 138` | horizontal line at y **370** |
| **Header row + body** | header `40, 64, 880, 56` | body `40, 128, 880, 388` | (no line needed) |

### Quad (2 × 2), gutter 16
| Cell | Coordinates |
|------|-------------|
| Top-left | `40, 64, 432, 218` |
| Top-right | `488, 64, 432, 218` |
| Bottom-left | `40, 298, 432, 218` |
| Bottom-right | `488, 298, 432, 218` |

### Three columns, gutter 14
| Column | Coordinates |
|--------|-------------|
| C1 | `40, 64, 284, 452` |
| C2 | `338, 64, 284, 452` |
| C3 | `636, 64, 284, 452` |

### Four columns, gutter 14
| Column | Coordinates |
|--------|-------------|
| C1 | `40, 64, 209, 452` |
| C2 | `263, 64, 209, 452` |
| C3 | `486, 64, 209, 452` |
| C4 | `709, 64, 209, 452` |

### 2 stacked left + 1 full right (gutter 16)
| Region | Coordinates |
|--------|-------------|
| Left-top | `40, 64, 580, 218` |
| Left-bottom | `40, 298, 580, 218` |
| Right (full height) | `636, 64, 284, 452` |
| Dividers | vertical x **624**; horizontal x 40→620 at y **290** |

### 1 full left + 2 stacked right (mirror)
| Region | Coordinates |
|--------|-------------|
| Left (full height) | `40, 64, 580, 452` |
| Right-top | `636, 64, 284, 218` |
| Right-bottom | `636, 298, 284, 218` |
| Dividers | vertical x **624**; horizontal x 636→920 at y **290** |

### 2 rows × 3 columns (6 cells), gutter 14 / 16
| | C1 | C2 | C3 |
|--|----|----|----|
| **Row 1** | `40, 64, 284, 218` | `338, 64, 284, 218` | `636, 64, 284, 218` |
| **Row 2** | `40, 298, 284, 218` | `338, 298, 284, 218` | `636, 298, 284, 218` |

### Header row + N-column body
Header band `40, 64, 880, 56`; body band y 128 → 516 (h 388), split into N columns
by the column formula (N=2 → w 432 each at x 40 / 488; N=3 → w 284 each at x 40 / 338 / 636).

> If the slide keeps a footer/source line, shorten any full-height region to
> end at y **516** (already assumed above) and put the footer at `40, 520, 880, 16`.

---

## Placing content inside a region

- **A chart / table / picture** → set its `pos` (or `move_shape` + `resize_shape`)
  to the region rectangle, optionally inset 0–8 pt. Tables: leave ~10–20 pt of
  slack below (row auto-grow).
- **Several small objects in one region** → sub-stack vertically: split the
  region's height by the number of objects (gutter ~8–12 pt).
- **Icon + caption rows** → icon is a small square (22–28 pt) at the region's
  left edge; caption text box starts ~6 pt to its right and takes the rest of
  the width; rows ~30–36 pt tall.
- **A "card"** → `add_shape` `kind:"rrect"` filling the region (inset a few pt),
  `fill` a brand color, `text` the content, `font_color` for contrast (white on
  dark fills, navy/black on light fills), `v_align:"middle"`, `h_align:"left"`.
- **A region header** → `add_text_box` at the top of the region, ALL-CAPS,
  bold, 11 pt, brand color, 16–18 pt tall.

---

## Recipe: a single box of bullets → a visual layout

This is the most common request: *"I wrote my points as bullets in one text box;
make it visual."* Mechanics:

1. **Read the bullets.** In the snapshot, the box has a `paragraphs` array; each
   `paragraphs[i].text` is one bullet. Count them → `N`.
2. **Pick a layout by `N`** (unless the user named one):
   - `N = 2` → **L50 / R50** (two big cards) — or **Top50 / Bot50**.
   - `N = 3` → **three columns**, or **1 full left + 2 stacked right** (use the
     left card for the headline/most-important point).
   - `N = 4` → **quad 2×2**, or **four columns**.
   - `N = 5` → quad + one wide card across the bottom, or a 2×3 grid with one
     cell empty.
   - `N = 6` → **2 rows × 3 columns**.
   - Bullets that are *sequential steps* → a **chevron process flow**
     (`add_shape kind:"chevron"` in a row, left → right).
   - Bullets that are *milestones with dates* → a **timeline** (a horizontal
     `add_line`, alternating `add_text_box`/`add_shape` above and below).
3. **Delete the bullet box** (`delete_shape`), keep the slide title.
4. **Add one card per bullet** (`add_shape kind:"rrect"`, or `add_text_box` for
   a flat look) positioned at that bullet's region. Use the bullet's text; you
   may shorten a long bullet and/or prefix a number badge (`"1   "`). Vary the
   fill across cards (e.g. a brand-color ramp dark → light) for visual rhythm.
5. **`enable_text_shrink_for_overflow`** (`scope:"slide:N"`) so the cards' text
   shrinks to fit.

If you also want a one-line lead-in above the cards, add a small `add_text_box`
between the title and the cards (and pull the cards' `top` down ~24 pt).

### Worked example — 4 bullets → quad 2×2 cards

*Snapshot excerpt:* `Slide 1: shape_id=2 = title "Our FY26 Operating Plan"; shape_id=3 = text box with 4 paragraphs: ["Accelerate enterprise ARR ...","Ship the AI copilot to GA ...","Expand gross margin +200 bps ...","Deepen the partner channel ..."]`
*VP prompt:* `On slide 1, take those four bullets out of the text box and lay them out as a 2x2 grid of rounded cards, dark-to-light, numbered. Keep the title.`
*Actions:*
```json
{"actions":[
  {"type":"delete_shape","slide":1,"shape_id":3},
  {"type":"add_shape","slide":1,"kind":"rrect","pos":{"left":40,"top":64,"width":432,"height":218},"fill":"#1F3864","text":"1   Accelerate enterprise ARR — focused land-and-expand in the top 200 named accounts","font_color":"#FFFFFF","font_size":15,"h_align":"left","v_align":"middle","ref_name":"q1"},
  {"type":"add_shape","slide":1,"kind":"rrect","pos":{"left":488,"top":64,"width":432,"height":218},"fill":"#2F5597","text":"2   Ship the AI copilot to GA and drive attach across the installed base","font_color":"#FFFFFF","font_size":15,"h_align":"left","v_align":"middle","ref_name":"q2"},
  {"type":"add_shape","slide":1,"kind":"rrect","pos":{"left":40,"top":298,"width":432,"height":218},"fill":"#8FAADC","text":"3   Expand gross margin +200 bps via the multi-cloud cost-optimization program","font_color":"#FFFFFF","font_size":15,"h_align":"left","v_align":"middle","ref_name":"q3"},
  {"type":"add_shape","slide":1,"kind":"rrect","pos":{"left":488,"top":298,"width":432,"height":218},"fill":"#D9E2F3","text":"4   Deepen the partner channel — ecosystem pipeline reaches one third of new bookings","font_color":"#1F3864","font_size":15,"h_align":"left","v_align":"middle","ref_name":"q4"},
  {"type":"enable_text_shrink_for_overflow","scope":"slide:1"}
]}
```

### Worked example — 3 bullets → 1 big left + 2 stacked right

*VP prompt:* `Slide 2 — pull the three pillars out of that box; make the first one a big card filling the left, and the other two as stacked cards on the right third.`
*Actions:*
```json
{"actions":[
  {"type":"delete_shape","slide":2,"shape_id":3},
  {"type":"add_shape","slide":2,"kind":"rrect","pos":{"left":40,"top":64,"width":580,"height":452},"fill":"#1F3864","text":"1\n\nCost discipline — take out 15% of run-rate opex; consolidate vendors & real estate","font_color":"#FFFFFF","font_size":20,"h_align":"left","v_align":"middle","ref_name":"big"},
  {"type":"add_shape","slide":2,"kind":"rrect","pos":{"left":636,"top":64,"width":284,"height":218},"fill":"#2F5597","text":"2\n\nRevenue quality — shift the mix toward recurring software, away from low-margin services","font_color":"#FFFFFF","font_size":14,"h_align":"left","v_align":"middle","ref_name":"r1"},
  {"type":"add_shape","slide":2,"kind":"rrect","pos":{"left":636,"top":298,"width":284,"height":218},"fill":"#8FAADC","text":"3\n\nTalent density — top-performer attrition below 8%; rebuild the senior engineering bench","font_color":"#FFFFFF","font_size":14,"h_align":"left","v_align":"middle","ref_name":"r2"},
  {"type":"enable_text_shrink_for_overflow","scope":"slide:2"}
]}
```

---

## Recipe: re-flow a slide of mixed objects into a two-column layout

*Snapshot excerpt:* `Slide 3, 960×540. Objects: m_chart (chart), m_table (table), m_box1/m_box2 (rrect boxes), icons id 8 & 11 + captions m_cap1/m_cap2, m_foot (footnote).`
*VP prompt:* `Re-flow slide 3 into 67/33. Financials on the left two-thirds: the chart on top, the KPI table under it. The two priority boxes and the icon callouts go in the right third, stacked. Footnote across the bottom. Add a vertical rule between. Don't change text or colors.`
*Actions:*
```json
{"actions":[
  {"type":"move_shape","slide":3,"shape_id":"m_chart","left":40,"top":64},
  {"type":"resize_shape","slide":3,"shape_id":"m_chart","width":580,"height":200},
  {"type":"move_shape","slide":3,"shape_id":"m_table","left":40,"top":276},
  {"type":"resize_shape","slide":3,"shape_id":"m_table","width":580,"height":130},
  {"type":"move_shape","slide":3,"shape_id":"m_box1","left":636,"top":64},
  {"type":"resize_shape","slide":3,"shape_id":"m_box1","width":284,"height":60},
  {"type":"move_shape","slide":3,"shape_id":"m_box2","left":636,"top":130},
  {"type":"resize_shape","slide":3,"shape_id":"m_box2","width":284,"height":60},
  {"type":"move_shape","slide":3,"shape_id":8,"left":636,"top":206},
  {"type":"resize_shape","slide":3,"shape_id":8,"width":24,"height":24},
  {"type":"move_shape","slide":3,"shape_id":"m_cap1","left":666,"top":206},
  {"type":"resize_shape","slide":3,"shape_id":"m_cap1","width":254,"height":24},
  {"type":"move_shape","slide":3,"shape_id":11,"left":636,"top":238},
  {"type":"resize_shape","slide":3,"shape_id":11,"width":24,"height":24},
  {"type":"move_shape","slide":3,"shape_id":"m_cap2","left":666,"top":238},
  {"type":"resize_shape","slide":3,"shape_id":"m_cap2","width":254,"height":24},
  {"type":"move_shape","slide":3,"shape_id":"m_foot","left":40,"top":520},
  {"type":"resize_shape","slide":3,"shape_id":"m_foot","width":880,"height":14},
  {"type":"add_line","slide":3,"x1":624,"y1":60,"x2":624,"y2":512,"color":"#BFBFBF","weight_pt":1.0},
  {"type":"enable_text_shrink_for_overflow","scope":"slide:3"}
]}
```

Same idea for **quad / 2-left+1-right / 50-50 / 3-column** etc — just assign each
object to a region from the catalog and emit the `move_shape`+`resize_shape` pair.

---

## Recipe: table → chart (or chart → table) in a layout slot

No single action. The brain reads the source's data out of the snapshot, deletes
it, and adds the replacement sized to the target region. See `EXAMPLES.md` §8.3.
Key points: give the new chart a `ref_name` so the `set_chart_*` follow-ups can
target it; `clean_style:true` on `add_chart` hides the value axis + gridlines for
a compact look; if the rows are on incompatible scales (%, $, counts), chart a
subset or normalize — and say so. Reverse direction: `delete_shape` the chart,
`add_table` + `set_cell_text` per cell in the same `pos`.

---

## Checklist before emitting a layout-redesign batch

- [ ] Read `deck.slide_width_pt` / `slide_height_pt` — using the right grid?
- [ ] Identified every object on the target slide and which region it goes to?
- [ ] `delete_shape` for any transforms comes *first*?
- [ ] Tables given vertical slack (don't touch the slide edge)?
- [ ] Title kept in the title band; footnote (if any) full-width at the bottom?
- [ ] Added the gutter divider line(s) if the request implies separated sections?
- [ ] Batch ends with `enable_text_shrink_for_overflow`?
- [ ] No text/color changes that weren't asked for?

# Decko.ai — Worked Examples (VP prompt → actions JSON)

A corpus of paired examples: a **VP prompt** (the English request a VP types) and
the exact **actions JSON** a model should return for it. Many examples include a
**snapshot excerpt** showing the shape IDs / slide info the model would have been
given — the model needs that to write correct `shape_id`s and `slide`s.

Read alongside:
- **[`PROMPTING_GUIDE.md`](PROMPTING_GUIDE.md)** — workflow, how to phrase
  requests, the step-by-step recipe.
- **[`ACTIONS_REFERENCE.md`](ACTIONS_REFERENCE.md)** — every action's full
  schema (required/optional fields, value vocabularies).

**Conventions used below:** slides are 1-based; `shape_id` is the number in the
snapshot; all sizes are points (16:9 slide = 960 × 540 pt); colors are `#RRGGBB`;
the output is always `{"actions":[ ... ]}` and nothing else.

---

## 1. Plain text edits (whole shape)

### 1.1 Change a title

*Snapshot excerpt:*
```
Slide 1 (layout: Title Slide, 960x540)
  shape_id=2  kind=placeholder(title)  box=(48,200,864,120)  text: "Project Falcon — Internal Draft"
```
*VP prompt:* `On slide 1, change the title to "Project Falcon — Board Review, Q3 FY24".`
*Actions:*
```json
{"actions":[
  {"type":"set_text","slide":1,"shape_id":2,"value":"Project Falcon — Board Review, Q3 FY24"}
]}
```

### 1.2 Bold + recolor + resize a callout text box

*Snapshot excerpt:*
```
Slide 3
  shape_id=9  kind=textbox  box=(620,40,300,40)  text: "DRAFT — NOT FINAL"
```
*VP prompt:* `On slide 3, make the "DRAFT — NOT FINAL" box bold, red (#C00000), 16pt, and shrink it to width 220, height 30.`
*Actions:*
```json
{"actions":[
  {"type":"set_font_bold","slide":3,"shape_id":9,"value":true},
  {"type":"set_font_color","slide":3,"shape_id":9,"value":"#C00000"},
  {"type":"set_font_size","slide":3,"shape_id":9,"value":16},
  {"type":"resize_shape","slide":3,"shape_id":9,"width":220,"height":30}
]}
```

### 1.3 Move and delete shapes

*Snapshot excerpt:*
```
Slide 4
  shape_id=5  kind=picture   box=(40,300,200,140)   (old logo)
  shape_id=6  kind=textbox   box=(40,460,400,30)    text: "Confidential"
```
*VP prompt:* `On slide 4, delete the old logo (shape 5) and move the "Confidential" line to left 40, top 510.`
*Actions:*
```json
{"actions":[
  {"type":"delete_shape","slide":4,"shape_id":5},
  {"type":"move_shape","slide":4,"shape_id":6,"left":40,"top":510}
]}
```

---

## 2. Paragraph-level text (bulleted bodies)

### 2.1 Replace one bullet, add another, delete a third

*Snapshot excerpt:*
```
Slide 5
  shape_id=4  kind=placeholder(body)  box=(60,140,840,300)
    p0: "Revenue up 12% YoY"
    p1: "Margin flat at 18%"
    p2: "Headcount +40"
```
*VP prompt:* `On slide 5 in the body box (shape 4): change bullet 1 (index 1) to "Operating margin expanded 120 bps to 19.2%", add a new bullet after it reading "Free cash flow of $84M", and delete the headcount bullet (index 2).`
*Actions:*
```json
{"actions":[
  {"type":"set_paragraph_text","slide":5,"shape_id":4,"paragraph_index":1,"value":"Operating margin expanded 120 bps to 19.2%"},
  {"type":"add_paragraph","slide":5,"shape_id":4,"after_paragraph_index":1,"value":"Free cash flow of $84M"},
  {"type":"delete_paragraph","slide":5,"shape_id":4,"paragraph_index":2}
]}
```
> Note: after `set_paragraph_text` and `add_paragraph` (which inserts at index 2),
> the old "Headcount +40" paragraph has shifted to index 3 — but `delete_paragraph`
> here uses index 2, which is now "Free cash flow of $84M"… To avoid index drift,
> **do deletions first, or delete from the bottom up.** Safer version:
```json
{"actions":[
  {"type":"delete_paragraph","slide":5,"shape_id":4,"paragraph_index":2},
  {"type":"set_paragraph_text","slide":5,"shape_id":4,"paragraph_index":1,"value":"Operating margin expanded 120 bps to 19.2%"},
  {"type":"add_paragraph","slide":5,"shape_id":4,"after_paragraph_index":1,"value":"Free cash flow of $84M"}
]}
```

### 2.2 Prepend a bullet (`after_paragraph_index: -1`)

*VP prompt:* `On slide 5, add "Q3 FY24 highlights:" as a new first bullet in the body box (shape 4).`
*Actions:*
```json
{"actions":[
  {"type":"add_paragraph","slide":5,"shape_id":4,"after_paragraph_index":-1,"value":"Q3 FY24 highlights:"}
]}
```

### 2.3 Bullet style, indent, alignment, line spacing

*VP prompt:* `On slide 5, body box (shape 4): make bullet 0 a numbered list item, indent bullets 1 and 2 one level, center bullet 0, and set 1.5 line spacing on bullet 0.`
*Actions:*
```json
{"actions":[
  {"type":"set_bullet_style","slide":5,"shape_id":4,"paragraph_index":0,"value":"number"},
  {"type":"set_indent_level","slide":5,"shape_id":4,"paragraph_index":1,"value":1},
  {"type":"set_indent_level","slide":5,"shape_id":4,"paragraph_index":2,"value":1},
  {"type":"set_paragraph_alignment","slide":5,"shape_id":4,"paragraph_index":0,"value":"center"},
  {"type":"set_paragraph_line_spacing","slide":5,"shape_id":4,"paragraph_index":0,"value":1.5}
]}
```

---

## 3. Run-level edits (mixed-format paragraphs)

### 3.1 Change only the bold drug name, keep the description plain

*Snapshot excerpt:*
```
Slide 7
  shape_id=11  kind=textbox  box=(60,200,840,40)
    p0: runs=[ r0(bold)="Zelboraf " , r1(plain)="— a BRAF inhibitor for metastatic melanoma" ]
```
*VP prompt:* `On slide 7, in shape 11, change the bold product name from "Zelboraf" to "Cotellic" but leave the description alone.`
*Actions:*
```json
{"actions":[
  {"type":"set_run_text","slide":7,"shape_id":11,"paragraph_index":0,"run_index":0,"value":"Cotellic "}
]}
```

### 3.2 Color and underline one run; add a hyperlink to another

*VP prompt:* `On slide 7, shape 11: make run 0 of paragraph 0 dark blue (#1F3864) and underlined, and add a hyperlink on run 1 to https://example.com/cotellic.`
*Actions:*
```json
{"actions":[
  {"type":"set_run_font_color","slide":7,"shape_id":11,"paragraph_index":0,"run_index":0,"value":"#1F3864"},
  {"type":"set_run_underline","slide":7,"shape_id":11,"paragraph_index":0,"run_index":0,"value":true},
  {"type":"set_run_hyperlink","slide":7,"shape_id":11,"paragraph_index":0,"run_index":1,"value":"https://example.com/cotellic"}
]}
```

---

## 4. Find / replace

### 4.1 Deck-wide company rename

*VP prompt:* `Replace "Globex Corporation" with "Initech Holdings" everywhere in the deck, and also "Globex" with "Initech".`
*Actions:*
```json
{"actions":[
  {"type":"find_replace_text","scope":"deck","find":"Globex Corporation","replace":"Initech Holdings"},
  {"type":"find_replace_text","scope":"deck","find":"Globex","replace":"Initech"}
]}
```

### 4.2 Slide-scoped replace + regex (years)

*VP prompt:* `On slide 8 only, change "FY23" to "FY24". Then deck-wide, bump every "FY2X" forward by one year using a regex if you can.`
*Actions:*
```json
{"actions":[
  {"type":"find_replace_text","scope":"slide:8","find":"FY23","replace":"FY24"},
  {"type":"find_replace_regex","scope":"deck","pattern":"FY2([0-9])","replacement":"FY2$1+1"}
]}
```
> Caveat: regex *replacement* is literal substitution (no arithmetic). For a true
> year bump, enumerate: `find_replace_text "FY24"→"FY25"`, `"FY23"→"FY24"`, etc.,
> ordered from newest year to oldest to avoid double-bumping.

---

## 5. Add shapes, text boxes, lines

### 5.1 Add a navy "key takeaway" banner

*VP prompt:* `On slide 6, add a navy (#15283C) rounded-rectangle banner at left 60, top 470, width 840, height 50, with white centered 18pt text "Key takeaway: the deal is accretive in year 1." Name it banner_takeaway.`
*Actions:*
```json
{"actions":[
  {"type":"add_shape","slide":6,"kind":"rrect","pos":{"left":60,"top":470,"width":840,"height":50},"fill":"#15283C","text":"Key takeaway: the deal is accretive in year 1.","font_color":"#FFFFFF","font_size":18,"font_bold":true,"h_align":"center","v_align":"middle","ref_name":"banner_takeaway"}
]}
```

### 5.2 Add a plain text box and a horizontal divider line

*VP prompt:* `On slide 2, add a 12pt gray (#888888) text box "Source: company filings, FactSet" at left 48, top 510, width 400, height 20, and a 1pt gray horizontal line just above it from (48,505) to (912,505).`
*Actions:*
```json
{"actions":[
  {"type":"add_text_box","slide":2,"text":"Source: company filings, FactSet","pos":{"left":48,"top":510,"width":400,"height":20},"font_size":12,"font_color":"#888888"},
  {"type":"add_line","slide":2,"x1":48,"y1":505,"x2":912,"y2":505,"color":"#888888","weight_pt":1}
]}
```

### 5.3 Morph a rectangle into a chevron

*Snapshot excerpt:* `Slide 9  shape_id=14  kind=autoshape(rect)  text: "Phase 2"`
*VP prompt:* `On slide 9, change shape 14 from a rectangle into a chevron.`
*Actions:*
```json
{"actions":[
  {"type":"set_shape_kind","slide":9,"shape_id":14,"kind":"chevron"}
]}
```

---

## 6. Layout & alignment

### 6.1 Align and distribute three KPI cards

*Snapshot excerpt:* `Slide 4 has three textboxes: shape_id=20, 21, 22 (the KPI cards).`
*VP prompt:* `On slide 4, align the three KPI cards (shapes 20, 21, 22) to the same top edge and distribute them evenly across the slide.`
*Actions:*
```json
{"actions":[
  {"type":"align_shapes","slide":4,"shape_ids":[20,21,22],"anchor":"top"},
  {"type":"distribute_horizontal","slide":4,"shape_ids":[20,21,22]}
]}
```

### 6.2 Tile six logos into a 3-column grid; make them uniform size

*VP prompt:* `On slide 10, lay out the six logo pictures (shapes 30–35) in a 3-column grid with 24pt gaps, and set them all to 120 × 60.`
*Actions:*
```json
{"actions":[
  {"type":"uniform_size","slide":10,"shape_ids":[30,31,32,33,34,35],"width_pt":120,"height_pt":60},
  {"type":"tile_grid","slide":10,"shape_ids":[30,31,32,33,34,35],"cols":3,"gap_pt":24}
]}
```

### 6.3 Match size to a reference, then center on slide

*VP prompt:* `On slide 11, make shapes 41 and 42 the same size as shape 40, then center shape 40 horizontally on the slide.`
*Actions:*
```json
{"actions":[
  {"type":"match_size","slide":11,"ref_shape_id":40,"target_shape_ids":[41,42]},
  {"type":"align_to_slide_center","slide":11,"shape_id":40,"axis":"h"}
]}
```

### 6.4 Swap two shapes; nudge one

*VP prompt:* `On slide 12, swap the positions of shapes 50 and 51, then nudge shape 52 right by 18pt.`
*Actions:*
```json
{"actions":[
  {"type":"swap_positions","slide":12,"shape_a_id":50,"shape_b_id":51},
  {"type":"nudge","slide":12,"shape_id":52,"direction":"r","amount_pt":18}
]}
```

### 6.5 Clear a slide but keep the title

*Snapshot excerpt:* `Slide 13  shape_id=2 = title placeholder; everything else is clutter.`
*VP prompt:* `On slide 13, delete everything except the title (shape 2).`
*Actions:*
```json
{"actions":[
  {"type":"clear_slide","slide":13,"keep_shape_ids":[2]}
]}
```

### 6.6 Re-flow a slide into a two-column layout

Converting a layout = a batch of `move_shape` + `resize_shape`, one pair per shape. Target each shape by `shape_name` if it has a `ref_name`, otherwise by numeric `shape_id` (icons / connectors usually need the id). Don't touch text or colors unless asked. Finish with `enable_text_shrink_for_overflow` so anything that got narrower auto-shrinks.

*Snapshot excerpt:* `Slide 1, 960×540pt. Top band (financial): hdr_fin, chart_rev, chart_nps, kpi_table. Bottom band (execution): hdr_strat, box_p1/p2/p3, flow_1/flow_2/flow_3, badge_okr, icons id 17/20/23/26 + captions ic_cap1..4. Plus title (id 2), divider (id 3), footnote.`
*VP prompt:* `Re-flow slide 1 into two columns. Put the financial content — the two charts, the KPI table, and its header — into the left 66% as a vertical stack. Take the execution content — the three priority boxes, the Discover/Build/Scale chevrons, the OKR badge, and the four icon callouts with their captions, plus the "Strategic Priorities" header — and re-stack it vertically into the right 33% as a narrow column. Keep the title across the top and the footnote across the bottom. Add a thin vertical rule between the columns. Don't recreate the slide or change any text or colors.`
*Actions (excerpt — the full batch is ~48 actions, two per shape):*
```json
{"actions":[
  {"type":"move_shape","slide":1,"shape_id":"title","left":40,"top":16},
  {"type":"resize_shape","slide":1,"shape_id":"title","width":880,"height":34},

  {"type":"move_shape","slide":1,"shape_id":"chart_rev","left":40,"top":82},
  {"type":"resize_shape","slide":1,"shape_id":"chart_rev","width":560,"height":206},
  {"type":"move_shape","slide":1,"shape_id":"chart_nps","left":40,"top":296},
  {"type":"resize_shape","slide":1,"shape_id":"chart_nps","width":272,"height":118},
  {"type":"move_shape","slide":1,"shape_id":"kpi_table","left":328,"top":296},
  {"type":"resize_shape","slide":1,"shape_id":"kpi_table","width":272,"height":118},

  {"type":"move_shape","slide":1,"shape_id":"box_p1","left":636,"top":82},
  {"type":"resize_shape","slide":1,"shape_id":"box_p1","width":284,"height":52},
  {"type":"move_shape","slide":1,"shape_id":"flow_1","left":636,"top":258},
  {"type":"resize_shape","slide":1,"shape_id":"flow_1","width":284,"height":26},
  {"type":"move_shape","slide":1,"shape_id":17,"left":636,"top":392},
  {"type":"resize_shape","slide":1,"shape_id":17,"width":22,"height":22},
  {"type":"move_shape","slide":1,"shape_id":"ic_cap1","left":664,"top":392},
  {"type":"resize_shape","slide":1,"shape_id":"ic_cap1","width":256,"height":22},

  {"type":"add_line","slide":1,"x1":624,"y1":60,"x2":624,"y2":516,"color":"#BFBFBF","weight_pt":1.0},
  {"type":"enable_text_shrink_for_overflow","scope":"slide:1"}
]}
```
Note: tables auto-grow their row heights, so a table may end a few pt taller than the `height` you set — leave slack below it.

### 6.7 Explode a bullet box into visual cards

The most common "make it visual" ask. Read the box's `paragraphs` out of the snapshot, delete the box (keep the title), and add one card per bullet positioned per a region preset chosen by bullet count (see `LAYOUT_RECIPES.md`). Vary the card fills for rhythm; finish with `enable_text_shrink_for_overflow`.

*Snapshot excerpt:* `Slide 1: shape_id=2 = title "Our FY26 Operating Plan"; shape_id=3 = text box, 4 paragraphs: ["Accelerate enterprise ARR — focused land-and-expand in the top 200 named accounts", "Ship the AI copilot to GA and drive attach across the installed base", "Expand gross margin +200 bps via the multi-cloud cost-optimization program", "Deepen the partner channel — ecosystem pipeline reaches one third of new bookings"]`
*VP prompt:* `On slide 1, take those four bullets out of the text box and lay them out as a 2x2 grid of rounded cards, dark to light, numbered. Keep the title.`
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
For 3 bullets → 1 big left card + 2 stacked right cards; sequential steps → `add_shape kind:"chevron"` in a left-to-right row; 6 bullets → a 2×3 grid. See `LAYOUT_RECIPES.md` for the region rectangles.

---

## 7. Connectors, groups, diagrams

### 7.1 Three-step process flow (new shapes + connectors via ref_name)

*VP prompt:* `On slide 14, build a left-to-right 3-step flow: navy (#15283C) rounded rectangles "Originate", "Underwrite", "Fund", white 20pt bold text, at top 230, each 240 wide / 90 tall, at left 60 / 360 / 660, with navy elbow arrows between them.`
*Actions:*
```json
{"actions":[
  {"type":"add_shape","slide":14,"kind":"rrect","pos":{"left":60,"top":230,"width":240,"height":90},"fill":"#15283C","text":"Originate","font_color":"#FFFFFF","font_size":20,"font_bold":true,"ref_name":"s1"},
  {"type":"add_shape","slide":14,"kind":"rrect","pos":{"left":360,"top":230,"width":240,"height":90},"fill":"#15283C","text":"Underwrite","font_color":"#FFFFFF","font_size":20,"font_bold":true,"ref_name":"s2"},
  {"type":"add_shape","slide":14,"kind":"rrect","pos":{"left":660,"top":230,"width":240,"height":90},"fill":"#15283C","text":"Fund","font_color":"#FFFFFF","font_size":20,"font_bold":true,"ref_name":"s3"},
  {"type":"add_connector","slide":14,"kind":"elbow","from_shape_name":"s1","to_shape_name":"s2","arrow_end":"filled","color":"#15283C","weight_pt":2},
  {"type":"add_connector","slide":14,"kind":"elbow","from_shape_name":"s2","to_shape_name":"s3","arrow_end":"filled","color":"#15283C","weight_pt":2}
]}
```

### 7.2 Simple org chart (CEO over two VPs)

*VP prompt:* `On slide 15, build an org chart: a "CEO" box centered at top (left 380, top 80, 200×60), two boxes below — "VP Sales" (left 200, top 220, 200×60) and "VP Eng" (left 560, top 220, 200×60) — all light gray (#E7E7E7) with black 16pt text, straight connectors from CEO to each VP.`
*Actions:*
```json
{"actions":[
  {"type":"add_shape","slide":15,"kind":"rect","pos":{"left":380,"top":80,"width":200,"height":60},"fill":"#E7E7E7","text":"CEO","font_color":"#000000","font_size":16,"ref_name":"ceo"},
  {"type":"add_shape","slide":15,"kind":"rect","pos":{"left":200,"top":220,"width":200,"height":60},"fill":"#E7E7E7","text":"VP Sales","font_color":"#000000","font_size":16,"ref_name":"vp1"},
  {"type":"add_shape","slide":15,"kind":"rect","pos":{"left":560,"top":220,"width":200,"height":60},"fill":"#E7E7E7","text":"VP Eng","font_color":"#000000","font_size":16,"ref_name":"vp2"},
  {"type":"add_connector","slide":15,"kind":"straight","from_shape_name":"ceo","to_shape_name":"vp1","arrow_end":"none","color":"#888888","weight_pt":1.5},
  {"type":"add_connector","slide":15,"kind":"straight","from_shape_name":"ceo","to_shape_name":"vp2","arrow_end":"none","color":"#888888","weight_pt":1.5},
  {"type":"group_shapes","slide":15,"shape_ids":["ceo","vp1","vp2"],"ref_name":"org_chart"}
]}
```

---

## 8. Tables

### 8.1 Build a 4×3 financials table from scratch

*VP prompt:* `On slide 16, add a 4-row × 3-col table at left 80, top 130, width 800, height 220. Headers: "Metric", "FY23", "FY24E". Rows: Revenue / $1,210M / $1,360M; EBITDA / $218M / $272M; Net income / $96M / $128M. Apply a clean medium style and bold the header row.`
*Actions:*
```json
{"actions":[
  {"type":"add_table","slide":16,"rows":4,"cols":3,"pos":{"left":80,"top":130,"width":800,"height":220},"ref_name":"fin_table"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":1,"col":1,"value":"Metric"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":1,"col":2,"value":"FY23"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":1,"col":3,"value":"FY24E"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":2,"col":1,"value":"Revenue"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":2,"col":2,"value":"$1,210M"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":2,"col":3,"value":"$1,360M"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":3,"col":1,"value":"EBITDA"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":3,"col":2,"value":"$218M"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":3,"col":3,"value":"$272M"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":4,"col":1,"value":"Net income"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":4,"col":2,"value":"$96M"},
  {"type":"set_cell_text","slide":16,"shape_id":"fin_table","row":4,"col":3,"value":"$128M"},
  {"type":"apply_table_style","slide":16,"shape_id":"fin_table","style_id":"MediumStyle2Accent1"},
  {"type":"set_cell_text_align","slide":16,"shape_id":"fin_table","row":1,"col":1,"h_align":"left"},
  {"type":"set_cell_fill","slide":16,"shape_id":"fin_table","row":1,"col":1,"color":"#15283C"},
  {"type":"set_cell_fill","slide":16,"shape_id":"fin_table","row":1,"col":2,"color":"#15283C"},
  {"type":"set_cell_fill","slide":16,"shape_id":"fin_table","row":1,"col":3,"color":"#15283C"}
]}
```
> Style IDs are Office table-style names/GUIDs; `"NoStyleNoGrid"` and
> `"MediumStyle2Accent1"` are common safe choices.

### 8.2 Edit an existing table: fix a cell, add a row, widen a column

*Snapshot excerpt:* `Slide 16  shape_id=7  kind=table  rows=4 cols=3 (the financials table).`
*VP prompt:* `On slide 16 in the table (shape 7): change cell (3,2) to "$225M", add a new row after row 4 with "EPS / $1.40 / $1.86", and set column 1 width to 280pt.`
*Actions:*
```json
{"actions":[
  {"type":"set_cell_text","slide":16,"shape_id":7,"row":3,"col":2,"value":"$225M"},
  {"type":"add_table_row","slide":16,"shape_id":7,"after_row":4},
  {"type":"set_cell_text","slide":16,"shape_id":7,"row":5,"col":1,"value":"EPS"},
  {"type":"set_cell_text","slide":16,"shape_id":7,"row":5,"col":2,"value":"$1.40"},
  {"type":"set_cell_text","slide":16,"shape_id":7,"row":5,"col":3,"value":"$1.86"},
  {"type":"set_table_col_width","slide":16,"shape_id":7,"col":1,"width_pt":280}
]}
```

### 8.3 Convert a table into a chart (and back)

There's no single "table→chart" action — the model reads the table's cells out of the snapshot, deletes the table, and adds a chart with that data in the same `pos`. Give the new chart a `ref_name` so the styling follow-ups can target it.

*Snapshot excerpt:* `Slide 1  shape_id=7  kind=table  rows=4 cols=3  cells: [["Metric","FY24","FY25"],["Gross %","61.2%","63.8%"],["FCF ($M)","88","121"],["Headcount","2,140","2,387"]]  pos: L328 T296 W272 H138`
*VP prompt:* `On slide 1, turn that KPI table into a small clustered-column chart in the same spot — categories are the metric names, two series FY24 and FY25. Headcount is on a totally different scale from the % and the $ figure, so leave it out; just chart Gross Margin % and FCF ($M). Title it "FY24 vs FY25 — Key Financials", legend at the bottom, show data labels, hide the gridlines, keep it compact.`
*Actions:*
```json
{"actions":[
  {"type":"delete_shape","slide":1,"shape_id":7},
  {"type":"add_chart","slide":1,"chart_type":"columnclustered","pos":{"left":328,"top":296,"width":272,"height":118},
     "categories":["Gross Margin %","FCF ($M)"],
     "series":[{"name":"FY24","values":[61.2,88],"color":"#1F3864"},
               {"name":"FY25","values":[63.8,121],"color":"#9DC3E6"}],
     "title":"FY24 vs FY25 — Key Financials","show_legend":true,"clean_style":true,"ref_name":"kpi_chart"},
  {"type":"set_chart_legend","slide":1,"shape_id":"kpi_chart","props":{"position":"bottom","font_size":8}},
  {"type":"set_chart_series","slide":1,"shape_id":"kpi_chart","series_index":1,"props":{"show_labels":true,"label_format":"#,##0.0","label_size":8}},
  {"type":"set_chart_series","slide":1,"shape_id":"kpi_chart","series_index":2,"props":{"show_labels":true,"label_format":"#,##0.0","label_size":8}},
  {"type":"set_chart_gridlines","slide":1,"shape_id":"kpi_chart","props":{"major":false}},
  {"type":"set_chart_format","slide":1,"shape_id":"kpi_chart","props":{"gap_width":80}},
  {"type":"enable_text_shrink_for_overflow","scope":"slide:1"}
]}
```
Reverse direction (chart → table): `delete_shape` the chart, then `add_table` + `set_cell_text` per cell in the same `pos`. Note: `clean_style:true` on `add_chart` already hides the value axis + gridlines; the explicit `set_chart_gridlines` here is belt-and-suspenders. If the auto title wraps to two lines in a small chart, add `set_chart_title` with a smaller `font_size` or drop the title.

---

## 9. Charts (native chart objects)

### 9.1 Clustered column — revenue & EBITDA by year

*VP prompt:* `On slide 17, add a clustered column chart at left 80, top 110, width 800, height 360: categories FY21, FY22, FY23, FY24; series "Revenue ($M)" 980, 1110, 1210, 1360; series "EBITDA ($M)" 150, 188, 218, 272; title "Revenue & EBITDA"; legend on; label values as "$#,##0M".`
*Actions:*
```json
{"actions":[
  {"type":"add_chart","slide":17,"chart_type":"columnclustered","pos":{"left":80,"top":110,"width":800,"height":360},"categories":["FY21","FY22","FY23","FY24"],"series":[{"name":"Revenue ($M)","values":[980,1110,1210,1360]},{"name":"EBITDA ($M)","values":[150,188,218,272]}],"title":"Revenue & EBITDA","show_legend":true,"show_values":true,"value_format":"$#,##0\"M\""}
]}
```

### 9.2 Line chart with markers — a clean minimalist look

*VP prompt:* `On slide 18, add a line-with-markers chart at left 100, top 130, width 760, height 320 of monthly active users: Jan 1.2, Feb 1.35, Mar 1.5, Apr 1.72, May 1.9, Jun 2.1 (millions); one series "MAU (M)"; title "Monthly Active Users"; no legend; minimalist style; values as "0.0".`
*Actions:*
```json
{"actions":[
  {"type":"add_chart","slide":18,"chart_type":"linemarkers","pos":{"left":100,"top":130,"width":760,"height":320},"categories":["Jan","Feb","Mar","Apr","May","Jun"],"series":[{"name":"MAU (M)","values":[1.2,1.35,1.5,1.72,1.9,2.1]}],"title":"Monthly Active Users","show_legend":false,"clean_style":true,"show_values":true,"value_format":"0.0"}
]}
```

### 9.3 Pie chart — revenue mix, custom slice colors

*VP prompt:* `On slide 19, add a pie chart at left 280, top 120, width 400, height 360: categories Subscription, Services, Hardware; values 62, 23, 15 (percent); one series "Revenue mix"; slice colors #1F4E79, #2E75B6, #9DC3E6; title "Revenue Mix (%)"; legend on the right.`
*Actions:*
```json
{"actions":[
  {"type":"add_chart","slide":19,"chart_type":"pie","pos":{"left":280,"top":120,"width":400,"height":360},"categories":["Subscription","Services","Hardware"],"series":[{"name":"Revenue mix","values":[62,23,15]}],"title":"Revenue Mix (%)","show_legend":true}
]}
```
> Per-slice colors on a single-series pie aren't expressible via the `series[].color`
> field (that colors the whole series). To color slices, follow up with
> `set_chart_format` `vary_by_categories: true`, then ask the VP to fine-tune, or
> add `set_series_color` per index in a second pass if the chart engine supports
> point coloring in your build.

### 9.4 100% stacked bar — segment share over time

*VP prompt:* `On slide 20, add a 100% stacked bar chart at left 80, top 120, width 800, height 340: categories FY22, FY23, FY24; series "North America" 48, 46, 44; "EMEA" 31, 32, 33; "APAC" 21, 22, 23; title "Revenue by Region (% of total)"; legend on top.`
*Actions:*
```json
{"actions":[
  {"type":"add_chart","slide":20,"chart_type":"barstackedpercent","pos":{"left":80,"top":120,"width":800,"height":340},"categories":["FY22","FY23","FY24"],"series":[{"name":"North America","values":[48,46,44]},{"name":"EMEA","values":[31,32,33]},{"name":"APAC","values":[21,22,23]}],"title":"Revenue by Region (% of total)","show_legend":true},
  {"type":"set_chart_legend_position","slide":20,"shape_id":"<the new chart's id once known>","value":"top"}
]}
```
> If you can't pre-set legend position in `add_chart`, do it as a separate
> `set_chart_legend_position` after you know the chart's `shape_id` — or rely on
> the chart-type default. (The example shows the pattern; a real run would split
> these into two prompts: add the chart, re-snapshot, then position the legend.)

### 9.5 Modern chart type — waterfall (placeholder-data caveat)

*VP prompt:* `On slide 21, add a waterfall chart at left 100, top 120, width 760, height 340 titled "Bridge: FY23 → FY24 EBITDA".`
*Actions:*
```json
{"actions":[
  {"type":"add_chart","slide":21,"chart_type":"waterfall","pos":{"left":100,"top":120,"width":760,"height":340},"categories":["FY23","Price","Volume","Cost","FY24"],"series":[{"name":"EBITDA bridge","values":[218,18,22,-6,272]}],"title":"Bridge: FY23 → FY24 EBITDA"}
]}
```
> ⚠️ Waterfall, pareto, funnel, histogram, boxwhisker, treemap, and sunburst are
> created with the **right chart type but PowerPoint's placeholder data** — your
> `categories`/`series`/`title` won't be applied (a PowerPoint automation
> limitation). Tell the VP: "the chart frame is in; double-click it and paste your
> data." Include the data in the prompt anyway so the VP has it handy.

### 9.6 Edit an existing chart: change type, recolor series, axis titles, legend

*Snapshot excerpt:* `Slide 17  shape_id=8  kind=chart  type=barClustered  series: ["Revenue","Net Income"]  categories: ["2021","2022","2023","2024"]`
*VP prompt:* `On slide 17, the chart (shape 8): switch it to a clustered column chart, set series 1 color to #1F4E79 and series 2 to #C55A11, set the x-axis title to "Fiscal Year" and y-axis title to "USD Millions", and move the legend to the bottom.`
*Actions:*
```json
{"actions":[
  {"type":"set_chart_type","slide":17,"shape_id":8,"value":"columnclustered"},
  {"type":"set_series_color","slide":17,"shape_id":8,"series_index":1,"value":"#1F4E79"},
  {"type":"set_series_color","slide":17,"shape_id":8,"series_index":2,"value":"#C55A11"},
  {"type":"set_chart_axis_title","slide":17,"shape_id":8,"axis":"category","value":"Fiscal Year"},
  {"type":"set_chart_axis_title","slide":17,"shape_id":8,"axis":"value","value":"USD Millions"},
  {"type":"set_chart_legend_position","slide":17,"shape_id":8,"value":"bottom"}
]}
```

### 9.7b Clean up an EBITDA-bridge chart: kill gridlines, hide vertical axis, format data labels, recolor specific bars

*Snapshot excerpt:* `Slide 1  shape_id=2  kind=chart  type=type_119 (waterfall)  series: ["Series1"]` — note: a waterfall is **one** series with N points; recolor *points*, not series.
*VP prompt:* `Slide 1, the bridge chart (shape 2): remove the horizontal gridlines, hide the vertical value axis, add a page title "EBITDA Bridge", format the data labels with a thousands separator and negatives in parentheses like (40), and make bars 1/5/8 black and bars 4/6 yellow. Resize the chart to fill the slide better. Don't recreate the slide.`
*Actions:*
```json
{"actions":[
  {"type":"add_text_box","slide":1,"text":"EBITDA Bridge","pos":{"left":40,"top":18,"width":600,"height":44},"font_size":24,"font_bold":true,"font_color":"#000000","h_align":"left","ref_name":"slide_title"},
  {"type":"move_shape","slide":1,"shape_id":2,"left":40,"top":74},
  {"type":"resize_shape","slide":1,"shape_id":2,"width":880,"height":426},
  {"type":"set_chart_gridlines","slide":1,"shape_id":2,"props":{"major":false}},
  {"type":"set_chart_axis","slide":1,"shape_id":2,"axis":"y","props":{"visible":false}},
  {"type":"set_chart_series","slide":1,"shape_id":2,"series_index":1,"props":{"show_labels":true,"label_format":"#,##0;(#,##0)","point_fills":["#000000","","","#FFFF00","#000000","#FFFF00","","#000000"]}},
  {"type":"enable_text_shrink_for_overflow","scope":"slide:1"}
]}
```
Notes: (1) `set_chart_gridlines {major:false}` defaults to the value (y) axis — the horizontal lines. (2) Hiding the value axis (`set_chart_axis y {visible:false}`) also drops its gridlines, so on a normal column chart either action alone clears the horizontal lines. (3) Data-label format goes on `set_chart_series` (`label_format`), **not** the axis `number_format`. (4) Native waterfall (`type_119`) charts auto-color points by Increase/Decrease/Total — per-point `point_fills` usually sticks, but if it doesn't, set the colors in the PowerPoint UI (click a bar twice → Format Data Point → Fill).

### 9.7 Update existing chart data + add a trendline

*VP prompt:* `On slide 18, chart (shape 9): replace the categories with 2019, 2020, 2021, 2022, 2023; set series 1 ("Sales") to 410, 388, 502, 631, 770; rename series 1 to "Net Sales ($M)"; add a linear trendline on series 1 showing the R² value.`
*Actions:*
```json
{"actions":[
  {"type":"set_chart_categories","slide":18,"shape_id":9,"categories":["2019","2020","2021","2022","2023"]},
  {"type":"set_series_values","slide":18,"shape_id":9,"series_index":1,"values":[410,388,502,631,770]},
  {"type":"set_series_name","slide":18,"shape_id":9,"series_index":1,"value":"Net Sales ($M)"},
  {"type":"add_chart_trendline","slide":18,"shape_id":9,"series_index":1,"props":{"kind":"linear","display_r2":true}}
]}
```

---

## 10. Images & icons

### 10.1 Insert a local logo; replace an existing picture

*Snapshot excerpt:* `Slide 1  shape_id=3  kind=picture  box=(700,30,200,80)  (old logo)`
*VP prompt:* `On slide 1, replace the existing logo (shape 3) with C:\\brand\\acme_logo.png. Also add C:\\brand\\seal.png on slide 22 at left 820, top 460, width 100, height 100.`
*Actions:*
```json
{"actions":[
  {"type":"replace_picture","slide":1,"shape_id":3,"path":"C:\\brand\\acme_logo.png"},
  {"type":"insert_picture","slide":22,"path":"C:\\brand\\seal.png","pos":{"left":820,"top":460,"width":100,"height":100}}
]}
```

### 10.2 Insert Fluent UI icons

*VP prompt:* `On slide 12, add three filled navy (#15283C) Fluent icons, 48pt: a factory, a globe, and a group of people — at top 240, left 180 / 470 / 760, each 64×64.`
*Actions:*
```json
{"actions":[
  {"type":"insert_icon","slide":12,"icon":"building_factory","style":"filled","size":48,"color":"#15283C","left":180,"top":240,"width":64,"height":64,"ref_name":"ic_factory"},
  {"type":"insert_icon","slide":12,"icon":"globe","style":"filled","size":48,"color":"#15283C","left":470,"top":240,"width":64,"height":64,"ref_name":"ic_globe"},
  {"type":"insert_icon","slide":12,"icon":"people","style":"filled","size":48,"color":"#15283C","left":760,"top":240,"width":64,"height":64,"ref_name":"ic_people"}
]}
```
> Use Fluent UI icon names (`lowercase_underscore`, e.g. `chart_multiple`,
> `money`, `arrow_trending`, `shield`, `vehicle_car`). If the export prompt
> injected an allow-list, use only names from it; otherwise pick the closest
> semantic match.

### 10.3 Pull images from a website → picker grid → caption table

*VP prompt:* `Get all the images from https://acme.example.com/products into a local folder, build a thumbnail-grid picker slide from them (4 columns), and then build a 2-column image+caption table on slide 24 from the ones I'll pick.`
*Actions:*
```json
{"actions":[
  {"type":"fetch_page_images","url":"https://acme.example.com/products","dest_folder":"C:\\temp\\acme_products","ref_name":"acme_imgs"},
  {"type":"build_image_picker_slide","folder":"C:\\temp\\acme_products","cols":4,"insert_at":0,"max_per_slide":24}
]}
```
> `build_image_grid_table` needs the specific picked images + captions (it takes a
> `rows` array of `{image_path, text}` plus `slide` and `pos`) — that's a
> follow-up prompt once the VP has chosen images from the picker slide.

---

## 11. Deck-wide operations

### 11.1 Rebrand: colors + font in one shot

*VP prompt:* `Rebrand the deck: replace accent color #C00000 with #0B5394 everywhere (fills and fonts), replace #ED7D31 with #2E75B6, and change the body font from Calibri to Arial.`
*Actions:*
```json
{"actions":[
  {"type":"recolor_palette_deck_wide","from_hex":"#C00000","to_hex":"#0B5394","target":"both"},
  {"type":"recolor_palette_deck_wide","from_hex":"#ED7D31","to_hex":"#2E75B6","target":"both"},
  {"type":"swap_font_deck_wide","from_name":"Calibri","to_name":"Arial"}
]}
```

### 11.2 Theme fonts + slide size + layout sweep

*VP prompt:* `Set the heading theme font to "Georgia" and the body theme font to "Verdana", switch the deck to 16:9, and force the "Title and Content" layout (index 1) on slides 3, 4 and 5.`
*Actions:*
```json
{"actions":[
  {"type":"set_theme_font","major":"Georgia","minor":"Verdana"},
  {"type":"set_slide_size","preset":"16:9"},
  {"type":"apply_layout_to_slides","slide_indices":[3,4,5],"layout_index":1}
]}
```

### 11.3 Recolor just fills matching a color on one slide

*VP prompt:* `On slide 7 only, change every shape currently filled #D9D9D9 to #15283C.`
*Actions:*
```json
{"actions":[
  {"type":"recolor_fill_match","scope":"slide:7","from":"#D9D9D9","to":"#15283C"}
]}
```

### 11.4 Bulk-insert a confidentiality footer on several slides

*VP prompt:* `Add the text box "Confidential — for discussion purposes only" at left 48, top 515, width 500, height 18 on slides 2 through 8.`
*Actions:*
```json
{"actions":[
  {"type":"bulk_insert_text_box","slide_indices":[2,3,4,5,6,7,8],"text":"Confidential — for discussion purposes only","left":48,"top":515,"width":500,"height":18}
]}
```

---

## 12. Slide management

### 12.1 Add, duplicate, move, delete slides

*VP prompt:* `Add a blank slide after slide 5, duplicate slide 3, move slide 9 to position 2, and delete slide 12.`
*Actions:*
```json
{"actions":[
  {"type":"add_slide","position":6,"layout_index":6},
  {"type":"duplicate_slide","slide":3},
  {"type":"move_slide","from_slide":9,"to_slide":2},
  {"type":"delete_slide","slide":12}
]}
```
> Order matters: each of these renumbers subsequent slides. If you mean specific
> slides, do them one at a time and re-snapshot between, or reason carefully about
> the renumbering. Safest is to put deletions/moves last and from the bottom up.

### 12.2 Extract a subset to a new deck; import slides from another deck

*VP prompt:* `Export slides 3, 4 and 7 to C:\\out\\teaser.pptx, and import slides 1 and 2 from C:\\decks\\appendix.pptx into this deck at position 10.`
*Actions:*
```json
{"actions":[
  {"type":"extract_slides","slide_indices":[3,4,7],"output_path":"C:\\out\\teaser.pptx"},
  {"type":"import_slides_from_deck","source_path":"C:\\decks\\appendix.pptx","slide_indices":[1,2],"target_position":10}
]}
```

---

## 13. Speaker notes, backgrounds, slide numbers

### 13.1 Set / append speaker notes

*VP prompt:* `On slide 8, set the speaker notes to "Walk through the bridge waterfall; pause on the cost line." On slide 9, append "Q: expect a question on FX assumptions." to the existing notes.`
*Actions:*
```json
{"actions":[
  {"type":"set_speaker_notes","slide":8,"value":"Walk through the bridge waterfall; pause on the cost line."},
  {"type":"append_speaker_notes","slide":9,"value":"Q: expect a question on FX assumptions."}
]}
```

### 13.2 Slide background color + slide number

*VP prompt:* `On slide 1, set a dark navy (#0B1F33) background, and add a slide-number placeholder on slide 1 at left 900, top 515, width 40, height 20, white 10pt text.`
*Actions:*
```json
{"actions":[
  {"type":"set_slide_background_color","slide":1,"color":"#0B1F33"},
  {"type":"insert_slide_number","slide":1,"pos":{"left":900,"top":515,"width":40,"height":20},"font_color":"#FFFFFF","font_size":10}
]}
```

---

## 14. Visual effects

### 14.1 Shadow + glow on a hero shape

*Snapshot excerpt:* `Slide 2  shape_id=15  kind=autoshape(rrect)  (the hero card)`
*VP prompt:* `On slide 2, give shape 15 a soft drop shadow (offset 4,4, blur 8, color #000000, 60% transparent) and a subtle blue glow (color #2E75B6, radius 6, 40% transparent).`
*Actions:*
```json
{"actions":[
  {"type":"set_shadow","slide":2,"shape_id":15,"offset_x":4,"offset_y":4,"blur":8,"color":"#000000","transparency":0.6},
  {"type":"set_glow","slide":2,"shape_id":15,"color":"#2E75B6","radius":6,"transparency":0.4}
]}
```

### 14.2 Gradient fill, transparency, 3-D bevel

*VP prompt:* `On slide 2, fill shape 15 with a left-to-right gradient from #15283C to #2E75B6 at 0°, set it 15% transparent, and add a "circle" 3-D bevel 4pt deep.`
*Actions:*
```json
{"actions":[
  {"type":"set_gradient_fill","slide":2,"shape_id":15,"color1":"#15283C","color2":"#2E75B6","angle":0},
  {"type":"set_transparency","slide":2,"shape_id":15,"value":0.15},
  {"type":"set_3d_bevel","slide":2,"shape_id":15,"bevel_type":"circle","depth_pt":4}
]}
```

### 14.3 Picture: grayscale, crop, brightness

*Snapshot excerpt:* `Slide 22  shape_id=4  kind=picture  box=(60,120,840,360)`
*VP prompt:* `On slide 22, make picture 4 grayscale, crop 20pt off the left and right, and bump brightness +0.1.`
*Actions:*
```json
{"actions":[
  {"type":"recolor_picture","slide":22,"shape_id":4,"color_type":"grayscale"},
  {"type":"crop_picture","slide":22,"shape_id":4,"left":20,"right":20,"top":0,"bottom":0},
  {"type":"set_brightness","slide":22,"shape_id":4,"value":0.1}
]}
```

### 14.4 Line styling, rotation, flip, z-order

*VP prompt:* `On slide 9, give shape 14 a 2pt dashed #888888 outline, rotate it 90°, flip it horizontally, and send it to the back.`
*Actions:*
```json
{"actions":[
  {"type":"set_line_color","slide":9,"shape_id":14,"value":"#888888"},
  {"type":"set_line_weight","slide":9,"shape_id":14,"weight_pt":2},
  {"type":"set_line_style","slide":9,"shape_id":14,"style":"dash"},
  {"type":"rotate_shape","slide":9,"shape_id":14,"degrees":90},
  {"type":"flip_shape","slide":9,"shape_id":14,"axis":"h"},
  {"type":"z_order","slide":9,"shape_id":14,"order":"back"}
]}
```

---

## 15. Text-frame behaviour

### 15.1 Fix overflowing text deck-wide; vertical-center one shape

*VP prompt:* `A bunch of body text is overflowing its boxes — turn on shrink-to-fit for every text frame in the deck (leave titles alone). Also on slide 4, vertically center the text in shape 30.`
*Actions:*
```json
{"actions":[
  {"type":"enable_text_shrink_for_overflow","scope":"deck","include_titles":false},
  {"type":"set_text_vertical_align","slide":4,"shape_id":30,"value":"middle"}
]}
```

### 15.2 Internal margins; autofit mode; fit-to-content

*VP prompt:* `On slide 5, shape 4: set internal margins to 12pt all around, set autofit to "resize" (grow the box to fit text), and on slide 6, shape 9, shrink the box to its text content.`
*Actions:*
```json
{"actions":[
  {"type":"set_text_margin","slide":5,"shape_id":4,"left":12,"right":12,"top":12,"bottom":12},
  {"type":"set_text_autofit","slide":5,"shape_id":4,"mode":"resize"},
  {"type":"fit_to_content","slide":6,"shape_id":9}
]}
```

---

## 16. Composite: rebuild a slide for a target company

*VP prompt:* `Rebuild slide 4 as a one-page overview of "Acme Robotics" — keep the title (shape 2), but clear everything else. Then add: a subtitle "FY24 snapshot" under the title; three KPI cards in a row at top 150 (light gray #F2F2F2, 240×110, left 60/360/660) showing "$1.36B Revenue", "+18% YoY Growth", "$272M EBITDA" (bold 22pt navy headline + 12pt gray label inside each — just put the headline as the card text for now); and a clustered column revenue chart at left 60, top 290, width 840, height 220 with categories FY21–FY24 and series "Revenue ($M)" 980, 1110, 1210, 1360, title "Revenue", no legend, clean style.`
*Actions:*
```json
{"actions":[
  {"type":"clear_slide","slide":4,"keep_shape_ids":[2]},
  {"type":"add_text_box","slide":4,"text":"FY24 snapshot","pos":{"left":60,"top":110,"width":840,"height":30},"font_size":16,"font_color":"#888888","ref_name":"subtitle"},
  {"type":"add_shape","slide":4,"kind":"rrect","pos":{"left":60,"top":150,"width":240,"height":110},"fill":"#F2F2F2","text":"$1.36B Revenue","font_color":"#15283C","font_size":22,"font_bold":true,"h_align":"center","v_align":"middle","ref_name":"kpi1"},
  {"type":"add_shape","slide":4,"kind":"rrect","pos":{"left":360,"top":150,"width":240,"height":110},"fill":"#F2F2F2","text":"+18% YoY Growth","font_color":"#15283C","font_size":22,"font_bold":true,"h_align":"center","v_align":"middle","ref_name":"kpi2"},
  {"type":"add_shape","slide":4,"kind":"rrect","pos":{"left":660,"top":150,"width":240,"height":110},"fill":"#F2F2F2","text":"$272M EBITDA","font_color":"#15283C","font_size":22,"font_bold":true,"h_align":"center","v_align":"middle","ref_name":"kpi3"},
  {"type":"add_chart","slide":4,"chart_type":"columnclustered","pos":{"left":60,"top":290,"width":840,"height":220},"categories":["FY21","FY22","FY23","FY24"],"series":[{"name":"Revenue ($M)","values":[980,1110,1210,1360]}],"title":"Revenue","show_legend":false,"clean_style":true}
]}
```

---

## 17. Things to get right (recap for the model)

- **Wrap:** always `{"actions":[ ... ]}`, never a bare array. Output **only** the JSON.
- **`slide`** = 1-based. **`shape_id`** = the number from the snapshot (or a `ref_name` you assigned earlier in the same batch).
- **Units** = points; **colors** = `#RRGGBB`; **booleans** = `true`/`false`; **`pos`** = `{left,top,width,height}`.
- **Create before reference**; do **deletions/moves last** (and bottom-up) to avoid index/number drift.
- **Don't invent data.** **Don't guess fields** — if unsure, omit the action.
- **7 modern chart types** (waterfall, pareto, funnel, histogram, boxwhisker, treemap, sunburst) get the right type but placeholder data — say so to the VP.
- **Large batch?** Tell the VP to use the Execute window's **"Load from file..."** button, or emit **one action object per line**.

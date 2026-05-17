# Decko.ai — VP Prompting Guide

This guide shows how to tell the LLM what to do with your PowerPoint deck.
It covers every capability Decko currently supports, with real example prompts
you can paste or adapt.

**Two audiences:**
- **VPs / users** — read [§ Example Prompts](#example-prompts-by-use-case) and
  [§ Pro Tips](#pro-tips-for-better-prompts).
- **AI assistants** (Hermes, OpenClaw, GPT/Claude-class — any model) helping a
  VP — read [§ For AI Assistants](#for-ai-assistants-turning-a-vp-request-into-actions),
  the companion **[`ACTIONS_REFERENCE.md`](ACTIONS_REFERENCE.md)** (the
  exhaustive, machine-precise schema for all ~130 actions), and
  **[`EXAMPLES.md`](EXAMPLES.md)** (~45 paired examples: VP prompt → exact
  `actions` JSON, across every category — the best file to learn the mapping
  from). You do **not** need to have built Decko to use it correctly — follow
  the Hard Rules and the schema literally.

---

## How Decko Works (30-second recap)

1. Open your deck + `PPT_AI_Editor.pptm` in PowerPoint.
2. **Alt+F8 → ExportSnapshot** — opens the snapshot form.
3. Click **Copy snapshot + prompt template**. Your clipboard now has a full
   prompt block ending with `[REPLACE THIS LINE WITH YOUR REQUEST]`.
4. Paste into your LLM (Claude, ChatGPT, Gemini, Hermes, etc.).
5. Replace the placeholder line with your request. Hit send.
6. Copy the LLM's JSON response.
7. **Alt+F8 → ExecuteInstructions** — paste JSON, click **Parse** then **Apply**.
   For a large batch, click **"Load from file..."** instead of pasting (see the
   warning). Then **Apply**.
8. **Ctrl+S** to save. (Decko edits the open deck in memory — close without
   saving and you lose the changes.)

> ⚠️ **Large pastes get corrupted.** The Execute window's text box is a Windows
> MSForms control; pasting a big block (more than ~2 KB, or one very long line)
> can silently inject whitespace into numbers and keys (`"left":50` →
> `"left":5  0`), making the JSON invalid or making actions get skipped. For
> anything substantial, **save the LLM's JSON to a `.json` file and use the
> "Load from file..." button** — that path bypasses the text box entirely. If
> you must paste, ask the LLM to emit **one action per line**.

### What Decko can do (current surface)

~130 actions across 14 modules: text at whole-shape, paragraph, and run level;
layout & alignment; shapes, lines, connectors, groups; tables (build + cell
formatting); **native charts — all 39 PowerPoint chart types create real chart
objects, not images**; images (local insert, web scrape, picker grid); Fluent UI
icons; deck-wide font/color/theme swaps; slide add/delete/move/extract/import;
speaker notes; slide backgrounds & numbers; visual effects (shadow, glow,
gradient, 3-D bevel, picture recolor/crop/brightness). Full per-action schema:
**[`ACTIONS_REFERENCE.md`](ACTIONS_REFERENCE.md)**.

---

## Example Prompts by Use Case

The prompts below go in place of `[REPLACE THIS LINE WITH YOUR REQUEST]`.

---

### 1. Rebrand / Company Swap

**Swap one company for another across the whole deck:**
```
Replace all instances of "Amgen" with "Eli Lilly" and all instances of
"Amgen Inc." with "Eli Lilly and Company" throughout the entire deck.
Update the ticker symbol from "AMGN" to "LLY". Do not touch picture shapes
(the user will swap logos manually).
```

**Swap company on a single slide:**
```
On slide 5, update the company profile from Amgen to Pfizer. Replace the
company name, ticker, headquarters, CEO, employee count, and revenue with
the latest publicly available figures for Pfizer.
```

**Full pitch book rebrand:**
```
This deck was built for Amgen. Rebuild it for Eli Lilly. Replace every
mention of the company name, drug names, pipeline assets, financial figures,
and executive names with current Eli Lilly data. Use the latest public
information. Skip picture shapes — the user will replace logos manually.
```

---

### 2. Updating Text & Data

**Update a single data point:**
```
On slide 3, change the revenue figure from "$28.2B" to "$33.4B".
```

**Update multiple fields surgically:**
```
On slide 7:
- Change "CEO: Robert Bradway" to "CEO: John Smith"
- Change "Headquarters: Thousand Oaks, CA" to "Headquarters: Indianapolis, IN"
- Change "Employees: 24,900" to "Employees: 48,000"
Preserve all existing formatting.
```

**Find and replace a phrase across the deck:**
```
Find "Q3 2024" everywhere in the deck and replace it with "Q4 2024".
```

**Find and replace with regex (e.g., update all dollar figures):**
```
Find every figure in the format "$XX.XB" on slide 4 and replace:
- "$12.3B" → "$14.1B"
- "$8.7B" → "$9.2B"
```

**Update a bullet list on a slide (same number of bullets):**
```
On slide 6, update the three product bullet points. Keep the bold drug name
on each bullet but change the indications:
- Bullet 1: "Keytruda — Treatment of non-small cell lung cancer"
- Bullet 2: "Lynparza — Treatment of BRCA-mutated ovarian cancer"
- Bullet 3: "Imbruvica — Treatment of chronic lymphocytic leukemia"
```

**Add a new bullet point:**
```
On slide 9, add a new bullet after the last bullet in shape 12 that says:
"2024 acquisition of Morphic Therapeutic for $3.2B"
```

**Delete a bullet point:**
```
On slide 9, delete the second bullet point in the list.
```

---

### 3. Formatting Text

**Change font color for emphasis:**
```
On slide 2, change the font color of shape 5 (the headline) to #1F4E79.
```

**Make a paragraph bold:**
```
On slide 4, make paragraph 0 of shape 8 bold.
```

**Indent a bullet one level deeper:**
```
On slide 3, increase the indent level of paragraph 2 in shape 6 to level 2.
```

**Fix text overflow on a slide:**
```
Some text on slide 5 may be getting clipped. Enable auto-shrink for all
text shapes on slide 5 so nothing overflows.
```

**Fix text overflow across the entire deck:**
```
Enable auto-shrink for overflow on every text shape in the deck.
```

---

### 4. Building & Rebuilding Slides

**Clear a slide and start fresh:**
```
Slide 4 needs a full rebuild. Clear everything on it except shape 2 (the
header band). Then create:
- A large title text box at the top: "Key Investment Highlights"
- Three rectangular cards side by side in the middle of the slide, each
  with a headline and two supporting bullet points about Eli Lilly's growth
  drivers (pipeline strength, GLP-1 leadership, international expansion).
- Style the cards with fill color #1F4E79 and white text.
```

**Duplicate a slide and update it:**
```
Duplicate slide 3. On the new slide, change the company name from "Amgen"
to "Biogen" and update the financial figures with the latest Biogen data.
```

**Add a new blank slide:**
```
Add a new slide at position 5 using layout index 2 (blank layout).
```

---

### 5. Web Images & Product Grid Tables

Use this when you want a slide that shows product images + descriptions
scraped from a company's website.

**Step 1 — Scrape images from a URL:**
```
Fetch all images from https://www.stanley.com/industries and drop them onto
a new picker slide (4 images per row) so I can see which filenames
correspond to which industry photos.
```

*(After seeing the picker slide, you identify the image filenames and send
a second prompt:)*

**Step 2 — Build the product grid table:**
```
Build a 2-column image+description table on slide 6, using the images from
the folder the previous step downloaded. Use these rows:
- Row 1: name="Aerospace", image="img_003.jpg", bullets=["Lightweight fasteners","FAA compliant"]
- Row 2: name="Automotive", image="img_007.jpg", bullets=["High-volume assembly","Vibration-resistant"]
- Row 3: name="Construction", image="img_012.jpg", bullets=["Heavy-duty anchors","Code compliant"]
Position the table at left=30, top=60, width=900, height=480.
Name font: 12pt bold #15283C. Description font: 10pt #333333.
```

**Download a single image on the fly:**
```
Download the image at [URL] and save it to C:\Users\vinit\Downloads\logo.png,
then insert it on slide 1 at position left=820, top=20, width=80, height=40.
```

---

### 6. Icons (Microsoft Fluent UI)

The prompt template embeds a curated allow-list of **~664 IB-relevant Fluent UI
icon names** at export time. You do **not** need to look up icon names manually
— just describe the concept you want and the LLM picks the closest valid name
from the list. If no exact match exists (e.g., "oil barrel"), the LLM falls
back to the nearest semantic equivalent from the allow-list (e.g., `drop`).

Icon names are in `lowercase_underscore` format (e.g., `building_factory`).
The allow-list covers business, finance, technology, infrastructure, and
industry concepts used in IB pitch books.

**Insert a single icon:**
```
On slide 3, add a filled "factory" icon at
left=100, top=200, width=60, height=60, color #15283C.
```

**Add icons to a set of feature cards:**
```
On slide 5, add a Fluent UI icon to the top-left corner of each of the
three feature cards (shapes 8, 12, 16):
- Shape 8 (Growth): growth/trending icon, filled, 48px, color #1F4E79
- Shape 12 (Innovation): lightbulb icon, filled, 48px, color #1F4E79
- Shape 16 (Global): globe icon, filled, 48px, color #1F4E79
Place each icon at top=110 and left= the card's left edge + 10pt.
```

---

### 7. Shapes, Lines & Layout

**Add a shape:**
```
On slide 2, add a dark navy capsule shape (fill #15283C) at
left=50, top=480, width=180, height=32. Put the text "Q4 2024" in it,
white 11pt bold, centered.
```

**Add a divider line:**
```
On slide 3, add a horizontal line from x=30, y=95 to x=930, y=95,
color #15283C, weight 1.5pt.
```

**Align shapes:**
```
On slide 4, align shapes 5, 8, and 11 to their top edges.
```

**Distribute shapes evenly:**
```
On slide 4, distribute shapes 5, 8, and 11 horizontally with equal spacing.
```

**Arrange shapes into a grid:**
```
On slide 6, arrange shapes 3, 4, 5, 6, 7, 8 into a 3-column grid
with a 12pt gap between cells.
```

**Move a shape:**
```
On slide 2, move shape 7 to left=820, top=490.
```

**Nudge a shape:**
```
On slide 3, nudge shape 5 down by 20pt.
```

**Resize a shape:**
```
On slide 3, resize shape 9 to width=340, height=60.
```

**Delete a shape:**
```
Delete shape 14 on slide 2 — it's a leftover placeholder.
```

**Duplicate a shape:**
```
Duplicate shape 6 on slide 4 and place the copy at left=500, top=200.
```

**Change shape kind:**
```
Change shape 7 on slide 3 from a rectangle to a rounded rectangle (rrect).
```

**Z-order:**
```
Bring shape 9 on slide 5 to the front so it overlaps shape 7.
```

---

### 8. Org Charts & Process Flows

```
On slide 7, create a simple 3-level org chart:
- Top box: "CEO" (fill #1F4E79, white text)
- Two boxes below: "CFO" and "COO" (fill #2E75B6, white text)
Connect CEO → CFO and CEO → COO with elbow connectors (arrow at bottom,
entering at top of each subordinate box).
```

```
Build a 4-step process flow on slide 9:
Step 1: "Identify Target" → Step 2: "Due Diligence" →
Step 3: "Negotiate Terms" → Step 4: "Close Deal"
Use chevron shapes, fill #1F4E79, white bold text. Arrange left-to-right
with 10pt gaps. Add right-pointing arrows between each chevron.
```

---

### 9. Groups

**Group shapes:**
```
On slide 5, group shapes 8, 9, and 10 into a single object.
```

**Ungroup:**
```
Ungroup shape 15 on slide 5 so I can edit the individual parts.
```

---

### 10. Tables

**Create a new table:**
```
On slide 4, create a 5-row by 3-column table at left=50, top=100,
width=860, height=300. Label it "pipeline_tbl".
Set column widths: col 1 = 280pt, col 2 = 400pt, col 3 = 180pt.
Set header row height to 40pt.
```

**Fill a table with data:**
```
In the pipeline table on slide 4 (shape 8), fill in:
Row 1 (header): "Asset", "Indication", "Phase"
Row 2: "LY3437943", "Type 2 Diabetes", "Phase 3"
Row 3: "Tirzepatide", "Obesity", "Approved"
Row 4: "Mirikizumab", "Crohn's Disease", "Phase 3"
Row 5: "Lebrikizumab", "Atopic Dermatitis", "Approved"
```

**Style a table:**
```
Apply the "medium_style_2_accent1" built-in style to the table on slide 4
(shape 8). Then set the header row (row 1) fill color to #1F4E79 and
header text color to white.
```

**Update a single cell:**
```
In the table on slide 6 (shape 5), change row 3, column 2 to "Phase 2".
```

**Add/delete table rows or columns:**
```
In the table on slide 4 (shape 8), add a new row after row 4.
Then delete column 3 — we no longer need the Phase column.
```

**Swap columns or rows:**
```
In the table on slide 4 (shape 8), swap columns 1 and 2.
```

**Merge cells:**
```
In the table on slide 4 (shape 8), merge cells (1,1) through (1,3)
to create a full-width header spanning all columns.
```

---

### 11. Charts

Decko creates **real native PowerPoint chart objects** (editable, with an
embedded data sheet) — not pictures of charts. All **39** PowerPoint chart types
work: 2-D/3-D column & bar, line, area, pie/doughnut, scatter, radar, surface,
and the modern types (waterfall, pareto, funnel, histogram, box-and-whisker,
treemap, sunburst). See [`ACTIONS_REFERENCE.md` § chart types](ACTIONS_REFERENCE.md#chart-types-add_chart-chart_type-set_chart_type-value) for the full name list.

**Add a new chart:**
```
On slide 4, add a clustered column chart at left 60, top 120, width 560,
height 340:
- Categories: FY21, FY22, FY23, FY24
- Series "Revenue ($M)": 120, 138, 151, 170
- Series "EBITDA ($M)": 22, 28, 33, 41
- Title: "Revenue & EBITDA"
- Show the legend, format values as "$#,##0M".
```

**Change an existing chart's type:**
```
On slide 8, change the chart (shape 4) from clustered bar to clustered column.
```

**Update an existing chart's data:**
```
On slide 8, update the chart (shape 4):
- Categories: 2021, 2022, 2023, 2024
- Series "Revenue": 24.0, 26.8, 28.5, 33.4
- Series "Net Income": 5.9, 6.1, 7.0, 8.8
- Chart title: "Eli Lilly Financial Performance ($B)"
- X-axis title: "Year"; Y-axis title: "USD Billions"
```

**Recolor series / move legend:**
```
On slide 8, set series 1 color to #1F4E79 and series 2 to #2E75B6, and move
the legend to the bottom.
```

> **Heads-up on 7 chart types:** `waterfall`, `pareto`, `funnel`, `histogram`,
> `boxwhisker`, `treemap`, `sunburst` are created with the correct chart type
> but with PowerPoint's **placeholder data** — Decko cannot write your
> categories/series/title into them (a PowerPoint automation limitation). You
> double-click the chart and edit its data manually after insertion. The other
> 32 types take your data normally.

---

### 12. Visual Effects & Polish

**Add a drop shadow:**
```
On slide 3, add a soft drop shadow to shape 7:
offset X=3pt, offset Y=3pt, blur=8pt, color=#000000, transparency=0.6.
```

**Add a gradient fill:**
```
On slide 2, give shape 5 a gradient fill from #1F4E79 (left) to #2E75B6
(right), angle=0 degrees.
```

**Make a shape semi-transparent:**
```
On slide 4, set shape 9 to 30% transparent fill so the background shows
through.
```

**Add a glow effect:**
```
On slide 5, add a blue outer glow (color #2E75B6, radius=10pt,
transparency=0.4) to shape 6 to make it pop.
```

**Rotate a shape:**
```
Rotate shape 8 on slide 3 by 45 degrees.
```

**Recolor a picture (grayscale, sepia, etc.):**
```
On slide 1, make the background photo (shape 2) grayscale.
```

**Adjust image brightness/contrast:**
```
On slide 1, reduce the brightness of picture shape 2 to -0.3 and
increase contrast to 0.2 so the text overlay is more readable.
```

---

### 13. Deck-Wide Color & Font Operations

**Recolor all shapes of one color across the deck:**
```
Replace every shape with fill color #0070C0 with #1F4E79 across the
entire deck.
```

**Recolor all text of one color across the deck:**
```
Replace every text element colored #FF0000 with #15283C across the deck.
```

**Swap fonts across the deck:**
```
Replace all uses of "Calibri" with "Inter" throughout the entire deck.
```

**Change theme fonts:**
```
Set the heading (major) theme font to "Helvetica Neue" and the body
(minor) theme font to "Helvetica Neue" across the deck.
```

**Set slide dimensions:**
```
Change the deck to standard widescreen 16:9 (960pt × 540pt).
```

---

### 14. Slide Management

**Reorder slides:**
```
Move slide 8 to position 3 in the deck.
```

**Delete a slide:**
```
Delete slide 12 — it's a scratch slide we no longer need.
```

**Export selected slides to a new file:**
```
Extract slides 1, 3, 5, 7, and 9 into a new file at
C:\Users\vinit\Downloads\exec_summary.pptx.
```

**Import slides from another deck:**
```
Import slides 1 and 2 from C:\Users\vinit\Downloads\template.pptx
and insert them at position 3 in the current deck.
```

---

### 15. Speaker Notes

**Set notes on a slide:**
```
Set the speaker notes on slide 4 to:
"Key message: Eli Lilly's GLP-1 franchise (Mounjaro + Zepbound) now
accounts for 40% of revenue. Emphasize the durable competitive moat
from manufacturing scale and IP protection through 2030."
```

**Add to existing notes:**
```
Append to the speaker notes on slide 6:
"Q&A talking point: The $3.2B Morphic acquisition closes in Q1 2025
and adds an oral integrin inhibitor to the IBD pipeline."
```

---

### 16. Slide Background

**Change slide background color:**
```
Set the background of slide 1 to #15283C (dark navy).
```

---

### 17. Insert Slide Number

```
On slide 4, add a slide number placeholder at the bottom right:
left=880, top=515, width=60, height=20, font color #888888, size 9pt.
```

---

### 18. Copy Formatting Between Shapes

```
Copy all formatting (font, fill, border, effects) from shape 5 on slide 3
and apply it to shape 9 on the same slide.
```

---

## Pro Tips for Better Prompts

**Be specific about the slide.** Say "slide 4" or "the company profile slide"
rather than "the slide" — the LLM reads the snapshot and will find the right
one, but naming it avoids ambiguity.

**Don't specify shape IDs yourself.** The LLM reads shape IDs from the
snapshot. You can say "the revenue chart" or "the pipeline table" and it will
identify the right shape_id.

**Say "preserve formatting" when you want surgical edits.** This tells the LLM
to use `find_replace_text` or `set_run_text` instead of blunt `set_text`,
which strips all formatting.

**For full slide rebuilds, say so explicitly.** "Rebuild this slide from
scratch" lets the LLM use `clear_slide` first and then add shapes cleanly.

**Chain prompts for web images.** The web image workflow is a two-step
conversation: first ask for the scrape + picker slide, then (after you've
identified filenames) send a second prompt to build the final table.

**Always Ctrl+S after applying.** Decko edits the open deck in memory. Close
without saving = all changes lost. Ctrl+S right after every Apply.

**For table-anchored layouts, describe the anchor — don't guess coordinates.**
Say "position it above the third column of the table on slide 4" rather than
inventing pixel offsets. The model computes the exact position from the snapshot's
table geometry.

**For "add something similar to existing", describe the pattern, not the dimensions.**
Say "add a section above column 4 matching the style of the existing sections above
columns 2 and 3" — the model reads existing shapes' dimensions from the snapshot
and copies them. Don't say "width 200, height 44" unless you actually know.

**For compound geometric operations, describe the outcome, not the steps.**
Say "three header sections, each flush with its table column, same height, evenly
distributed above columns 2, 3, and 4." The model handles the inspect → compute
→ reposition sequence.

**Use relative spatial language.** "Above column N", "same top edge as the existing
sections", "matching the height of the boxes above column 2" — this is easier to
say and more reliable than specific point values.

---

## Full Action Reference (for AI Assistants)

This section is a quick map of every action Decko supports. For the **complete,
machine-precise schema** (every required/optional field, every value vocabulary,
a minimal example per action), use **[`ACTIONS_REFERENCE.md`](ACTIONS_REFERENCE.md)**
— that file is the single source of truth and is meant to be read literally by
any model.

### Action Categories & Counts (246 total, 17 modules)

| Module | Count | What it does |
|--------|-------|--------------|
| `modActionsTemplate.bas` | 1 | `apply_template` — build a whole slide from a named layout in one action |
| `modActionsSpec.bas` | 3 | Decks-as-code: `build_deck_from_spec`, `extract_spec`, `generate_variants` |
| `modActionsCapture.bas` | 4 | Deck DNA: `capture_template` / `list_templates` / `delete_template` / `rename_template` |
| `modActions.bas` | ~17 | Core shape/slide/notes/table-cell ops |
| `modActionsText.bas` | ~17 | Paragraph text, bullets, alignment, autofit, find/replace |
| `modActionsRun.bas` | ~11 | Run-level formatting (bold/italic/color/size/font/hyperlink/strikethrough) |
| `modActionsLayout.bas` | ~30 | Align, distribute, grid, spacing, match size/pos, add_shape, add_text_box, add_line, clear_slide, z_order, duplicate, recolor batch |
| `modActionsTable.bas` | ~13 | add_table, add/del row/col, merge, col/row size, cell border/fill/align, table style, image grid table |
| `modActionsChart.bas` | ~18 | add_chart (39 types), set type/title/axis/legend/series/categories/values/colors, trendlines, error bars, chart format |
| `modActionsImage.bas` | ~4 | Insert/replace picture, image picker grid |
| `modActionsConnector.bas` | 1 | add_connector |
| `modActionsGroup.bas` | 2 | group/ungroup |
| `modActionsSlide.bas` | ~6 | move/extract/import slides, slide background, slide number |
| `modActionsDeck.bas` | ~12 | Regex find/replace, font swap, theme, recolor palette, batch palette remap (`recolor_deck`), palette scanner (`scan_palette`), slide size, bulk insert, layout apply |
| `modActionsEffects.bas` | ~17 | Rotate/flip, line color/weight/style, shadow, glow, reflection, transparency, gradient, 3-D bevel, preset effect, crop/recolor/brightness/contrast picture, shape adjustment |
| `modActionsIcon.bas` | 1 | insert_icon (Fluent UI SVG) |
| `modActionsWeb.bas` | ~3 | fetch_page_images, download_image, open_image_picker |

### Action Quick-Reference

#### High-Level Authoring — TRY THESE FIRST

When the request is "make a slide that …" or "build a deck about …" (not
"move/recolor this existing box"), reach for these BEFORE placing
individual shapes. One action replaces 10–30 atomic ops and produces a
composed, on-brand slide.

| Action | When to suggest |
|--------|----------------|
| `apply_template` | "Make a title / section / bullets / two-column / comparison / KPI / quote slide." `template` is ONE literal name + its literal slots: `title`(`title,subtitle`), `section`(`section_number,section_title`), `bullets`(`heading,bullets[]`), `two_col`(`heading,left_body,right_body`), `comparison`(`heading,left_label,left_body,right_label,right_body`), `kpi_dashboard`(`heading,tiles[]`={stat,label}), `quote`(`quote_text,attribution`) — or any captured-template name. `content` holds those exact slots. `"slide":N` replaces; omit to append. **No `two_column`/`image_caption` — use names above verbatim.** |
| `build_deck_from_spec` | "Build me a 5-slide deck on X." Pass `spec` = array of `{"template":"…","content":{…}}` (one per slide). `"clear_existing":true` replaces the deck; omit to append. Whole deck in ONE action. |
| `extract_spec` | "Clone / restyle this deck." Reads the current deck back out as a `spec` you can edit and feed to `build_deck_from_spec`. Round-trips. |
| `generate_variants` | "Show me a few different layouts of this." Renders the content as N distinct **principled archetypes** — Hero, Split, Stack, Quote, Tiles (cycling). Form A: `template`+`n` (heading taken from `title`/`heading`/`section_title`/`quote_text`, rest → body). Form B: `templates:[names]` renders the SAME content across each named template (content must satisfy all their slots). NOT a position shuffle. |
| `capture_template` | "I like this slide, save it as a template." Requires `name`; optional `slide:N` (default active). Stored in the Deck DNA registry; thereafter usable as a `template` value in `apply_template`/`build_deck_from_spec`. |
| `list_templates` | "What templates do I have?" Returns the captured registry (names + slots). |
| `delete_template` | "Delete the X template." Requires `name`. |
| `rename_template` | "Rename template X to Y." Params **`from`** and **`to`** (NOT `name`/`new_name`). |

**Captured-template manifest:** the snapshot prompt you were given ends with
a live list of the user's captured Deck DNA templates and their content
slots. If the user references a template by name, target it directly — its
slots are listed there; you do not need to guess.

**DECK DESIGN PRINCIPLES:** the prompt also ends with an injected
design-principles block (visual hierarchy, contrast, alignment, restraint).
Treat it as binding: when you emit `apply_template`/`build_deck_from_spec`
content, follow those principles — don't over-stuff slots, keep one idea
per slide, respect contrast.

**Icons:** there is no exhaustive allow-list. To add an icon, source the
SVG by semantic name from the CDN per the guidance in the prompt and use
`insert_icon`. Prefer the short curated names listed; for anything else,
pick the closest semantic Fluent name.

#### Atomic Ops
| Action | When to suggest |
|--------|----------------|
| `set_text` | Replace all text in a shape (plain text only, single run). |
| `set_font_size` | Change font size for the whole shape. |
| `set_font_bold` | Bold/unbold the whole shape. |
| `set_font_italic` | Italic the whole shape. |
| `set_font_color` | Recolor all text in a shape. |
| `set_fill_color` | Change shape background color. |
| `move_shape` | Position a shape by absolute left/top. |
| `resize_shape` | Change width/height. |
| `delete_shape` | Remove a shape. |
| `add_slide` | Insert new slide at position. |
| `delete_slide` | Remove a slide. |
| `duplicate_slide` | Clone a slide. |
| `set_cell_text` | Update one table cell. |
| `swap_table_columns` | Swap two table columns. |
| `swap_table_rows` | Swap two table rows. |
| `set_speaker_notes` | Replace slide speaker notes. |
| `append_speaker_notes` | Add to existing speaker notes. |

#### Granular Text (paragraph-level)
| Action | When to suggest |
|--------|----------------|
| `set_paragraph_text` | Replace one paragraph (single-run paragraphs only). |
| `add_paragraph` | Add a bullet/paragraph. |
| `delete_paragraph` | Remove a bullet/paragraph. |
| `set_bullet_style` | none / disc / number. |
| `set_indent_level` | Indent depth 0–4. |
| `set_paragraph_font_size` | Per-paragraph size. |
| `set_paragraph_font_color` | Per-paragraph color. |
| `find_replace_text` | Safe find/replace, scope=deck or slide:N. |
| `set_paragraph_alignment` | left / center / right / justify. |
| `set_paragraph_line_spacing` | 1.0 / 1.5 / 2.0 etc. |
| `set_text_vertical_align` | top / middle / bottom (whole shape). |
| `set_text_margin` | Internal padding in pt. |
| `set_text_autofit` | none / shrink / resize (per shape). |
| `enable_text_shrink_for_overflow` | Auto-shrink all text on slide:N or deck. |

#### Run-Level Formatting (sub-paragraph precision)
| Action | When to suggest |
|--------|----------------|
| `set_run_text` | Change one run's text without touching formatting of siblings. Use this for mixed-format paragraphs (bold drug name + plain description). |
| `set_run_bold` | Bold one run. |
| `set_run_italic` | Italic one run. |
| `set_run_underline` | Underline one run. |
| `set_run_subscript` | Subscript. |
| `set_run_superscript` | Superscript. |
| `set_run_font_color` | Recolor one run. |
| `set_run_font_size` | Resize one run. |
| `set_run_font_name` | Different font on one run. |
| `set_run_hyperlink` | Add / clear hyperlink on a run. |

#### Layout & Composition
| Action | When to suggest |
|--------|----------------|
| `align_shapes` | Align left/right/top/bottom/hcenter/vcenter. |
| `distribute_horizontal` | Even horizontal gaps. |
| `distribute_vertical` | Even vertical gaps. |
| `tile_grid` | N-column grid with gap. |
| `fit_to_slide_margins` | Shrink to fit inside margin. |
| `add_line` | Horizontal/vertical divider. |
| `add_shape` | New autoshape (rect/oval/capsule/arrow/chevron/callout/star/…). |
| `set_shape_kind` | Morph existing shape to new kind. |
| `clear_slide` | Delete all shapes except keep list. |
| `move_shape_relative` | Nudge by dx/dy. |
| `snap_to_grid` | Round position to grid_pt. |
| `align_to_slide_center` | Center on slide horizontally/vertically/both. |
| `nudge` | Shift l/r/u/d by amount_pt. |
| `fit_to_content` | Auto-resize shape to its text. |
| `match_size` | Copy one shape's size to others. |
| `uniform_size` | Set all listed shapes to same size. |
| `smart_spacing` | Place each shape gap_pt from previous edge. |
| `equalize_spacing` | Equal gaps along axis. |
| `match_position` | Align target edge to reference edge. |
| `swap_positions` | Swap two shapes' positions and sizes. |
| `group_by_overlap` | Group shapes whose bboxes intersect. |
| `recolor_fill_match` | Replace fill color across scope. |
| `recolor_font_match` | Replace font color across scope. |
| `delete_shapes_match` | Delete shapes matching text/kind filter. |
| `duplicate_shape` | Clone shape at new position. |
| `copy_formatting` | Copy formatting from source to target. |
| `z_order` | front / back / forward one / back one. |

#### Tables
| Action | When |
|--------|------|
| `add_table` | Create new table. |
| `add_table_row` | Add row after row N. |
| `delete_table_row` | Remove row N. |
| `add_table_col` | Add column after col N. |
| `delete_table_col` | Remove column N. |
| `merge_cells` | Merge cell range. |
| `set_table_col_width` | Set column width in pt. |
| `set_table_row_height` | Set row height in pt. |
| `set_cell_border` | Border style on a cell side. |
| `set_cell_text_align` | h_align + v_align per cell. |
| `set_cell_fill` | Cell background color. |
| `apply_table_style` | Named Office table style. |

#### Charts (native charts only; pasted images skipped)
| Action | When |
|--------|------|
| `set_chart_type` | barClustered / columnClustered / line / pie / area / scatter / etc. |
| `set_chart_title` | Title text + enabled flag (+ font props). |
| `set_chart_axis_title` | x or y axis title. |
| `set_chart_axis` | Fine axis control: `visible`/`line_visible` (hide the axis — value-axis also drops its gridlines), `min`/`max`/`major_unit`, `number_format`, `tick_label_position`, scale type. |
| `set_chart_gridlines` | Show / hide / style major & minor gridlines per axis (`x`/`y`/`both`) — color, weight, dash. Use `{"props":{"major":false}}` to kill the horizontal lines. |
| `set_chart_legend_position` / `set_chart_legend` | top / right / bottom / left; or `props` for position + visible + font. |
| `set_series_color` | Color a *series* by 1-based index. (For per-bar color in a single-series chart use `set_chart_series` `props.point_fills`.) |
| `set_chart_series` | Per-series props: `show_labels`, `label_format` (the data-label number format — e.g. `#,##0;(#,##0)`), `label_position`, `fill_color`, `point_fills` (array, per-point color), `point_marker_styles` (per-point marker — pass `"none"` to hide one), `point_line_visible` (per-point line-segment hide — break a line before a sentinel last point), markers, axis_group, custom_labels, etc. |
| `set_chart_format` | Chart-group props: `gap_width`, `overlap`, `bar_shape`, `doughnut_hole_size`, reverse cats/series, **`plot_area_left/top/width/height`** (pin the plot rectangle inside the frame so caller-side overlay text — segment totals, CAGR callouts, growth brackets — lands on bars deterministically). |
| `set_series_values` | Array of data values. |
| `set_chart_categories` | Array of category labels. |
| `set_series_name` | Series label. |

#### Images
| Action | When |
|--------|------|
| `insert_picture` | Insert local image file at position. |
| `replace_picture` | Swap existing picture, keep frame. |
| `fetch_page_images` | Scrape all images from a URL. |
| `build_image_picker_slide` | Visual thumbnail grid from downloaded folder. |
| `build_image_grid_table` | 2-col image+desc table from row spec. |
| `download_image` | Download one URL to local path. |

#### Icons (Microsoft Fluent UI)
| Action | When |
|--------|------|
| `insert_icon` | Add a Fluent UI SVG icon. Describe the concept — the LLM picks from the ~664-icon IB allow-list in the prompt. |

Params: `slide`, `icon` (lowercase_underscore name from allow-list), `style` (filled/regular),
`size` (16/20/24/28/32/48), `color` (#RRGGBB), `left`, `top`, `width`, `height` (all in pt).

Note: the allow-list is injected at export time. LLM must use only names from that list;
if no exact match exists it picks the nearest semantic equivalent.

#### Visual Effects
| Action | When |
|--------|------|
| `rotate_shape` | Rotate by degrees. |
| `flip_shape` | h or v flip. |
| `set_line_color` | Outline color. |
| `set_line_weight` | Outline weight in pt. |
| `set_line_style` | solid / dash / dot / dashdot. |
| `set_shadow` | Drop shadow params. |
| `set_glow` | Outer glow. |
| `set_reflection` | Reflection effect. |
| `set_transparency` | Fill transparency 0.0–1.0. |
| `set_gradient_fill` | Two-color gradient with angle. |
| `set_3d_bevel` | 3D bevel style + depth. |
| `apply_preset_effect` | Office texture preset 1–24. |
| `crop_picture` | Crop edges in pt. |
| `recolor_picture` | grayscale / sepia / washout / bw / auto. |
| `set_brightness` | -1.0 to 1.0. |
| `set_contrast` | -1.0 to 1.0. |

#### Connectors & Groups
| Action | When |
|--------|------|
| `add_connector` | Elbow/straight/curved connector between shapes. |
| `group_shapes` | Group shape list into one. |
| `ungroup` | Ungroup a group shape. |

#### Slide Structure
| Action | When |
|--------|------|
| `move_slide` | Reorder slide A to position B. |
| `extract_slides` | Export selected slides to new .pptx. |
| `import_slides_from_deck` | Pull slides from another .pptx at position. |

#### Deck-Wide
| Action | When |
|--------|------|
| `find_replace_regex` | Regex find/replace across deck or slide. |
| `swap_font_deck_wide` | Replace one font name everywhere. |
| `recolor_palette_deck_wide` | Replace one color across fill/font/both. |
| `recolor_deck` | Batch palette remap — N color pairs in one pass. Covers fills, fonts, borders, tables, charts, slide backgrounds, groups. |
| `scan_palette` | Scan active deck for explicit RGB colors; writes role-tagged JSON to clipboard. Use before `recolor_deck` to discover what colors exist. |
| `apply_theme` | Apply .thmx or .potx theme file. |
| `set_slide_size` | Set slide dimensions (pt) or preset (16:9 / 4:3). |
| `set_theme_font` | Set major/minor theme fonts. |
| `bulk_insert_image` | Same image at same position on multiple slides. |
| `bulk_insert_text_box` | Same text box on multiple slides. |
| `apply_layout_to_slides` | Force layout index N on listed slides. |

#### Slide Background & Numbers
| Action | When |
|--------|------|
| `set_slide_background_color` | Solid color background. |
| `insert_slide_number` | Add slide number placeholder with font/pos. |
| `set_run_strikethrough` | Strikethrough on a run. |

---

## For AI Assistants — turning a VP request into actions

This section is for any model (Hermes, OpenClaw, GPT-class, Claude-class) acting
as the VP's assistant. **You have two jobs:**

1. **Help the VP phrase a request** — turn a vague business ask into a precise,
   unambiguous instruction (the "VP prompt") that a model can act on.
2. **Emit the `actions` JSON** — given the snapshot + the VP's request, output
   the exact `{"actions":[ ... ]}` array Decko will execute.

You can do both even though you didn't build Decko. Just follow the Hard Rules
and the schema in [`ACTIONS_REFERENCE.md`](ACTIONS_REFERENCE.md) **literally**.

### Hard Rules (memorize — breaking any one fails the batch)

1. Output is **one JSON object: `{"actions":[ ... ]}`**. Never a bare array.
2. Every item has a **`"type"`** plus that action's fields. Actions run **in order, top to bottom**.
3. **`slide` = 1-based slide number.**
4. **`shape_id` = the number from the snapshot** — never the shape's name, never its list position. (Exception: a shape created earlier in the same batch via `ref_name` may be referenced by that `ref_name` string.)
5. **All sizes/distances are points (pt).** 1 inch = 72 pt. 16:9 slide = 960 × 540 pt (the snapshot reports the actual size).
6. **Colors are `"#RRGGBB"`.** **Booleans are `true`/`false`.**
7. **`pos` is `{"left":N,"top":N,"width":N,"height":N}`** for `add_shape`, `add_text_box`, `add_chart`, `add_table`, `insert_picture`, `insert_slide_number`.
8. **Scope strings are `"deck"` or `"slide:N"`.**
9. **Create before you reference.** `add_slide` before charts on it; `add_table` before `set_cell_text` on it; `add_shape` (with `ref_name`) before `add_connector` to it.
10. **Don't guess.** Missing required field / unknown type / out-of-range ID → that action is skipped (batch continues). Malformed JSON → whole batch fails. When unsure, omit.
11. **Never invent data.** Facts come from the VP or public sources they approve.
12. **Snapshot is helpful but not a blocker.** If available, read shape IDs, positions, and text from it. If not available, proceed with best-effort assumptions (reasonable coordinates, placeholder text, inferred layout) and **state the assumptions inline** so the VP can correct them. Never stop and wait for a snapshot.
13. **Big batches: tell the VP to use "Load from file...", or emit one action per line.** The Execute text box corrupts large pastes.
14. **Geometry-sensitive operations need snapshot-derived coordinates, never guesses.** When the request involves aligning shapes to a table, matching existing shape dimensions, or redistributing objects — read positions from the snapshot. Compute `left`/`width` from the table's reported `box` and column count. See [Geometry-Sensitive Operations](#geometry-sensitive-operations) below.

### Step-by-step: VP request → actions JSON

1. **Read the snapshot.** Note slide count, slide size, and for each shape its
   `shape_id`, kind, position, and current text (with paragraph/run indices).
2. **Restate the request as concrete edits.** "Make it look cleaner" → which
   slides? which shapes? remove what / align what / recolor to what? If the VP
   was vague, ask one clarifying question rather than guessing.
3. **Build-a-slide / build-a-deck? Use High-Level Authoring first.** If the
   request is "make a slide that …" or "build a deck about …" (creating new
   content, not tweaking an existing box), emit ONE `apply_template` or
   `build_deck_from_spec` instead of 20 atomic shape ops. Check the
   captured-template manifest at the end of the prompt — if the user has a
   matching Deck DNA template, use its name. Only drop to atomic ops for
   edits the templates can't express. See
   [High-Level Authoring](#high-level-authoring--try-these-first).
4. **Pick the smallest action for each edit.** Prefer `find_replace_text` over
   `set_text`; `set_run_text` over `set_paragraph_text` when formatting is mixed.
   Use the [Quick-Reference](#action-quick-reference) and
   [`ACTIONS_REFERENCE.md`](ACTIONS_REFERENCE.md) to match intent → action.
5. **Order the actions.** Creates first, then edits/format on the created
   things, then layout/alignment last (alignment needs final sizes).
6. **Fill every required field; add optional fields only when the request asks
   for them.** Copy field names exactly from the schema.
7. **For new layouts, compute coordinates.** 16:9 = 960 × 540 pt. Leave ~40 pt
   margins. A 3-up row at top 120: boxes at left 60 / 360 / 660, width ~240.
8. **Wrap in `{"actions":[ ... ]}` and output ONLY that JSON** — no prose, no
   markdown fences. (Decko's sanitizer tolerates fences/prose, but clean output
   is safer, especially for weaker models.)
9. **If the batch is large, output one action object per line** so the VP can
   paste it without the text box mangling it (or tell them to save it to a
   `.json` file and use "Load from file...").

### Geometry-Sensitive Operations

Some requests require computing exact coordinates from the snapshot before emitting
any action. Recognize them by these phrases:

| Phrase | What it means geometrically |
|--------|-----------------------------|
| "above column N of the table" | shape left = table.left + sum of first N-1 column widths; shape width = column N width |
| "aligned with / flush with the table column" | same left + width as the column |
| "matching the size / style of existing sections" | read the existing shapes' `box=(L,T,W,H)` from the snapshot; copy those dimensions |
| "redistribute evenly above the columns" | set each section's left = its column's left edge; width = column width |
| "same top edge as the others" | read the lowest `top` among the existing sections; use that for new shapes |

#### How to derive table column boundaries

The snapshot reports the table as a single shape with `box=(table_left, table_top, table_width, table_height)`.
Unless per-column widths are shown, assume equal columns:

```
col_width = table_width / num_columns
col_left[i] = table_left + i * col_width          (i is 0-based)
col_right[i] = col_left[i] + col_width
```

"Above column N" (1-based) means:
- `shape.left  = col_left[N-1]`
- `shape.width = col_width`
- `shape.top`  = some consistent value above `table_top` (match existing sections or leave gap ~8–16 pt)

#### The inspect-first rule

Before emitting any move/resize/add that depends on existing geometry:

1. Read the snapshot for all relevant shape positions.
2. Identify the reference object (the table, an existing section, etc.) and record its `box`.
3. Compute all target coordinates from those values.
4. Only then emit the actions.

If the snapshot does not show per-column widths explicitly, use equal-split
arithmetic. If no snapshot is available at all, make reasonable assumptions
(e.g. equal column widths, standard 16:9 slide = 960×540 pt) and **call them
out clearly** in a comment or note above the JSON so the VP can adjust. Never
stop and wait for a snapshot — always emit a best-effort actions batch.

#### Pattern: add new section + redistribute to align with columns

This is the most common table-header pattern. Steps:

1. **Read** existing sections' `shape_id`, `box`, fill, font from the snapshot.
2. **Compute** column boundaries from the table's `box`.
3. **Add** the new section's boxes (`add_shape`) at the new column's left + width.
4. **Reposition** all existing sections (`move_shape` + `resize_shape`) to their
   respective column left + width — even if they already look right; this makes
   all three consistent.
5. Use the same `top` for all top-boxes and the same `top` for all bottom-boxes.

#### Pattern: "match existing style"

When the VP says "add something similar to the existing X", do **not** invent
dimensions, colors, or fonts. Read them from the snapshot:

- `box` → `width`, `height` for the new shape
- fill color → use `"fill"` field verbatim from snapshot's reported fill
- font size → copy from existing shape's text run
- gap between stacked boxes → `existing_bottom_box.top - (existing_top_box.top + existing_top_box.height)`

### How to write a good "VP prompt"

When the VP is going to paste a prompt into a model themselves, help them write
one that any model can execute. A good VP prompt:

- **Names the target.** "On slide 3, the title text box (top-left)" beats "the title".
- **States the change concretely.** "Change to 'Q3 FY24 Results'" beats "update the title".
- **Gives the data inline.** Numbers, labels, colors, fonts — spelled out, not "use the latest figures".
- **Specifies position/size in points** when adding things ("at left 60, top 120, width 560, height 340").
- **One slide at a time** for complex rebuilds; batch only simple repeated edits.
- **Ends with a reminder** that the model must output only the `{"actions":[...]}` JSON.

Template you can hand the VP:

```
[Using the snapshot above.]
On slide <N>, <do X to shape <id>/the <description>>:
- <concrete change 1, with exact text/numbers/colors>
- <concrete change 2>
Output only the Decko actions JSON ({"actions":[ ... ]}).
```

### Worked examples

**Example A — surgical text + color edit.**
*Snapshot excerpt:*
```
Slide 2  (layout: Title and Content, 960x540)
  shape_id=3  kind=placeholder(title)  box=(48,30,864,90)   text: "Company Overview"
  shape_id=7  kind=textbox              box=(60,140,500,40)  text: "Revenue: $1.2B (FY23)"
```
*VP request:* "On slide 2, change the title to 'Acme — Company Overview', and
make the revenue line bold and dark blue (#1F3864)."
*Output:*
```json
{"actions":[
  {"type":"set_text","slide":2,"shape_id":3,"value":"Acme — Company Overview"},
  {"type":"set_font_bold","slide":2,"shape_id":7,"value":true},
  {"type":"set_font_color","slide":2,"shape_id":7,"value":"#1F3864"}
]}
```

**Example B — add a chart to a new slide.**
*VP request:* "Add a new blank slide after slide 5 with a clustered column chart
of revenue by year — FY21 120, FY22 138, FY23 151, FY24 170 (in $M) — titled
'Revenue Trend', legend on, values labeled."
*Output:*
```json
{"actions":[
  {"type":"add_slide","position":6,"layout_index":6},
  {"type":"add_chart","slide":6,"chart_type":"columnclustered",
   "pos":{"left":80,"top":90,"width":800,"height":380},
   "categories":["FY21","FY22","FY23","FY24"],
   "series":[{"name":"Revenue ($M)","values":[120,138,151,170]}],
   "title":"Revenue Trend","show_legend":true,"show_values":true,
   "value_format":"$#,##0\"M\""}
]}
```
(Note: `add_slide` is action 1 so slide 6 exists when `add_chart` runs.)

**Example C — build a 3-step process flow on slide 4.**
*VP request:* "On slide 4, clear everything except the title and build a
3-step process flow left-to-right: 'Source', 'Refine', 'Deliver', dark navy
boxes (#15283C) with white text, arrows between them." *(Snapshot says the title
on slide 4 is shape_id=2.)*
*Output:*
```json
{"actions":[
  {"type":"clear_slide","slide":4,"keep_shape_ids":[2]},
  {"type":"add_shape","slide":4,"kind":"rrect","pos":{"left":60,"top":230,"width":240,"height":90},"fill":"#15283C","text":"Source","font_color":"#FFFFFF","font_size":20,"font_bold":true,"ref_name":"step1"},
  {"type":"add_shape","slide":4,"kind":"rrect","pos":{"left":360,"top":230,"width":240,"height":90},"fill":"#15283C","text":"Refine","font_color":"#FFFFFF","font_size":20,"font_bold":true,"ref_name":"step2"},
  {"type":"add_shape","slide":4,"kind":"rrect","pos":{"left":660,"top":230,"width":240,"height":90},"fill":"#15283C","text":"Deliver","font_color":"#FFFFFF","font_size":20,"font_bold":true,"ref_name":"step3"},
  {"type":"add_connector","slide":4,"kind":"elbow","from_shape_name":"step1","to_shape_name":"step2","arrow_end":"filled","color":"#15283C","weight_pt":2},
  {"type":"add_connector","slide":4,"kind":"elbow","from_shape_name":"step2","to_shape_name":"step3","arrow_end":"filled","color":"#15283C","weight_pt":2}
]}
```

**Example D — deck-wide rebrand.**
*VP request:* "Rebrand the whole deck from 'Globex' to 'Initech', swap the
accent color #C00000 → #0B5394 everywhere, and change the body font from
Calibri to Arial."
*Output:*
```json
{"actions":[
  {"type":"find_replace_text","scope":"deck","find":"Globex","replace":"Initech"},
  {"type":"recolor_palette_deck_wide","from_hex":"#C00000","to_hex":"#0B5394","target":"both"},
  {"type":"swap_font_deck_wide","from_name":"Calibri","to_name":"Arial"}
]}
```

**Example E — add a section above a table column; redistribute all sections to align with columns.**

*Snapshot excerpt:*
```
Slide 3 (960x540)
  shape_id=5   kind=table   box=(40,200,880,200)   rows=5 cols=4
  shape_id=10  kind=rect    box=(260,110,220,44)   text: "Label A"
  shape_id=11  kind=rect    box=(260,158,220,36)   text: "Value A"
  shape_id=12  kind=rect    box=(480,110,220,44)   text: "Label B"
  shape_id=13  kind=rect    box=(480,158,220,36)   text: "Value B"
```
*VP request:* "Slide 3 has a 4-column table. There are two sections above it — each
is two stacked boxes — one above column 2, one above column 3. Add a third section
above column 4 (blank text), matching the style of the existing ones. Then make all
three sections flush with their respective table columns."

*Geometry derivation:*
- Table: left=40, width=880 → 4 equal columns → col_width=220
- col2_left=260, col3_left=480, col4_left=700
- Existing top-box height=44, bottom-box height=36, gap=4 pt (158-110-44=4)
- Existing top y=110, bottom y=158

*Actions:*
```json
{"actions":[
  {"type":"move_shape","slide":3,"shape_id":10,"left":260,"top":110},
  {"type":"resize_shape","slide":3,"shape_id":10,"width":220,"height":44},
  {"type":"move_shape","slide":3,"shape_id":11,"left":260,"top":158},
  {"type":"resize_shape","slide":3,"shape_id":11,"width":220,"height":36},
  {"type":"move_shape","slide":3,"shape_id":12,"left":480,"top":110},
  {"type":"resize_shape","slide":3,"shape_id":12,"width":220,"height":44},
  {"type":"move_shape","slide":3,"shape_id":13,"left":480,"top":158},
  {"type":"resize_shape","slide":3,"shape_id":13,"width":220,"height":36},
  {"type":"add_shape","slide":3,"kind":"rect","pos":{"left":700,"top":110,"width":220,"height":44},"fill":"#FFFFFF","text":"","font_size":12,"ref_name":"sec3_top"},
  {"type":"add_shape","slide":3,"kind":"rect","pos":{"left":700,"top":158,"width":220,"height":36},"fill":"#FFFFFF","text":"","font_size":12,"ref_name":"sec3_bot"}
]}
```
> Copy the exact fill color and font from the snapshot's existing shapes — the
> values above use `#FFFFFF` as a placeholder. `col_left` always comes from
> `table.left + (col_index * col_width)`. Column 1 (left=40) has no section above it.

---

### Palette Retheme Workflow

Use this workflow when the VP wants to apply a client's brand colors to an existing deck. It requires no Python — only the VBA `scan_palette` action and this LLM.

#### Steps (user side)

1. **Open the source deck** (client brand palette slide) in PowerPoint alongside `PPT_AI_Editor.pptm`.
2. Run: `{"actions":[{"type":"scan_palette","scope":"slide:1"}]}` — this writes a JSON color list to the clipboard. Copy it.
3. **Open the destination deck** (the deck to recolor) and make it active.
4. Run: `{"actions":[{"type":"scan_palette"}]}` — writes destination color list to clipboard. Copy it.
5. Send both lists to the LLM with the request below.
6. Paste the LLM's `recolor_deck` action into Execute. Apply.

#### LLM prompt template

```
SOURCE PALETTE (colors I want to use):
<paste scan_palette output from brand/source deck>

DESTINATION PALETTE (current colors in the deck to recolor):
<paste scan_palette output from destination deck>

Map ALL destination colors to source colors and return a single recolor_deck action.
Rules:
1. Role match first: high-count fill colors map to the dominant fill colors in the source; font-only colors map to source font colors.
2. For unmatched: find the perceptually closest source color by hue + lightness similarity.
3. If multiple destination colors map to the same source color, preserve their relative lightness as tints/shades of that source color (lightest destination -> lightest tint, darkest -> darkest).
4. Never leave a destination color unmapped — every entry must appear in the mappings array.
5. Explain your mapping choices in plain English first, then output ONLY the JSON.

Output format:
{"actions":[{"type":"recolor_deck","mappings":[{"from":"#RRGGBB","to":"#RRGGBB"}, ...]}]}
```

#### How the LLM maps colors

The LLM reads **role + count** from the destination palette to understand each color's purpose:
- High count + `"fill"` role = primary brand fill → maps to the most prominent source fill color
- Low count + warm/red hue = accent → maps to source accent
- `"font"` only = body text color → maps to source text color
- Light/near-neutral `"fill"` = background or subtle fill → maps to source background color

For colors with no clear role match, the LLM computes perceptual distance (hue + lightness) and picks the closest source color. If a source palette has fewer colors than the destination, multiple destination colors may map to the same source color — tint scaling preserves visual hierarchy in that case.

#### Example F — client retheme

*Source palette scan (brand slide):*
```json
[
  {"hex":"#003087","count":3,"roles":["fill"]},
  {"hex":"#D03027","count":2,"roles":["fill"]},
  {"hex":"#F5F5F5","count":4,"roles":["fill"]},
  {"hex":"#1A1A1A","count":3,"roles":["font"]}
]
```

*Destination palette scan (deck to recolor):*
```json
[
  {"hex":"#113D63","count":158,"roles":["fill","border"]},
  {"hex":"#CCD1D7","count":68,"roles":["fill","border"]},
  {"hex":"#5E7C9E","count":71,"roles":["fill"]},
  {"hex":"#C0504D","count":53,"roles":["fill"]},
  {"hex":"#404040","count":42,"roles":["font"]}
]
```

*LLM reasoning:*
- `#113D63` (158 fills) = dominant brand fill → maps to source primary `#003087`
- `#5E7C9E` (71 fills, medium blue, lighter shade of `#113D63`) = secondary blue → maps to `#003087` tinted ~40% lighter → `#3A6099`
- `#CCD1D7` (68 fills, near-neutral) = background/subtle → maps to `#F5F5F5`
- `#C0504D` (53 fills, warm red) = accent → maps to source red `#D03027`
- `#404040` (42 font-only) = body text → maps to source text `#1A1A1A`

*Output:*
```json
{"actions":[{"type":"recolor_deck","mappings":[
  {"from":"#113D63","to":"#003087"},
  {"from":"#5E7C9E","to":"#3A6099"},
  {"from":"#CCD1D7","to":"#F5F5F5"},
  {"from":"#C0504D","to":"#D03027"},
  {"from":"#404040","to":"#1A1A1A"}
]}]}
```

---

### Intent → action cheat-sheet

| VP says | Use |
|---------|-----|
| "Change the company name throughout" | `find_replace_text`, `scope:"deck"` |
| "Fix this one number / phrase" | `find_replace_text` (literal) or `set_run_text` (if it's one run in a mixed paragraph) |
| "Rebuild this slide for Company X" | `clear_slide` (keep the title) → `add_text_box` / `add_shape` / `add_table` / `add_chart` |
| "Add a chart" | `add_chart` (39 types; data inline for the 32 standard ones) |
| "Add a logo / picture" | `insert_picture` (local path) or `replace_picture` |
| "Pull images from their website" | `fetch_page_images` → `build_image_picker_slide` and/or `build_image_grid_table` |
| "Add an icon / pictogram" | `insert_icon` — give the concept; use a name from the Fluent UI set / the allow-list in the export prompt |
| "Recolor the deck / rebrand colors" | Full retheme: `scan_palette` on source deck + `scan_palette` on destination → LLM maps palettes → `recolor_deck` (one action, all N pairs). Quick single-swap: `recolor_palette_deck_wide`. |
| "Change the font everywhere" | `swap_font_deck_wide` (or `set_theme_font` for theme-driven decks) |
| "Align / space out these shapes" | `align_shapes` + `distribute_horizontal` / `distribute_vertical` (or `tile_grid`, `smart_spacing`, `uniform_size`) |
| "Make a table" | `add_table` → `set_cell_text` per cell → `apply_table_style` |
| "Remove the gridlines / hide the axis on this chart" | `set_chart_gridlines` `{"props":{"major":false}}` (kills the horizontal lines); `set_chart_axis` `axis:"y"` `{"props":{"visible":false}}` to drop the vertical value axis (that also removes its gridlines) |
| "Format the data labels (thousands separator, negatives in parens)" | `set_chart_series` `{"props":{"show_labels":true,"label_format":"#,##0;(#,##0)"}}` — `label_format`, **not** the axis `number_format` |
| "Turn this table into a chart" | Read the table cells from the snapshot → `delete_shape` (the table) → `add_chart` with that data in the *same* `pos` (give it a `ref_name`, then `set_chart_*` follow-ups). If the rows are on incompatible scales (%, $, counts), pick a subset or normalize — say so. Reverse direction: `delete_shape` (chart) → `add_table` + `set_cell_text`. |
| "Re-flow / convert the layout (67/33, 50/50, quad, 3-/4-column, 2-left + 1-right, top/bottom, …)" | A batch of `move_shape` + `resize_shape` (target by `shape_name` if it has a `ref_name`, else `shape_id`); `delete_shape` + `add_*` for any transforms; `add_line` rules in the gutters; end with `enable_text_shrink_for_overflow`. Don't change text/colors unless asked. **Pick the region rectangles from [`LAYOUT_RECIPES.md`](LAYOUT_RECIPES.md)** — it has exact pt coordinates for every preset on a 960×540 slide. |
| "I wrote my points as bullets in one box — make it visual" | Read the box's `paragraphs` from the snapshot → `delete_shape` the box (keep the title) → `add_shape kind:"rrect"` (or `add_text_box`) one card per bullet, positioned per a region preset chosen by bullet count (2→L50/R50, 3→3-col or 1-left+2-right, 4→quad or 4-col, 6→2×3 grid; sequential steps → `chevron` flow). Vary the card fills (brand ramp). End with `enable_text_shrink_for_overflow`. Full recipe + worked examples in [`LAYOUT_RECIPES.md`](LAYOUT_RECIPES.md). |
| "Add shadow / glow / gradient / 3-D" | `set_shadow` / `set_glow` / `set_gradient_fill` / `set_3d_bevel` |
| "Build an org chart / process flow" | `add_shape` (each box with a `ref_name`) + `add_connector` (by `from_shape_name`/`to_shape_name`) |
| "Reorder / extract / merge slides" | `move_slide` / `extract_slides` / `import_slides_from_deck` |
| "Add speaker notes" | `set_speaker_notes` / `append_speaker_notes` |
| "Text is overflowing / getting cut off" | `enable_text_shrink_for_overflow` (slide or deck) |
| "Change slide size to 16:9" | `set_slide_size`, `preset:"16:9"` |
| "Add slide numbers" | `insert_slide_number` per slide (or via the template's master) |
| "Color the slide background" | `set_slide_background_color` |
| "Position shapes above / aligned with table columns" | Derive `col_left = table.left + (col_index * col_width)`, `col_width = table.width / num_cols`. Use `move_shape` + `resize_shape` to set each shape's `left` and `width` to the column's values. |
| "Add a section matching the style of existing ones" | Read existing shape `box` (L,T,W,H), fill, font from snapshot. `add_shape` with those exact dimensions and fill at the new column's left. |
| "Redistribute / realign sections above table columns" | Set each section's `left = col_left[N]`, `width = col_width` via `move_shape` + `resize_shape`. Use the same `top` for all. Anchor: table geometry, not hardcoded offsets. |

### Always remind the VP

1. **Share the snapshot if you can** — it gives the model shape IDs, slide numbers,
   and current text, enabling precise actions. Export via Alt+F8 → ExportSnapshot.
   If you can't share it, the model will proceed with best-effort assumptions and
   state them clearly — you can correct anything that looks wrong before applying.
2. **Workflow:** (Optional) Alt+F8 → ExportSnapshot → Copy snapshot+template →
   paste into model → replace the placeholder line → Alt+F8 → ExecuteInstructions →
   Parse → (large batch: "Load from file...") → Apply → **Ctrl+S to save**.
3. **Decko auto-backs-up before every Apply** (`<deck>_backup_<timestamp>`) and
   logs every action to `<deck>.action_log.jsonl` — if something looks wrong,
   the log says which actions ran, were skipped, or errored, and why.

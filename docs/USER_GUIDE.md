# Decko.ai — A Practical Guide

**For VPs, MDs, and senior bankers who want to stop wrestling with PowerPoint.**

Decko.ai is a teammate sitting inside your PowerPoint deck. You describe what you want in plain English, an LLM (ChatGPT, Claude, or your firm's internal assistant) writes a precise set of instructions, and Decko applies them to your deck in seconds.

You keep full control. Nothing is uploaded anywhere. No autosave, no surprises. You hit `Ctrl+S` when you're happy.

---

## The 30-Second Mental Model

Think of editing a deck as three steps:

1. **Show the LLM your slide.** Decko exports a structured "snapshot" — every shape, every word, every color, every font — into your clipboard.
2. **Tell the LLM what you want.** "Rebuild this slide for Eli Lilly." "Make the comps table use our navy palette." "Add a deal structure org chart on a new slide."
3. **Apply.** Decko reads the LLM's response and executes every change in one shot.

That's it. The LLM is the brain. Decko is the hands.

---

## The Daily Workflow

Open your deck in PowerPoint. Make sure `PPT_AI_Editor.pptm` (the carrier) is also open in the background — this is the engine.

### Step 1 — Export the snapshot

Press `Alt + F8`, run **`ExportSnapshot`**.

A dark-themed window appears with two big text boxes:
- **Snapshot** (top): the structured description of your slide(s)
- **Prompt template** (bottom): instructions baked in for the LLM — explains the action vocabulary, prevents formatting destruction, sets edit priorities.

Click **Copy Both**. Paste into ChatGPT, Claude, or your firm's LLM aggregator.

Add **one sentence** describing what you want. Examples:

> "Rebuild this slide swapping Amgen for Eli Lilly using current market data."

> "Add a 6-row trading comps table on slide 8 with our navy header."

> "Convert the bullet list into a 4-step process flow with chevron shapes."

### Step 2 — Get the action plan

The LLM responds with a JSON block (a structured list of instructions). Copy it.

### Step 3 — Apply

Press `Alt + F8`, run **`ExecuteInstructions`**. Paste. Click **Parse**, review the action count, click **Apply**.

Done. Hit `Ctrl+S` if you like the result.

---

## What Decko Can Do — A Tour by Use Case

Think of these as the things you used to ask an analyst to do at midnight. (16 categories, ~109 actions total.)

### 1. Edit Existing Text (without breaking formatting)

This is the bread and butter. The most common mistake bankers make is editing in PowerPoint and losing the bold/colored runs in a paragraph (e.g., a drug name in bold mid-sentence). Decko's edit hierarchy preserves formatting:

- **Find & replace a word or phrase** anywhere on the slide or deck
- **Edit a single bold/colored fragment** without touching surrounding text
- **Replace a whole paragraph** as a last resort
- **Bulk swap a name throughout the deck** (Amgen → Eli Lilly)
- **Regex-based replacements** for patterns ($1.2B → $2.4B)

### 2. Update Numbers in Tables

- Change a single cell value
- Add or delete rows/columns
- Swap entire columns or rows
- Merge cells
- Apply borders, fills, alignment per cell

### 3. Build Tables From Scratch

This is new and powerful. You can ask the LLM to build a comps table, deal-team table, or sources & uses table from nothing:

- Specify rows × columns + position
- Apply a built-in PowerPoint style (light, medium, dark — with accent colors)
- Set custom column widths and row heights
- Add cell borders (color, weight, side)
- Color individual cells or whole header rows
- Center, left, or right align cell text
- Vertically center cell contents

### 4. Update Chart Data

For native PowerPoint charts (the kind you build with Insert → Chart):

- Change the bar values
- Update category labels
- Rename a series
- Switch chart type (bar → line, etc.)
- Toggle title and axis titles
- Change legend position
- Recolor a series

**Caveat:** if your chart was pasted in as a picture or as an embedded Excel range (common in IB pitchbooks), this won't apply. The chart needs to be a real PowerPoint chart object.

### 5. Build Boxes, Badges, and Callouts

You can create any shape from scratch with text, fill, stroke, and font properties — all in one instruction. Available shape kinds include:

- **Basics:** rectangles, rounded rectangles, circles, ovals, triangles, diamonds, hexagons, octagons, pentagons, parallelograms
- **Arrows:** right, left, up, down, double-headed, curved, bent, U-turn, striped, notched
- **Process:** chevrons, chevron-pentagons (home-plate)
- **Callouts:** rectangular, rounded, oval, cloud, line callouts
- **Stars:** 4-, 5-, 8-, 10-, 12-, 16-, 24-, 32-point
- **Business motifs:** rings/donuts, brackets, braces, plaques, banners, "no" symbol, cloud
- **Capsule** for badge-style elements

You can also stack shapes (rectangle behind a circle for a numbered badge), control z-order (front/back/forward/backward), duplicate a shape to a new position, and copy formatting between shapes.

### 6. Connectors and Arrows (org charts done right)

This used to be impossible without manual drawing. Now:

- **Connector types:** straight, elbow, curved
- **Connection points:** specify exactly where a connector attaches — top, right, bottom, left of a shape
- **Arrow heads:** filled triangle, open, stealth, diamond, oval, none — on either end
- **Arrow size:** small, medium, large
- **Line styles:** solid, dash, dot, round-dot, dash-dot, long-dash, long-dash-dot
- **Color and weight**

This is what makes a real org chart possible — bottom-of-parent-box to top-of-child-box, with proper triangle arrowheads.

### 7. Slide-Level Changes

- **Add a slide** at any position with any layout
- **Delete a slide**
- **Duplicate a slide**
- **Move a slide** to a different position
- **Extract slides** into a new deck file
- **Import slides** from another deck (preserves formatting and contiguous-range optimization)
- **Set slide background color** (instant section divider)
- **Insert a slide number** as a styled text box
- **Apply a layout** to a range of slides

### 8. Deck-Wide Polish

- **Find/replace across the whole deck**
- **Swap a font everywhere** (Calibri → Garamond)
- **Recolor a palette deck-wide** (replace one hex with another, on fills, fonts, or both)
- **Apply a theme** from a `.thmx` file
- **Change slide size** (16:9, 4:3, custom dimensions)
- **Set theme major/minor fonts**
- **Bulk insert an image** on multiple slides at once
- **Bulk insert a text box** on multiple slides at once

### 9. Visual Effects

- **Rotate, flip** (horizontal or vertical)
- **Line color, weight, style**
- **Shadow** (offset, blur, color, transparency)
- **Glow** (color, radius, transparency)
- **Reflection** (size, transparency, distance)
- **Transparency** (0–1)
- **Gradient fill** (two-color with angle)
- **3D bevel** (circle, slope, cross, angle, soft round)
- **24 preset effects**
- **Picture crop** (from each side)
- **Recolor picture** (grayscale, sepia, washout, B&W)
- **Brightness and contrast** (-1 to +1)

### 10. Layout & Alignment Toolbox

This is what separates a sloppy deck from a polished one:

- **Align shapes** to a reference (left, right, top, bottom, horizontal-center, vertical-center)
- **Distribute** shapes evenly horizontally or vertically
- **Tile in a grid** (specify columns, gap)
- **Fit a shape to slide margins** (with custom margin)
- **Move relative** by dx, dy
- **Snap to grid**
- **Center on slide** (h, v, or both)
- **Nudge** by direction and amount
- **Match size** of one shape to another reference
- **Uniform size** across many shapes
- **Smart spacing** (place shapes next to each other with gap)
- **Equalize spacing** between shapes
- **Match position** edge-to-edge
- **Swap positions** of two shapes
- **Group by overlap** (auto-group only the overlapping ones)
- **Match fit-to-content** (auto-resize a shape to fit its text)

### 11. Run-Level Formatting (mixed formatting in one paragraph)

When a sentence has a bold drug name in the middle, you don't want to lose that. Decko handles "runs" (the formatting fragments inside a paragraph):

- Bold, italic, underline, **strikethrough**
- Subscript, superscript
- Per-run font color, size, name
- Per-run text replacement (preserves surrounding bold/italic)
- Per-run hyperlink (set or clear)

### 12. Text Frame Polish

- **Paragraph alignment** (left, center, right, justify)
- **Line spacing**
- **Vertical alignment** (top, middle, bottom)
- **Text auto-fit modes** (none, shrink-to-fit, resize-shape)
- **Text margins**
- **Auto-shrink overflowing text** across a slide or whole deck (smart — only shrinks shapes whose text actually overflows, leaves badge circles alone)

### 13. Pictures

- **Insert** a picture by file path
- **Replace** an existing picture (perfect for logo swaps)
- **Crop, recolor, brightness, contrast** as listed above

### 14. Web Images — Scrape, Pick, and Build a Grid Table

When you want a slide that shows images pulled from a company's website
(product photos, industry applications, use-case imagery), Decko uses a
3-step workflow:

1. **Scrape** — download all images from a URL into a local folder
2. **Pick** — drop thumbnails onto a new slide (4 per row) so you can
   identify which filenames match which images
3. **Build** — assemble a polished 2-column image + description table on
   the target slide, with image overlay, name strip, and bullet text per row

Example prompt flow:

*Step 1 (you ask):*
> "Fetch all images from https://www.stanley.com/industries and show me a
> picker slide so I can identify filenames."

*Step 2 (after you see the picker and note filenames — you ask):*
> "Build the industry applications table on slide 6 using img_003.jpg for
> Aerospace, img_007.jpg for Automotive, img_012.jpg for Construction."

You can also **download a single image on the fly** by URL and insert it
at a specific position without the full scrape workflow.

### 15. Icons (Microsoft Fluent UI)

Insert scalable vector icons from Microsoft's Fluent UI icon set — the same
icon library used in Microsoft 365. Icons are fetched directly from the CDN
at render time (no manual download needed).

- **3,000+ icons** covering business, tech, industry, and UI concepts
- **Styles:** filled (default) or regular (outline)
- **Sizes:** 16 / 20 / 24 / 28 / 32 / 48 pt
- **Color:** any hex color — or omit for default dark grey
- **Position:** same left/top/width/height system as shapes

How to find an icon name: go to [fluenticons.co](https://fluenticons.co),
search for a concept (e.g., "factory", "globe", "chart"), and use the icon
name in `lowercase_underscore` format.

Example prompt:
> "On slide 3, add a filled 'building_factory' icon in our brand navy
> (#15283C) at the top-left of each industry card, 48pt."

### 16. Speaker Notes

- **Set speaker notes** (replace)
- **Append** to existing notes

---

## Talking to the LLM — Prompt Guidance

The LLM does the thinking. The clearer you brief it, the better the result.

### Rule 1 — Always paste the snapshot AND the prompt template

The template tells the LLM:
- What action types are available
- That find_replace_text is preferred over set_paragraph_text (preserves formatting)
- That whole-paragraph rewrites lose bold/italic
- That trailing line breaks cause duplicate paragraphs
- That overflow guards should be appended

If you skip the template, the LLM defaults to clumsy choices (like rewriting a whole paragraph just to delete one word).

### Rule 2 — One sentence is usually enough

> "Rebuild this slide for Eli Lilly using current market data."

The template covers the rest.

### Rule 3 — Be explicit when the change is small or surgical

> "Delete the word 'the' from the third bullet."

> "Change the price target from $145 to $172 in the analyst commentary callout."

If you only mark something visually (e.g., strikethrough one word in a screenshot), also state it in plain text. LLM vision can misread small marks.

### Rule 4 — Use snapshot coordinates for moves, not screenshot estimation

When you ask the LLM to align or move a shape, tell it:

> "Use the snapshot for exact positions. For moves, compute from existing shape coordinates — do not estimate from the screenshot."

This prevents the LLM from guessing pixel positions and ending up off by 5pt.

### Rule 5 — Always name new shapes with `ref_name`

When the LLM creates shapes from scratch, every shape should get a `ref_name`. This is a label you can use to reference it later in the same instruction batch:

```json
{"type":"add_shape","kind":"rect","ref_name":"deal_box",...}
{"type":"add_shape","kind":"circle","ref_name":"badge_4",...}
{"type":"z_order","shape_name":"badge_4","order":"front"}
```

You can chain create → name → reference in one batch.

### Rule 6 — When in doubt, ask the LLM to explain its plan first

> "Walk me through what you'd change before generating the JSON."

Then once you agree, ask for the JSON. Saves applying something you didn't intend.

---

## Worked Examples — Real Prompts and Outcomes

These are scenarios from actual pitch books. Each one is something a VP would routinely ask for.

### Example 1 — Profile rebuild (Amgen → Eli Lilly)

**Prompt to LLM:**
> "Rebuild slide 12 (Amgen company profile) for Eli Lilly. Use current market data. Preserve all formatting and the layout."

**What you get back:**
A list of `find_replace_text` and `set_run_text` actions that swap the company name, market cap, key drugs, geographic footprint, and headline metrics — without touching the slide's structure or formatting.

### Example 2 — Add a comps table to a blank slide

**Prompt:**
> "On slide 8, build a 6-row trading comps table with columns Company, EV ($B), EV/Rev, EV/EBITDA, P/E. Use medium_style_2_accent1 with a navy header. Fill in Vertex, Alkermes, BioMarin, Ironwood, Lannett."

**What you get back:**
- One `add_table` action
- One `apply_table_style`
- A header `set_cell_fill` (navy) + `set_cell_text` for column titles
- 5 × 5 = 25 `set_cell_text` actions for the data
- A few `set_table_col_width` to make the Company column wider

All in a single batch.

### Example 3 — Org chart from a list of names

**Prompt:**
> "Slide 14, build an org chart: parent = JAZZ Pharmaceuticals plc, three children = Jazz Therapeutics, GW Pharmaceuticals, Pharmos Inc. Connect each child to parent with elbow connectors, bottom-of-parent to top-of-child, gray triangle arrowheads."

**What you get back:**
- 4 `add_shape` actions (rounded rectangles with company names)
- 3 `add_connector` actions with `from_point: bottom`, `to_point: top`, `arrow_end: filled`

### Example 4 — Section divider slide

**Prompt:**
> "Add a section divider before slide 5: navy background, large white centered title 'Valuation Analysis', and a 4-step chevron flow showing Trading Comps, Precedent Tx, DCF SoTP, Football Field."

**What you get back:**
- `add_slide`
- `set_slide_background_color`
- `add_text_box` for title and subtitle
- 4 `add_shape` calls with chevron kind

### Example 5 — Surgical edit from a marked-up screenshot

**Prompt** (with marked-up screenshot attached):
> "Red strikethrough on the screenshot = delete that text. Use find_replace_text for word deletions, never set_paragraph_text."

**What you get back:**
- `find_replace_text` actions targeting only the marked spans
- One `enable_text_shrink_for_overflow` at the end as an overflow guard

### Example 6 — Deck-wide font and color refresh

**Prompt:**
> "Swap Calibri to Garamond throughout the deck. Replace the old blue #2E75B6 with our brand navy #1F4E79 on all fills."

**What you get back:**
- One `swap_font_deck_wide` action
- One `recolor_palette_deck_wide` action

Two lines, full deck refresh.

### Example 7 — Logo swap

**Prompt:**
> "Replace the company logo on slide 3 (currently shape_id 7) with the file at C:\\Users\\me\\Downloads\\new_logo.png."

**What you get back:**
- One `replace_picture` action

### Example 8 — Numbered badge pattern (from a sketch)

**Prompt** (with hand sketch attached):
> "Build the methodology rows shown in the sketch. Each row = a navy rectangle with white bold text + a small dark-navy circle on the left with a letter A through G + a red X mark on the far right if the methodology is not used in the valuation."

**What you get back:**
- For each row: `add_shape` (rect), `add_shape` (circle for badge), `z_order` to bring badge to front, optional `add_text_box` for the X.

### Example 9 — Convert a bullet list into a process flow

**Prompt:**
> "Slide 6 has a 4-bullet list ('Research', 'Outreach', 'Diligence', 'Execution'). Replace it with a horizontal chevron process flow using our brand blue."

**What you get back:**
- 4 `delete_paragraph` (high-to-low order) on the original bullet shape — or `delete_shape` if it's a standalone text frame
- 4 `add_shape` with `kind: chevron`, evenly spaced

### Example 10 — Industry applications slide from a company website

**Prompt (part 1):**
> "Fetch all images from https://www.stanley.com/industries and show me a
> picker slide so I can identify which filenames I want."

*(After seeing the picker slide, you note img_003 = Aerospace, img_007 = Automotive, img_012 = Construction)*

**Prompt (part 2):**
> "On slide 6, build a 2-column image+description table using those images:
> Row 1: Aerospace, img_003, bullets = Lightweight fasteners / FAA compliant
> Row 2: Automotive, img_007, bullets = High-volume assembly / Vibration-resistant
> Row 3: Construction, img_012, bullets = Heavy-duty anchors / Code compliant
> Table at left=30, top=60, width=900, height=480. Name font 12pt bold navy. Desc font 10pt grey."

**What you get back:**
- `fetch_page_images` → `build_image_picker_slide` (part 1)
- `build_image_grid_table` with 3 fully styled rows (part 2)

### Example 11 — Icons on a feature card slide

**Prompt:**
> "On slide 5 I have three feature cards side by side. Add a filled Fluent UI
> icon above the headline on each card:
> - Card 1 (Growth): 'arrow_trending' icon, navy #1F4E79, 48pt
> - Card 2 (Innovation): 'lightbulb' icon, navy #1F4E79, 48pt
> - Card 3 (Global): 'globe' icon, navy #1F4E79, 48pt
> Use the snapshot to find each card's left edge and place the icon 10pt from
> the left, 10pt from the card top."

**What you get back:**
- 3 `insert_icon` actions, each with exact position derived from snapshot coordinates

### Example 12 — Force-fit overflowing text after a content swap

**Prompt:**
> "Replace 'Q4 results' with 'fourth-quarter results across all major operating segments' wherever it appears, then make sure no text overflows."

**What you get back:**
- A few `find_replace_text` actions
- One `enable_text_shrink_for_overflow` with `scope: deck` at the end

---

## Hidden Power Moves

Things that are not obvious from the feature list:

### A. The "create + style + position" chain

In one batch, you can: create a shape, name it, apply text + font + fill, then move/resize it, then send it to back, then connect to another named shape. No need to wait for one action to complete and then ask for another.

### B. ref_name everywhere

Any new shape can be given a name. Any subsequent action can reference it by `shape_name` instead of `shape_id`. You can stop thinking in numeric IDs.

### C. The overflow guard

Append `{"type":"enable_text_shrink_for_overflow","scope":"slide:N"}` at the end of any content swap. Decko checks each shape — if its text actually overflows, it enables shrink-to-fit. Shapes whose text already fits are untouched. Badge circles and tight-fit labels are safe.

### D. Bulk operations across a slide range

`bulk_insert_image` and `bulk_insert_text_box` accept a list of slide indices — one action stamps a logo across slides 5, 7, 9, 11.

### E. Smart layout helpers

When a deck has been edited piecemeal and shapes are misaligned, ask the LLM to use `align_shapes`, `equalize_spacing`, `smart_spacing`, or `tile_grid`. Cleans up a sloppy slide in one batch.

### F. Recolor by current value

`recolor_fill_match` and `recolor_font_match` change the color of every shape that currently has a specific color. Useful for "make every red callout into our brand red instead of plain red."

### G. Style pass via copy_formatting

If you've built one perfect-looking shape, ask the LLM to "copy the formatting from shape `golden_box` to all other rectangles on slide 4." Decko has `copy_formatting` for that.

### H. Group by overlap

When you have a stack of overlapping shapes that should move together, `group_by_overlap` auto-groups only the overlapping subset.

### I. Live regex find/replace

`find_replace_regex` runs across the whole deck. Useful for patterns: "$1.2bn" → "$2.4bn" (different sums but same prefix), date formats, ticker symbol re-formatting.

### J. Hyperlink runs

You can put a hyperlink on a single word inside a paragraph without affecting other words. Set the URL or pass an empty string to clear the link entirely.

---

## Honest Limitations

So you don't waste time fighting the wrong battles:

- **Charts pasted as pictures or embedded Excel can't be edited** by the chart-data API. Decko only edits native PowerPoint chart shapes. For pictures, replace the picture. For embedded Excel, edit the linked range and refresh.
- **Connector connection points are most accurate on rectangles, ovals, rounded rectangles.** For triangles, arrows, and irregular shapes, the connection points may not match exactly. Use straight connectors or accept that the connector may attach in an unexpected spot.
- **Coordinate estimation from screenshots is approximate.** When using a marked-up screenshot to instruct the LLM, always give it the snapshot too — coordinates from the snapshot are exact.
- **Sketches → slide skeletons work well, sketches → pixel-perfect slides do not.** Treat the LLM's first pass as 80% of the way there.
- **Auto-backup and auto-save are intentionally OFF.** You are in charge of saving with `Ctrl+S`. Closing without saving loses everything Decko did.
- **Color must be specified.** A black-and-white sketch won't tell the LLM what colors to use. Either describe your brand palette in the prompt or include a colored reference slide.
- **Theme colors require a `.thmx` file** if you want to apply one — the LLM can't invent themes from scratch.

---

## A Few Power Tips

- **Slide range scopes:** Most deck-wide actions accept `scope: "deck"` or `scope: "slide:N"`. Use `slide:N` to keep changes local.
- **The action log:** Every batch writes a `.action_log.jsonl` file next to your deck. If something looks off, check the log to see what was applied vs skipped vs errored.
- **Reverse-order edits for paragraphs:** If you're adding/deleting paragraphs, tell the LLM to do it high-to-low. Decko already pre-sorts run-level actions, but paragraph-level deletes can shift indices if done low-to-high.
- **The longest-first rule for find/replace:** When swapping a name like "Amgen Inc." → "Eli Lilly and Company", do "Amgen Inc." first, then "Amgen", to avoid double substitution.
- **Sanitizer is forgiving:** If the LLM wraps the JSON in code fences, adds prose, uses smart quotes, or trails commas, Decko cleans it up. You don't have to babysit the LLM's output formatting.

---

## Quick Reference — Action Vocabulary

This is the full list of what Decko understands. Use these names when reviewing what the LLM produced.

**Text:** `set_text`, `set_paragraph_text`, `set_run_text`, `add_paragraph`, `delete_paragraph`, `find_replace_text`, `find_replace_regex`, `set_paragraph_alignment`, `set_paragraph_line_spacing`, `set_text_vertical_align`, `set_text_autofit`, `enable_text_shrink_for_overflow`, `set_text_margin`, `set_bullet_style`, `set_indent_level`

**Run-level:** `set_run_bold`, `set_run_italic`, `set_run_underline`, `set_run_strikethrough`, `set_run_subscript`, `set_run_superscript`, `set_run_font_color`, `set_run_font_size`, `set_run_font_name`, `set_run_hyperlink`

**Font (shape-level):** `set_font_size`, `set_font_bold`, `set_font_italic`, `set_font_color`, `set_fill_color`, `set_paragraph_font_size`, `set_paragraph_font_color`

**Shape lifecycle:** `add_shape`, `add_text_box`, `delete_shape`, `move_shape`, `resize_shape`, `move_shape_relative`, `nudge`, `duplicate_shape`, `copy_formatting`, `z_order`, `clear_slide`, `set_shape_kind`

**Layout:** `align_shapes`, `align_to_slide_center`, `distribute_horizontal`, `distribute_vertical`, `tile_grid`, `fit_to_slide_margins`, `snap_to_grid`, `match_size`, `uniform_size`, `smart_spacing`, `equalize_spacing`, `match_position`, `swap_positions`, `fit_to_content`, `group_shapes`, `ungroup`, `group_by_overlap`

**Tables:** `add_table`, `add_table_row`, `delete_table_row`, `add_table_col`, `delete_table_col`, `merge_cells`, `set_cell_text`, `swap_table_columns`, `swap_table_rows`, `set_table_col_width`, `set_table_row_height`, `set_cell_border`, `set_cell_text_align`, `set_cell_fill`, `apply_table_style`

**Charts:** `set_chart_type`, `set_chart_title`, `set_chart_axis_title`, `set_chart_legend_position`, `set_series_color`, `set_series_values`, `set_chart_categories`, `set_series_name`

**Connectors:** `add_connector`, `add_line`

**Pictures:** `insert_picture`, `replace_picture`, `crop_picture`, `recolor_picture`, `set_brightness`, `set_contrast`

**Web images:** `fetch_page_images`, `build_image_picker_slide`, `build_image_grid_table`, `download_image`

**Icons:** `insert_icon`

**Effects:** `rotate_shape`, `flip_shape`, `set_line_color`, `set_line_weight`, `set_line_style`, `set_shadow`, `set_glow`, `set_reflection`, `set_transparency`, `set_gradient_fill`, `set_3d_bevel`, `apply_preset_effect`

**Slides:** `add_slide`, `delete_slide`, `duplicate_slide`, `move_slide`, `extract_slides`, `import_slides_from_deck`, `set_slide_background_color`, `insert_slide_number`

**Deck-wide:** `recolor_fill_match`, `recolor_font_match`, `delete_shapes_match`, `swap_font_deck_wide`, `recolor_palette_deck_wide`, `apply_theme`, `set_slide_size`, `set_theme_font`, `bulk_insert_image`, `bulk_insert_text_box`, `apply_layout_to_slides`

**Notes:** `set_speaker_notes`, `append_speaker_notes`

---

## When Something Goes Wrong

- **"Invalid JSON" error:** the LLM included prose. The sanitizer handles most cases but if it fails, copy only the JSON block (between the first `{` and last `}`).
- **"shape not found":** the LLM used a `shape_id` that doesn't exist on that slide. Re-export the snapshot — IDs change between sessions.
- **Text formatting got crushed:** the LLM used `set_paragraph_text` instead of `set_run_text`. Re-prompt with "use set_run_text to preserve bold runs."
- **Shape ended up in the wrong place:** the LLM estimated coordinates from a screenshot. Re-prompt: "use snapshot coordinates, not visual estimation."
- **Double paragraph break:** trailing `\r` in a value. The action log will show it. Re-prompt: "no trailing line breaks in any text value."
- **PowerPoint refuses to close:** a leftover macro window. Save first, then close PowerPoint normally.

---

## The Bottom Line

Decko.ai turns "I need a junior to fix this slide" into "give me 30 seconds." Anything you'd hand to an analyst at 11pm — text edits, table builds, layout cleanup, branding refreshes, org charts, section dividers — Decko can do directly from a one-sentence English request.

The two skills that take you from "user" to "power user":

1. **Always paste the snapshot AND the prompt template.** Without the template the LLM is just guessing the action vocabulary.
2. **State exactly what you want in plain English.** No need to learn the JSON. The LLM writes that.

Everything else is gravy.

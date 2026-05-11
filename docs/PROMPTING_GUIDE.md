# Decko.ai — VP Prompting Guide

This guide shows how to tell the LLM what to do with your PowerPoint deck.
It covers every capability Decko currently supports, with real example prompts
you can paste or adapt. The last section is a technical reference for AI
assistants helping VPs compose prompts.

---

## How Decko Works (30-second recap)

1. Open your deck + `PPT_AI_Editor.pptm` in PowerPoint.
2. **Alt+F8 → ExportSnapshot** — opens the snapshot form.
3. Click **Copy snapshot + prompt template**. Your clipboard now has a full
   prompt block ending with `[REPLACE THIS LINE WITH YOUR REQUEST]`.
4. Paste into your LLM (Claude, ChatGPT, Gemini, Hermes, etc.).
5. Replace the placeholder line with your request. Hit send.
6. Copy the LLM's JSON response.
7. **Alt+F8 → ExecuteInstructions** — paste JSON, click **Apply**.
8. **Ctrl+S** to save. (Decko edits the open deck in memory — close without
   saving and you lose the changes.)

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

Icons are sourced from [fluenticons.co](https://fluenticons.co). Search for
a concept there to find the icon name, then use it in lowercase with
underscores (e.g., `building_factory`).

**Insert a single icon:**
```
On slide 3, insert a filled "building_factory" Fluent UI icon at
left=100, top=200, width=60, height=60, color #15283C.
```

**Add icons to a set of feature cards:**
```
On slide 5, add a Fluent UI icon to the top-left corner of each of the
three feature cards (shapes 8, 12, 16):
- Shape 8 (Growth): "arrow_trending" icon, filled, 48px, color #1F4E79
- Shape 12 (Innovation): "lightbulb" icon, filled, 48px, color #1F4E79
- Shape 16 (Global): "globe" icon, filled, 48px, color #1F4E79
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

**Change chart type:**
```
On slide 8, change the chart (shape 4) from a clustered bar to a
clustered column chart.
```

**Update chart data:**
```
On slide 8, update the chart (shape 4):
- Categories: "2021", "2022", "2023", "2024"
- Series "Revenue": 24.0, 26.8, 28.5, 33.4
- Series "Net Income": 5.9, 6.1, 7.0, 8.8
- Chart title: "Eli Lilly Financial Performance ($B)"
- X-axis title: "Year"
- Y-axis title: "USD Billions"
```

**Recolor chart series:**
```
On slide 8, set series 1 color to #1F4E79 and series 2 color to #2E75B6.
```

**Move chart legend:**
```
On slide 8, move the chart legend to the bottom.
```

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

---

## Full Action Reference (for AI Assistants)

This section is a complete map of every action Decko supports. AI assistants
(openclaw, hermes, etc.) should use this to guide VPs toward the right prompt
phrasing for any task.

### Action Categories & Counts

| Module | Count | Actions |
|--------|-------|---------|
| `modActions.bas` | 17 | Core shape/slide/notes/table |
| `modActionsText.bas` | 14 | Granular text, paragraph-level, find/replace |
| `modActionsRun.bas` | 10 | Run-level formatting (bold/italic/color/hyperlink) |
| `modActionsLayout.bas` | 24 | Align, distribute, grid, recolor batch, clear slide |
| `modActionsTable.bas` | 5 | Table row/col/merge |
| `modActionsChart.bas` | 5 | Chart type/title/axis/series |
| `modActionsImage.bas` | 2 | Insert/replace picture |
| `modActionsConnector.bas` | 1 | Add connector |
| `modActionsGroup.bas` | 2 | Group/ungroup |
| `modActionsSlide.bas` | 3 | Move/extract/import slides |
| `modActionsDeck.bas` | 9 | Regex, font swap, theme, recolor palette, bulk insert |
| `modActionsEffects.bas` | 16 | Shadow, glow, gradient, 3D, crop, recolor picture |
| `modActionsIcon.bas` | 1 | Insert Fluent UI SVG icon |
| `modActionsWeb.bas` | helper | fetch_page_images, download_image |
| **Total** | **~109** | |

### Action Quick-Reference

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
| `set_chart_title` | Title text + enabled flag. |
| `set_chart_axis_title` | x or y axis title. |
| `set_chart_legend_position` | top / right / bottom / left. |
| `set_series_color` | Color by series index. |
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
| `insert_icon` | Add a Fluent UI SVG icon. Find names at fluenticons.co. |

Params: `slide`, `icon` (lowercase_underscore name), `style` (filled/regular),
`size` (16/20/24/28/32/48), `color` (#RRGGBB), `left`, `top`, `width`, `height` (all in pt).

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
| `recolor_palette_deck_wide` | Replace color across fill/font/both. |
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

## Guidance for AI Assistants (openclaw / hermes)

When a VP describes what they want to do with a presentation, your job is to
suggest the right Decko prompt they should paste into the LLM. Use the table
above to match their intent to the action type, then write a clear prompt they
can use directly.

**Key translation rules:**

| VP says | Suggest |
|---------|---------|
| "Change the company throughout the deck" | `find_replace_text` with `scope: deck` |
| "Update the financials / data points" | `find_replace_text` (surgical) or `set_run_text` (mixed-format) |
| "Rebuild this slide for Company X" | `clear_slide` + `add_shape`/`add_text_box` sequence + latest public data |
| "Add a logo / image" | `insert_picture` (local path) |
| "Get images from their website" | `fetch_page_images` → `build_image_picker_slide` → `build_image_grid_table` |
| "Add an icon / pictogram" | `insert_icon` — look up name at fluenticons.co |
| "Recolor the whole deck" | `recolor_fill_match` / `recolor_font_match` / `recolor_palette_deck_wide` |
| "Align / space out these shapes" | `align_shapes` + `distribute_horizontal` / `distribute_vertical` |
| "Create a table" | `add_table` + `set_cell_text` + `apply_table_style` |
| "Add a shadow / gradient / effect" | `set_shadow` / `set_gradient_fill` / `set_glow` |
| "Build an org chart / process flow" | `add_shape` (with ref_name) + `add_connector` |
| "Extract these slides to send" | `extract_slides` |
| "Merge this slide from another file" | `import_slides_from_deck` |
| "Add speaker notes" | `set_speaker_notes` or `append_speaker_notes` |
| "Some text is getting cut off" | `enable_text_shrink_for_overflow` |

**Remind VPs of the workflow every time:**
1. Export snapshot (Alt+F8 → ExportSnapshot → Copy snapshot + prompt template)
2. Paste into LLM and replace the placeholder line
3. Execute (Alt+F8 → ExecuteInstructions → Apply)
4. **Ctrl+S to save**

**Always remind:** The LLM must see the snapshot — it contains shape IDs, slide
numbers, and existing text that the LLM needs to write correct actions.
Never send a prompt without the snapshot attached.

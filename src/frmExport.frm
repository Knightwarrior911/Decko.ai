VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmExport 
   Caption         =   "Decko.ai � Export Snapshot"
   ClientHeight    =   7200
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   10800
   OleObjectBlob   =   "frmExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function PromptTemplate() As String
    Dim s As String

    s = "You are editing a PowerPoint presentation. Below is the current state as JSON:" & vbCrLf & vbCrLf
    s = s & "```json" & vbCrLf & "{snapshot}" & vbCrLf & "```" & vbCrLf & vbCrLf
    s = s & "CANVAS DIMENSIONS (CRITICAL):" & vbCrLf
    s = s & "  Read deck.slide_width_pt and deck.slide_height_pt from the snapshot above." & vbCrLf
    s = s & "  Default widescreen deck = 960pt x 540pt. Default 4:3 deck = 720pt x 540pt." & vbCrLf
    s = s & "  All pos.left/top/width/height values must respect these bounds — do NOT" & vbCrLf
    s = s & "  hardcode 720 for width unless the snapshot says so." & vbCrLf
    s = s & "  Full-bleed bands: pos.left=0, pos.width=slide_width_pt." & vbCrLf
    s = s & "TEXT-FIT GUIDANCE:" & vbCrLf
    s = s & "  Big numbers (font_size>=72) need pos.width >= font_size * 0.7 * char_count" & vbCrLf
    s = s & "  to avoid wrap. ""14-15%"" at font_size=88 needs ~430pt width." & vbCrLf
    s = s & "  When in doubt, append {""type"":""enable_text_shrink_for_overflow"",""scope"":""slide:N""}" & vbCrLf
    s = s & "  at the end of the batch — every text shape on slide N gets auto-shrink." & vbCrLf & vbCrLf
    s = s & "I want the following changes:" & vbCrLf & vbCrLf
    s = s & "[REPLACE THIS LINE WITH YOUR REQUEST]" & vbCrLf & vbCrLf

    s = s & "Return ONLY a valid instructions JSON. No prose, no explanation, no markdown" & vbCrLf
    s = s & "code fences. Top-level shape:" & vbCrLf & vbCrLf
    s = s & "{""actions"": [ <action>, <action>, ... ]}" & vbCrLf & vbCrLf

    s = s & "Each action is one of EXACTLY these schemas. Field names are STRICT - do not" & vbCrLf
    s = s & "rename ""value"" to ""text""/""color""/""size""/""fill"". Use names verbatim." & vbCrLf & vbCrLf

    s = s & "ATOMIC OPS (V1):" & vbCrLf
    s = s & "  {""type"":""set_text"",""slide"":1,""shape_id"":3,""value"":""Hello""}" & vbCrLf
    s = s & "  {""type"":""set_font_size"",""slide"":1,""shape_id"":3,""value"":28}" & vbCrLf
    s = s & "  {""type"":""set_font_bold"",""slide"":1,""shape_id"":3,""value"":true}" & vbCrLf
    s = s & "  {""type"":""set_font_italic"",""slide"":1,""shape_id"":3,""value"":false}" & vbCrLf
    s = s & "  {""type"":""set_font_color"",""slide"":1,""shape_id"":3,""value"":""#FF0000""}" & vbCrLf
    s = s & "  {""type"":""set_fill_color"",""slide"":1,""shape_id"":4,""value"":""#2E75B6""}" & vbCrLf
    s = s & "  {""type"":""move_shape"",""slide"":1,""shape_id"":4,""left"":100,""top"":200}" & vbCrLf
    s = s & "  {""type"":""resize_shape"",""slide"":1,""shape_id"":4,""width"":250,""height"":80}" & vbCrLf
    s = s & "  {""type"":""delete_shape"",""slide"":1,""shape_id"":7}" & vbCrLf
    s = s & "  {""type"":""add_slide"",""position"":3,""layout_index"":1}" & vbCrLf
    s = s & "  {""type"":""delete_slide"",""slide"":4}" & vbCrLf
    s = s & "  {""type"":""duplicate_slide"",""slide"":2}" & vbCrLf
    s = s & "  {""type"":""set_cell_text"",""slide"":1,""shape_id"":5,""row"":2,""col"":1,""value"":""Revenue""}" & vbCrLf
    s = s & "  {""type"":""swap_table_columns"",""slide"":1,""shape_id"":5,""col_a"":1,""col_b"":2}" & vbCrLf
    s = s & "  {""type"":""swap_table_rows"",""slide"":1,""shape_id"":5,""row_a"":1,""row_b"":2}" & vbCrLf & vbCrLf

    s = s & "GRANULAR TEXT:" & vbCrLf
    s = s & "  {""type"":""set_paragraph_text"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""...""}" & vbCrLf
    s = s & "  {""type"":""add_paragraph"",""slide"":1,""shape_id"":3,""after_paragraph_index"":-1,""value"":""...""}" & vbCrLf
    s = s & "  {""type"":""delete_paragraph"",""slide"":1,""shape_id"":3,""paragraph_index"":2}" & vbCrLf
    s = s & "  {""type"":""set_bullet_style"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""disc""}" & vbCrLf
    s = s & "  {""type"":""set_indent_level"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":1}" & vbCrLf
    s = s & "  {""type"":""set_paragraph_font_size"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":24}" & vbCrLf
    s = s & "  {""type"":""set_paragraph_font_color"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""find_replace_text"",""scope"":""deck"",""find"":""ACME"",""replace"":""NewCo""}" & vbCrLf & vbCrLf

    s = s & "LAYOUT / COMPOSITION:" & vbCrLf
    s = s & "  {""type"":""align_shapes"",""slide"":1,""shape_ids"":[3,4,5],""anchor"":""top""}" & vbCrLf
    s = s & "  {""type"":""distribute_horizontal"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""distribute_vertical"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""tile_grid"",""slide"":1,""shape_ids"":[3,4,5,6],""cols"":2,""gap_pt"":10}" & vbCrLf
    s = s & "  {""type"":""fit_to_slide_margins"",""slide"":1,""shape_id"":3,""margin_pt"":36}" & vbCrLf
    s = s & "  {""type"":""add_line"",""slide"":1,""x1"":50,""y1"":50,""x2"":700,""y2"":50,""color"":""#000000"",""weight_pt"":1.5}" & vbCrLf
    s = s & "  {""type"":""add_shape"",""slide"":1,""kind"":""capsule"",""pos"":{""left"":100,""top"":100,""width"":200,""height"":60},""fill"":""#1F4E79"",""stroke"":null,""stroke_weight_pt"":0}" & vbCrLf
    s = s & "  {""type"":""set_shape_kind"",""slide"":1,""shape_id"":4,""kind"":""capsule""}" & vbCrLf
    s = s & "  {""type"":""clear_slide"",""slide"":1,""keep_shape_ids"":[3]}" & vbCrLf
    s = s & "  {""type"":""move_shape_relative"",""slide"":1,""shape_id"":4,""dx_pt"":0,""dy_pt"":50}" & vbCrLf & vbCrLf

    s = s & "CROSS-CUTTING BATCH:" & vbCrLf
    s = s & "  {""type"":""recolor_fill_match"",""scope"":""deck"",""from"":""#0000FF"",""to"":""#FF0000""}" & vbCrLf
    s = s & "  {""type"":""recolor_font_match"",""scope"":""slide:2"",""from"":""#000000"",""to"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""delete_shapes_match"",""scope"":""deck"",""text_contains"":""Confidential""}" & vbCrLf & vbCrLf

    s = s & "SPEAKER NOTES:" & vbCrLf
    s = s & "  {""type"":""set_speaker_notes"",""slide"":1,""value"":""Talking points...""}" & vbCrLf
    s = s & "  {""type"":""append_speaker_notes"",""slide"":1,""value"":""Add this too""}" & vbCrLf & vbCrLf

    s = s & "IMAGES (LOCAL FILE PATHS ONLY - no URLs):" & vbCrLf
    s = s & "  {""type"":""insert_picture"",""slide"":1,""path"":""C:\\path\\to\\img.png"",""pos"":{""left"":50,""top"":50,""width"":200,""height"":150}}" & vbCrLf
    s = s & "  {""type"":""replace_picture"",""slide"":1,""shape_id"":7,""path"":""C:\\path\\to\\new.png""}" & vbCrLf & vbCrLf

    s = s & "WEB IMAGE WORKFLOW (scrape + visual pick + grid table):" & vbCrLf
    s = s & "  Use this 3-action sequence when the user wants images from a website" & vbCrLf
    s = s & "  (e.g. ""build a slide on company X's industry applications from URL"")." & vbCrLf
    s = s & "  Action 1 - download all images from a page into <deck>\\assets\\<slug>_<ts>\\:" & vbCrLf
    s = s & "    {""type"":""fetch_page_images"",""url"":""https://example.com/industries""}" & vbCrLf
    s = s & "  Action 2 - drop the downloaded images onto a NEW slide as a labeled grid" & vbCrLf
    s = s & "  so the user can visually identify which filenames they want:" & vbCrLf
    s = s & "    {""type"":""build_image_picker_slide"",""folder"":""<from_step1>"",""cols"":4}" & vbCrLf
    s = s & "  Action 3 - after the user tells you which filenames map to which row," & vbCrLf
    s = s & "  build a 2-column image+name+bullets table on the target slide:" & vbCrLf
    s = s & "    {""type"":""build_image_grid_table"",""slide"":2,""ref_name"":""apps_tbl""," & vbCrLf
    s = s & "     ""pos"":{""left"":30,""top"":60,""width"":900,""height"":480}," & vbCrLf
    s = s & "     ""image_col"":1,""desc_col"":2,""name_position"":""bottom""," & vbCrLf
    s = s & "     ""name_strip_pt"":30,""image_pad_pt"":6," & vbCrLf
    s = s & "     ""col1_width_pt"":280,""col2_width_pt"":620," & vbCrLf
    s = s & "     ""name_font"":{""size"":12,""bold"":true,""color"":""#15283C""}," & vbCrLf
    s = s & "     ""desc_font"":{""size"":10,""color"":""#333333""}," & vbCrLf
    s = s & "     ""rows"":[" & vbCrLf
    s = s & "       {""name"":""Aerospace"",""image_path"":""C:\\\\path\\\\img_003.jpg"",""bullets"":[""Lightweight fasteners"",""FAA compliant""]}," & vbCrLf
    s = s & "       {""name"":""Automotive"",""image_path"":""C:\\\\path\\\\img_007.jpg"",""bullets"":[""High-volume assembly"",""Vibration-resistant""]}" & vbCrLf
    s = s & "     ]}" & vbCrLf
    s = s & "  Each row's image_path is a downloaded file from the picker slide; the row's" & vbCrLf
    s = s & "  ""image_url"" alternative downloads on the fly." & vbCrLf
    s = s & "  Image is overlaid as a separate shape on top of the cell - the cell text" & vbCrLf
    s = s & "  ""name"" sits at the bottom (or top) of the same cell visually." & vbCrLf
    s = s & "  Standalone helper: {""type"":""download_image"",""url"":""..."",""dest_path"":""C:\\\\...""}" & vbCrLf & vbCrLf

    s = s & "SLIDE STRUCTURE:" & vbCrLf
    s = s & "  {""type"":""move_slide"",""from"":3,""to"":1}" & vbCrLf
    s = s & "  {""type"":""extract_slides"",""slide_indices"":[1,3,5],""output_path"":""C:\\path\\out.pptx""}" & vbCrLf
    s = s & "  {""type"":""import_slides_from_deck"",""source_path"":""C:\\path\\other.pptx"",""slide_indices"":[1,2],""target_position"":3}" & vbCrLf & vbCrLf

    s = s & "TABLES (existing):" & vbCrLf
    s = s & "  {""type"":""add_table_row"",""slide"":1,""shape_id"":5,""after_row"":2}" & vbCrLf
    s = s & "  {""type"":""delete_table_row"",""slide"":1,""shape_id"":5,""row"":3}" & vbCrLf
    s = s & "  {""type"":""add_table_col"",""slide"":1,""shape_id"":5,""after_col"":2}" & vbCrLf
    s = s & "  {""type"":""delete_table_col"",""slide"":1,""shape_id"":5,""col"":3}" & vbCrLf
    s = s & "  {""type"":""merge_cells"",""slide"":1,""shape_id"":5,""row_a"":1,""col_a"":1,""row_b"":1,""col_b"":3}" & vbCrLf & vbCrLf

    s = s & "GROUPS:" & vbCrLf
    s = s & "  {""type"":""group_shapes"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""ungroup"",""slide"":1,""shape_id"":12}" & vbCrLf & vbCrLf

    s = s & "CREATION FROM SCRATCH (use ref_name to chain styling):" & vbCrLf
    s = s & "  Use 'ref_name' to label a newly created shape, then reference it in" & vbCrLf
    s = s & "  follow-up actions via 'shape_name' (alternative to 'shape_id')." & vbCrLf
    s = s & "  This unlocks: create rect -> add text -> color -> reposition -> z-order." & vbCrLf & vbCrLf
    s = s & "  {""type"":""add_text_box"",""slide"":1,""text"":""Hello"",""ref_name"":""tb1""," & _
            """pos"":{""left"":50,""top"":50,""width"":200,""height"":40},""font_color"":""#FF0000"",""font_size"":14,""font_bold"":true}" & vbCrLf
    s = s & "  {""type"":""add_shape"",""slide"":1,""kind"":""circle"",""ref_name"":""badge_B""," & _
            """pos"":{""left"":30,""top"":100,""width"":30,""height"":30},""fill"":""#1F4E79"",""stroke"":null," & _
            """text"":""B"",""font_color"":""#FFFFFF"",""font_size"":11,""font_bold"":true}" & vbCrLf
    s = s & "  {""type"":""z_order"",""slide"":1,""shape_name"":""badge_B"",""order"":""front""}" & vbCrLf
    s = s & "  {""type"":""duplicate_shape"",""slide"":1,""shape_id"":5,""left"":300,""top"":100,""ref_name"":""dup1""}" & vbCrLf
    s = s & "  {""type"":""copy_formatting"",""slide"":1,""source_shape_id"":5,""target_shape_id"":7}" & vbCrLf & vbCrLf

    s = s & "EXPANDED SHAPE KINDS for add_shape (kind field):" & vbCrLf
    s = s & "  Basic: rect, rrect, oval/circle, diamond, triangle, right_triangle," & vbCrLf
    s = s & "    parallelogram, trapezoid, hexagon, octagon, pentagon, cross/plus, capsule" & vbCrLf
    s = s & "  Arrows: arrow/right_arrow, left_arrow, up_arrow, down_arrow," & vbCrLf
    s = s & "    double_arrow/left_right_arrow, up_down_arrow, quad_arrow," & vbCrLf
    s = s & "    curved_right_arrow, curved_left_arrow, striped_arrow, notched_arrow" & vbCrLf
    s = s & "  Process: chevron, chevron_pentagon" & vbCrLf
    s = s & "  Callouts: callout_rect, callout_rrect, callout_oval, callout_cloud," & vbCrLf
    s = s & "    callout_line1, callout_line2" & vbCrLf
    s = s & "  Stars: star4/5/6/8/16, ribbon_up, ribbon_down" & vbCrLf
    s = s & "  Misc: donut, block_arc, brace_left/right, bracket_left/right, plaque, cloud" & vbCrLf
    s = s & "  Numeric escape: ""kind"":""mso_52"" or ""kind"":""52"" passes raw msoAutoShapeType" & vbCrLf & vbCrLf

    s = s & "CONNECTORS (org chart / process flow):" & vbCrLf
    s = s & "  {""type"":""add_connector"",""slide"":1,""from_shape_name"":""box1"",""to_shape_name"":""box2""," & _
            """kind"":""elbow"",""arrow_end"":""filled"",""arrow_start"":""none""," & _
            """from_point"":""bottom"",""to_point"":""top""," & _
            """color"":""#000000"",""weight_pt"":1.5,""arrow_size"":""medium"",""dash_style"":""solid""}" & vbCrLf
    s = s & "  (from_shape_name/to_shape_name = ref_name used when creating the shape; numeric from_shape_id/to_shape_id also accepted)" & vbCrLf
    s = s & "  - kind: straight | elbow | curved" & vbCrLf
    s = s & "  - arrow_end / arrow_start: filled | open | diamond | oval | none" & vbCrLf
    s = s & "  - from_point / to_point: top | right | bottom | left | auto (or 1-8 numeric)" & vbCrLf
    s = s & "  - arrow_size: small | medium | large" & vbCrLf
    s = s & "  - dash_style: solid | dash | dot | round_dot | dash_dot | long_dash | long_dash_dot" & vbCrLf & vbCrLf

    s = s & "TABLE CREATION + STYLING:" & vbCrLf
    s = s & "  {""type"":""add_table"",""slide"":1,""rows"":4,""cols"":3,""ref_name"":""t1""," & _
            """pos"":{""left"":50,""top"":100,""width"":600,""height"":200}}" & vbCrLf
    s = s & "  {""type"":""set_table_col_width"",""slide"":1,""shape_id"":5,""col"":1,""width_pt"":250}" & vbCrLf
    s = s & "  {""type"":""set_table_row_height"",""slide"":1,""shape_id"":5,""row"":1,""height_pt"":40}" & vbCrLf
    s = s & "  {""type"":""set_cell_border"",""slide"":1,""shape_id"":5,""row"":1,""col"":1,""side"":""all""," & _
            """color"":""#000000"",""weight_pt"":1.0,""visible"":true}" & vbCrLf
    s = s & "    - side: top | left | bottom | right | diag_down | diag_up | all" & vbCrLf
    s = s & "  {""type"":""set_cell_text_align"",""slide"":1,""shape_id"":5,""row"":1,""col"":1," & _
            """h_align"":""center"",""v_align"":""middle""}" & vbCrLf
    s = s & "    - h_align: left | center | right; v_align: top | middle | bottom" & vbCrLf
    s = s & "  {""type"":""set_cell_fill"",""slide"":1,""shape_id"":5,""row"":1,""col"":1,""color"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""apply_table_style"",""slide"":1,""shape_id"":5,""style_id"":""medium_style_2_accent1""}" & vbCrLf
    s = s & "    - style_id: named (no_style_no_grid, no_style_with_grid, themed_style_1[_accent1..2]," & vbCrLf
    s = s & "      themed_style_2[_accent1], medium_style_2[_accent1..2], dark_style_2[_accent1]," & vbCrLf
    s = s & "      light_style_1[_accent1], light_style_2[_accent1]) or {GUID} pass-through" & vbCrLf & vbCrLf

    s = s & "NATIVE CHARTS (Shape.HasChart=True only; pasted images skipped):" & vbCrLf
    s = s & "  {""type"":""set_chart_type"",""slide"":1,""shape_id"":4,""value"":""barClustered""}" & vbCrLf
    s = s & "  {""type"":""set_chart_title"",""slide"":1,""shape_id"":4,""value"":""Q3 Revenue"",""enabled"":true}" & vbCrLf
    s = s & "  {""type"":""set_chart_axis_title"",""slide"":1,""shape_id"":4,""axis"":""x"",""value"":""Quarter""}" & vbCrLf
    s = s & "  {""type"":""set_chart_legend_position"",""slide"":1,""shape_id"":4,""value"":""bottom""}" & vbCrLf
    s = s & "  {""type"":""set_series_color"",""slide"":1,""shape_id"":4,""series_index"":1,""value"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""set_series_values"",""slide"":1,""shape_id"":4,""series_index"":1,""values"":[10.5,12.1,15.3,18.7]}" & vbCrLf
    s = s & "  {""type"":""set_chart_categories"",""slide"":1,""shape_id"":4,""categories"":[""Q1"",""Q2"",""Q3"",""Q4""]}" & vbCrLf
    s = s & "  {""type"":""set_series_name"",""slide"":1,""shape_id"":4,""series_index"":1,""value"":""Revenue""}" & vbCrLf & vbCrLf

    s = s & "SLIDE-LEVEL + POLISH:" & vbCrLf
    s = s & "  {""type"":""set_slide_background_color"",""slide"":1,""color"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""insert_slide_number"",""slide"":1,""ref_name"":""pn""," & _
            """pos"":{""left"":600,""top"":500,""width"":80,""height"":25},""font_color"":""#888888"",""font_size"":10}" & vbCrLf
    s = s & "  {""type"":""set_run_strikethrough"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":0,""value"":true}" & vbCrLf & vbCrLf

    s = s & "RULES:" & vbCrLf
    s = s & "- Use only shape_ids that exist in the snapshot. Do not invent ids." & vbCrLf
    s = s & "- Slide / row / col / series / position numbers are 1-based." & vbCrLf
    s = s & "- paragraph_index is 0-based to match the snapshot's paragraphs[].index." & vbCrLf
    s = s & "- after_paragraph_index = -1 means insert at top." & vbCrLf
    s = s & "- Colors are #RRGGBB hex strings." & vbCrLf
    s = s & "- Lengths in points." & vbCrLf
    s = s & "- Booleans are JSON true / false (lowercase, no quotes)." & vbCrLf
    s = s & "- For 'every X with property P -> Y' requests: enumerate matching shape_ids" & vbCrLf
    s = s & "  in the snapshot and emit one action per match (or use a *_match helper)." & vbCrLf
    s = s & "- For 'rebuild this slide': use clear_slide first, then a sequence of" & vbCrLf
    s = s & "  add_shape / set_text / move_shape actions to populate." & vbCrLf
    s = s & "- File paths must be absolute and use double backslashes in JSON strings." & vbCrLf
    s = s & "- One field name per action - never substitute aliases." & vbCrLf & vbCrLf

    s = s & "STRICT JSON OUTPUT RULES (parser is intolerant - violations cause errors):" & vbCrLf
    s = s & "- Output ONLY ONE JSON object. The first character of your reply MUST be" & vbCrLf
    s = s & "  ``{`` and the last character MUST be ``}``." & vbCrLf
    s = s & "- NO prose, NO explanation, NO markdown code fences (no ```json or ```)" & vbCrLf
    s = s & "  before, after, or inside the JSON." & vbCrLf
    s = s & "- NO JavaScript-style comments inside the JSON. Forbidden: ``// ...`` and" & vbCrLf
    s = s & "  ``/* ... */``. Strict JSON does not allow comments." & vbCrLf
    s = s & "- NEVER emit the JSON twice or include both a ``with-comments`` draft and a" & vbCrLf
    s = s & "  ``cleaned`` version. One JSON, once, alone." & vbCrLf
    s = s & "- Use straight ASCII double-quotes (`""`) - never smart quotes." & vbCrLf
    s = s & "- Escape backslashes as `\\` and double-quotes as `\""` inside string values." & vbCrLf
    s = s & "- Trailing commas are forbidden (e.g. `[1, 2, 3,]` is invalid)." & vbCrLf & vbCrLf

    s = s & "SCOPE GUIDANCE - get the slide number RIGHT:" & vbCrLf
    s = s & "- For find_replace_text and find_replace_regex, ``scope`` must reference the" & vbCrLf
    s = s & "  slide where the text actually lives. Read the snapshot to find the correct" & vbCrLf
    s = s & "  slide_number BEFORE writing the action." & vbCrLf
    s = s & "- If you are editing a single specific slide, use ``""scope"":""slide:N""`` with" & vbCrLf
    s = s & "  N being that slide's 1-based index in the snapshot." & vbCrLf
    s = s & "- If you want a sweep across the whole presentation, use ``""scope"":""deck""``." & vbCrLf
    s = s & "- Never assume slide:1 unless slide 1 is genuinely the target." & vbCrLf & vbCrLf

    s = s & "AUTONOMOUS EDIT POLICY - do not require the user to spell out details:" & vbCrLf
    s = s & "The user's request will often be a high-level instruction such as ""rebuild" & vbCrLf
    s = s & "this slide for Eli Lilly"" or ""swap the company to Pfizer"". You are expected" & vbCrLf
    s = s & "to read the snapshot, understand the slide's purpose, and produce all the" & vbCrLf
    s = s & "actions needed. Apply these rules without being told:" & vbCrLf & vbCrLf

    s = s & "1. Pictures cannot be modified from this tool. Any shape with" & vbCrLf
    s = s & "   ``""type"":""picture""`` MUST be skipped - do NOT emit any action that" & vbCrLf
    s = s & "   targets a picture shape's id. The user will replace logos manually." & vbCrLf
    s = s & "2. Skip slide-template chrome unless the user explicitly asks otherwise:" & vbCrLf
    s = s & "   slide-number placeholders (shape_name like ""Slide Number Placeholder"")," & vbCrLf
    s = s & "   ALL-CAPS section header banners, dated footnotes (""(1) Financial data" & vbCrLf
    s = s & "   as of ..."")  and any shape whose ``type`` is ``other`` containing only" & vbCrLf
    s = s & "   a page number." & vbCrLf
    s = s & "3. Action priority - prefer surgical over destructive:" & vbCrLf
    s = s & "     a. find_replace_text  (preserves all formatting; safe for unique" & vbCrLf
    s = s & "        substrings)" & vbCrLf
    s = s & "     b. set_run_text       (when a paragraph has multiple runs and you" & vbCrLf
    s = s & "        only need to change one run; preserves bold/color/size of that" & vbCrLf
    s = s & "        run AND its neighbours)" & vbCrLf
    s = s & "     c. set_paragraph_text (whole-paragraph swap; collapses run-level" & vbCrLf
    s = s & "        formatting INSIDE that paragraph - acceptable for paragraphs" & vbCrLf
    s = s & "        that already have a single run)" & vbCrLf
    s = s & "     d. set_text           (LAST RESORT - destroys all formatting on the" & vbCrLf
    s = s & "        whole shape; use only for plain-text shapes with single-style" & vbCrLf
    s = s & "        content or when the user explicitly says ""replace everything"")" & vbCrLf
    s = s & "4. Run-aware editing for mixed-format paragraphs (CRITICAL):" & vbCrLf
    s = s & "   If the snapshot shows a paragraph with multiple runs of different" & vbCrLf
    s = s & "   formatting, you MUST use set_run_text per run, not set_paragraph_text." & vbCrLf
    s = s & "   set_paragraph_text on a multi-run paragraph collapses all bold/color" & vbCrLf
    s = s & "   formatting inside that paragraph." & vbCrLf
    s = s & "   Concrete pattern - product bullets like" & vbCrLf
    s = s & "       p0: [bold]""Enbrel""  + [plain]"" - To treat plaque psoriasis...""" & vbCrLf
    s = s & "   should be edited as TWO actions to preserve the bold drug name:" & vbCrLf
    s = s & "       {""type"":""set_run_text"",""slide"":N,""shape_id"":12,""paragraph_index"":0,""run_index"":0,""value"":""Mounjaro""}" & vbCrLf
    s = s & "       {""type"":""set_run_text"",""slide"":N,""shape_id"":12,""paragraph_index"":0,""run_index"":1,""value"":"" - To treat type 2 diabetes...""}" & vbCrLf
    s = s & "5. Slide inference - never silently default to slide 1:" & vbCrLf
    s = s & "   - If the user names a slide explicitly (""slide 33"", ""the Amgen" & vbCrLf
    s = s & "     profile""), use that one." & vbCrLf
    s = s & "   - If the user says ""this slide"" or doesn't specify, scan the snapshot" & vbCrLf
    s = s & "     and pick the slide whose content matches the user's request." & vbCrLf
    s = s & "   - Slide 1 is only the right answer when the snapshot has exactly one" & vbCrLf
    s = s & "     slide OR when slide 1 is demonstrably the target." & vbCrLf
    s = s & "6. Tab-aligned label/value rows (e.g. ``""Headquarters:\tCalifornia""``)" & vbCrLf
    s = s & "   are TWO logical fields separated by a tab. When updating just the" & vbCrLf
    s = s & "   value, use find_replace_text on the value substring (e.g. find" & vbCrLf
    s = s & "   ``""California""``, replace ``""Indianapolis""``) - do not rewrite the" & vbCrLf
    s = s & "   whole paragraph." & vbCrLf
    s = s & "7. ALWAYS include the latest publicly-available data. You have web access" & vbCrLf
    s = s & "   and your knowledge is current. Do NOT use stale training data, do NOT" & vbCrLf
    s = s & "   skip a field with ""I'm not sure"", and do NOT default to numbers from" & vbCrLf
    s = s & "   the snapshot. The user expects fresh, accurate values for every field" & vbCrLf
    s = s & "   you touch (financials, headcount, executives, recent acquisitions, etc.)." & vbCrLf
    s = s & "8. Longest-first order: when emitting multiple find_replace_text actions" & vbCrLf
    s = s & "   whose ``find`` strings overlap (e.g. ""Amgen Inc."" and ""Amgen""), put" & vbCrLf
    s = s & "   the LONGER substring FIRST so the longer phrase matches and is" & vbCrLf
    s = s & "   replaced before the shorter one consumes its prefix. Example:" & vbCrLf
    s = s & "       1) ""Amgen Inc."" -> ""Eli Lilly and Company""" & vbCrLf
    s = s & "       2) ""Amgen""      -> ""Eli Lilly and Company""" & vbCrLf
    s = s & "9. Trailing terminators - DO NOT include a trailing ``\r`` (or ``\n``)" & vbCrLf
    s = s & "   in any ``value``, ``find``, or ``replace`` string. The paragraph" & vbCrLf
    s = s & "   terminator is implicit; including one will insert a spurious blank" & vbCrLf
    s = s & "   row. Even though the snapshot shows ``""text"":""Foo\r""`` for each" & vbCrLf
    s = s & "   paragraph, your action's value should be just ``""Foo""``." & vbCrLf
    s = s & "10. List length changes - if the new list has the SAME number of items" & vbCrLf
    s = s & "    as the original, prefer set_run_text per (paragraph_index, run_index)." & vbCrLf
    s = s & "    If the count differs, first emit delete_paragraph from highest index" & vbCrLf
    s = s & "    DOWN to lowest, then add_paragraph in order. Never just append new" & vbCrLf
    s = s & "    paragraphs without removing the old - that produces a doubled list." & vbCrLf
    s = s & "12. Creation pattern - when adding NEW shapes/text boxes/tables that need" & vbCrLf
    s = s & "    follow-up styling, ALWAYS assign a unique ""ref_name"" on the create" & vbCrLf
    s = s & "    action, then reference it via ""shape_name"" in subsequent actions in" & vbCrLf
    s = s & "    the same batch. Example: badge pattern (rect + circle overlay):" & vbCrLf
    s = s & "        {""type"":""add_shape"",""slide"":1,""kind"":""rect"",""ref_name"":""r1"",...}" & vbCrLf
    s = s & "        {""type"":""add_shape"",""slide"":1,""kind"":""circle"",""ref_name"":""c1"",...}" & vbCrLf
    s = s & "        {""type"":""z_order"",""slide"":1,""shape_name"":""c1"",""order"":""front""}" & vbCrLf
    s = s & "    Inline text+font params on add_shape/add_text_box are preferred over" & vbCrLf
    s = s & "    chained set_text/set_font_color follow-ups (fewer round trips)." & vbCrLf
    s = s & "13. Org chart pattern - to wire connectors between shapes you just created," & vbCrLf
    s = s & "    use add_connector with ""from_point""/""to_point"" for clean routing:" & vbCrLf
    s = s & "    boxes stacked vertically -> from_point=""bottom"", to_point=""top""." & vbCrLf
    s = s & "    Connectors require shape_id (not shape_name) for from/to. Plan order:" & vbCrLf
    s = s & "    add_shape...ref_name=""parent"" -> add_shape...ref_name=""child1"" ->" & vbCrLf
    s = s & "    THEN look up ids from snapshot or use shape_name on add_connector if" & vbCrLf
    s = s & "    you need to forward-reference (currently from/to require ids; emit a" & vbCrLf
    s = s & "    second pass after creation if needed)." & vbCrLf
    s = s & "11. Overflow guard - new content is often longer than what it replaces" & vbCrLf
    s = s & "    (e.g. ""Chairman & CEO"" -> ""Chair, President & CEO""). Always APPEND" & vbCrLf
    s = s & "    one or two enable_text_shrink_for_overflow actions at the END of" & vbCrLf
    s = s & "    your action list so PowerPoint auto-shrinks any text that no longer" & vbCrLf
    s = s & "    fits its frame. Title shapes are skipped by default. Schema:" & vbCrLf
    s = s & "        {""type"":""enable_text_shrink_for_overflow"",""scope"":""slide:N""}" & vbCrLf
    s = s & "    or {""type"":""enable_text_shrink_for_overflow"",""scope"":""deck""}" & vbCrLf
    s = s & "    Use slide:N when you only edited one slide; deck for multi-slide" & vbCrLf
    s = s & "    edits. This action MUST be last so it sees the final content." & vbCrLf

    s = s & vbCrLf & "ADDITIONAL ACTIONS (use as needed):" & vbCrLf
    s = s & "  {""type"":""set_text_autofit"",""slide"":1,""shape_id"":11,""mode"":""shrink""}" & vbCrLf
    s = s & "  {""type"":""enable_text_shrink_for_overflow"",""scope"":""slide:1""}" & vbCrLf
    s = s & "  {""type"":""enable_text_shrink_for_overflow"",""scope"":""deck"",""include_titles"":""false""}" & vbCrLf

    PromptTemplate = s
End Function

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim scope As String: scope = modExportSnapshot.g_SnapshotScope
    txtSnapshot.Text = modExportSnapshot.BuildSnapshotJson(scope)
    If Err.Number <> 0 Then
        txtSnapshot.Text = "ERROR: " & Err.Description
        Err.Clear
    End If
    Dim suffix As String
    If LCase(Trim(scope)) = "active" Then
        Dim idx As Long: idx = 1
        On Error Resume Next
        idx = ActiveWindow.View.Slide.SlideIndex
        On Error GoTo 0
        suffix = "  (ACTIVE SLIDE ONLY: slide " & idx & ")"
    ElseIf Len(Trim(scope)) > 0 And LCase(Trim(scope)) <> "all" Then
        suffix = "  (SCOPE: " & scope & ")"
    Else
        suffix = "  (ALL SLIDES)"
    End If
    Me.Caption = "Decko.ai - Export Snapshot" & suffix
End Sub

' User can flip scope from the form without closing it.
' Run via VBE Immediate window or wire to a button if you add one to the .frx.
Public Sub RebuildSnapshot(scope As String)
    modExportSnapshot.g_SnapshotScope = scope
    On Error Resume Next
    txtSnapshot.Text = modExportSnapshot.BuildSnapshotJson(scope)
    If Err.Number <> 0 Then
        txtSnapshot.Text = "ERROR: " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub btnCopySnapshot_Click()
    CopyToClipboard txtSnapshot.Text
    lblStatus.Caption = "Snapshot copied to clipboard."
End Sub

Private Sub btnCopyWithTemplate_Click()
    Dim payload As String
    payload = Replace(PromptTemplate(), "{snapshot}", txtSnapshot.Text)
    CopyToClipboard payload
    lblStatus.Caption = "Snapshot + prompt template copied to clipboard."
End Sub

Private Sub btnSaveTxt_Click()
    Dim deckPath As String: deckPath = ActivePresentation.FullName
    Dim ts As String: ts = Format(Now, "yyyy-mm-dd_hhnnss")
    Dim outPath As String
    outPath = deckPath & "_snapshot_" & ts & ".txt"

    Dim f As Integer: f = FreeFile
    Open outPath For Output As #f
    Print #f, txtSnapshot.Text
    Close #f
    lblStatus.Caption = "Saved: " & outPath
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub CopyToClipboard(s As String)
    Dim doObj As Object
    Set doObj = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    doObj.SetText s
    doObj.PutInClipboard
End Sub


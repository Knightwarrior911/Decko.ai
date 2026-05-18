"""Emit slide1.actions.json — Citi 'Financial results overview' recreation.

Slide = 960x540pt 16:9. Built as deck slide 1 (blank), default slide deleted last.
"""
import json
from pathlib import Path

OUT = Path(__file__).resolve().parent / "slide1.actions.json"

NAVY = "#0F1632"; CITI = "#255BE3"; MIDBLUE = "#1F4E79"
CYAN = "#00B0F0"; LTBLUE = "#73C2FC"; GRAY = "#A6A6A6"
HDRBLUE = "#255BE3"; BANDLT = "#DCE6F7"; INK = "#0F1632"; SUB = "#7C7C7C"

A = []

# slide 1 blank
A.append({"type": "add_slide", "position": 1, "layout_index": 6})

# title + chrome
A.append({"type": "add_text_box", "slide": 1, "text": "Financial results overview",
          "pos": {"left": 24, "top": 14, "width": 700, "height": 30},
          "font_color": CITI, "font_size": 22, "font_bold": True})
A.append({"type": "add_line", "slide": 1, "x1": 24, "y1": 50, "x2": 936, "y2": 50,
          "color": CITI, "weight_pt": 1.5})
A.append({"type": "add_text_box", "slide": 1, "text": "6",
          "pos": {"left": 916, "top": 518, "width": 28, "height": 16},
          "font_color": SUB, "font_size": 9, "h_align": "right"})
A.append({"type": "add_text_box", "slide": 1,
          "text": "Note: Totals may not sum due to rounding. All footnotes are presented starting on Slide 28.",
          "pos": {"left": 24, "top": 520, "width": 640, "height": 14},
          "font_color": SUB, "font_size": 7})

# ---- LEFT: Financial Results table ----
ROWS = [
    ("Net Interest Income", "15,741", "-", "12%"),
    ("Non-Interest Revenue", "8,892", "111%", "17%"),
    ("Total Revenues", "24,633", "24%", "14%"),
    ("Expenses", "14,311", "3%", "7%"),
    ("NCLs", "2,208", "1%", "(10)%"),
    ("ACL Build and Other(1)", "597", "NM", "126%"),
    ("Provision for Credit Losses", "2,805", "26%", "3%"),
    ("EBT", "7,517", "97%", "38%"),
    ("Income Taxes", "1,578", "23%", "18%"),
    ("Net Income", "5,785", "134%", "42%"),
    ("Net Income to Common(2)", "5,442", "151%", "44%"),
    ("Diluted EPS", "$3.06", "157%", "56%"),
    ("Efficiency Ratio (Δ in bps)", "58.1%", "(1,150)", "(410)"),
    ("ROE", "11.5%", "", ""),
    ("RoTCE(d) (Δ in bps)", "13.1%", "800", "400"),
    ("CET1 Capital Ratio(c)", "12.7%", "", ""),
    ("Memo:", "", "", ""),
    ("NII ex-Markets (g)", "12,944", "-", "7%"),
    ("NIR ex-Markets (h)", "4,443", "88%", "29%"),
]
HEADERS = ("($ in MM, except EPS)", "1Q26", "% Δ QoQ", "% Δ YoY")
BOLD_ROWS = {"Total Revenues", "EBT", "Net Income", "Net Income to Common(2)", "Diluted EPS"}

nrows = 2 + len(ROWS)  # title band + header + data
TL, TT, TW, TH = 24, 58, 470, 454
A.append({"type": "add_table", "slide": 1, "rows": nrows, "cols": 4,
          "pos": {"left": TL, "top": TT, "width": TW, "height": TH},
          "ref_name": "tbl"})
A.append({"type": "apply_table_style", "slide": 1, "shape_id": "tbl",
          "style_id": "no_style_no_grid"})
# shrink ALL cell text so 21 rows fit (row-level font — table shape has no text frame)
for r in range(1, nrows + 1):
    A.append({"type": "set_row_font_size", "slide": 1, "shape_id": "tbl", "row": r, "value": 8})
A.append({"type": "merge_cells", "slide": 1, "shape_id": "tbl",
          "row_a": 1, "col_a": 1, "row_b": 1, "col_b": 4})
A.append({"type": "set_cell_text", "slide": 1, "shape_id": "tbl", "row": 1, "col": 1,
          "value": "Financial Results"})
A.append({"type": "set_cell_fill", "slide": 1, "shape_id": "tbl", "row": 1, "col": 1, "color": CITI})
A.append({"type": "set_cell_text_align", "slide": 1, "shape_id": "tbl", "row": 1, "col": 1,
          "h_align": "center", "v_align": "middle"})

for c, h in enumerate(HEADERS, 1):
    A.append({"type": "set_cell_text", "slide": 1, "shape_id": "tbl", "row": 2, "col": c, "value": h})
    A.append({"type": "set_cell_fill", "slide": 1, "shape_id": "tbl", "row": 2, "col": c, "color": MIDBLUE})
    A.append({"type": "set_cell_text_align", "slide": 1, "shape_id": "tbl", "row": 2, "col": c,
              "h_align": "left" if c == 1 else "center", "v_align": "middle"})

for ri, row in enumerate(ROWS):
    r = ri + 3
    for c, txt in enumerate(row, 1):
        A.append({"type": "set_cell_text", "slide": 1, "shape_id": "tbl", "row": r, "col": c, "value": txt})
        A.append({"type": "set_cell_text_align", "slide": 1, "shape_id": "tbl", "row": r, "col": c,
                  "h_align": "left" if c == 1 else "right", "v_align": "middle"})
    if row[0] in BOLD_ROWS:
        for c in range(1, 5):
            A.append({"type": "set_cell_fill", "slide": 1, "shape_id": "tbl", "row": r, "col": c, "color": BANDLT})

# faint horizontal rules between data rows (clean Citi look)
for r in range(2, nrows + 1):
    for c in range(1, 5):
        A.append({"type": "set_cell_border", "slide": 1, "shape_id": "tbl", "row": r, "col": c,
                  "side": "bottom", "color": "#E3E3E3", "weight_pt": 0.5, "visible": True})

# header rows: white bold
for r in (1, 2):
    for c in range(1, 5):
        A.append({"type": "set_cell_font_color", "slide": 1, "shape_id": "tbl", "row": r, "col": c, "value": "#FFFFFF"})
        A.append({"type": "set_cell_font_bold", "slide": 1, "shape_id": "tbl", "row": r, "col": c, "value": True})
# bold/banded subtotal rows
for ri, row in enumerate(ROWS):
    if row[0] in BOLD_ROWS:
        r = ri + 3
        for c in range(1, 5):
            A.append({"type": "set_cell_font_bold", "slide": 1, "shape_id": "tbl", "row": r, "col": c, "value": True})

# tighten table
for c, w in enumerate((196, 92, 92, 90), 1):
    A.append({"type": "set_table_col_width", "slide": 1, "shape_id": "tbl", "col": c, "width_pt": w})
for r in range(1, nrows + 1):
    for c in range(1, 5):
        A.append({"type": "set_cell_padding", "slide": 1, "shape_id": "tbl",
                  "row": r, "col": c, "left": 3, "right": 3, "top": 1, "bottom": 1})

# ---- RIGHT TOP: highlights ----
A.append({"type": "add_text_box", "slide": 1, "text": "1Q26 Financial Overview Highlights",
          "pos": {"left": 508, "top": 58, "width": 428, "height": 22},
          "font_color": CITI, "font_size": 15, "font_bold": True})
A.append({"type": "add_line", "slide": 1, "x1": 508, "y1": 80, "x2": 936, "y2": 80,
          "color": CITI, "weight_pt": 1.0})
# REAL PowerPoint bullets: plain text paragraphs (no glyph chars) + per-paragraph
# set_indent_level + set_bullet_style + level color/size. (lvl, style, text)
L0C, L1C, L2C = INK, "#264F82", "#7C7C7C"
hl_paras = [
    (0, "disc",   "Revenues – Up 14% YoY, driven by continued growth in each of our businesses", L0C, 9),
    (1, "dash",   "NII up 12% YoY, driven by increases in each business and Legacy Franchises, partially offset by a decline in Corporate/Other", L1C, 8),
    (2, "square", "NII ex-Markets up 7%, driven by increases across businesses and Legacy Franchises", L2C, 8),
    (1, "dash",   "NIR up 17% YoY, driven by growth in each business and All Other", L1C, 8),
    (2, "square", "NIR ex-Markets up 29% YoY, driven by growth across businesses and All Other", L2C, 8),
    (0, "disc",   "Expenses – Up 7% YoY, driven by higher compensation and benefits (incl. severance and FX translation) and transactional and product servicing expenses", L0C, 9),
    (0, "disc",   "Provision for Credit Losses – Cost of $2.8 billion, primarily net credit losses in U.S. Consumer Cards plus a firmwide net ACL build of $597 million", L0C, 9),
    (0, "disc",   "RoTCE of 13.1%", L0C, 9),
]
A.append({"type": "add_text_box", "slide": 1, "text": "\n".join(p[2] for p in hl_paras),
          "pos": {"left": 508, "top": 86, "width": 428, "height": 250},
          "font_color": INK, "font_size": 9, "ref_name": "hl"})
for i, (lvl, style, _txt, col, sz) in enumerate(hl_paras):
    A.append({"type": "set_indent_level", "slide": 1, "shape_id": "hl", "paragraph_index": i, "value": lvl})
    A.append({"type": "set_bullet_style", "slide": 1, "shape_id": "hl", "paragraph_index": i, "value": style})
    A.append({"type": "set_paragraph_font_color", "slide": 1, "shape_id": "hl", "paragraph_index": i, "value": col})
    A.append({"type": "set_paragraph_font_size", "slide": 1, "shape_id": "hl", "paragraph_index": i, "value": sz})
A.append({"type": "set_text_autofit", "slide": 1, "shape_id": "hl", "mode": "shrink"})

# ---- RIGHT BOTTOM: Revenue by Segment chart ----
A.append({"type": "add_shape", "slide": 1, "kind": "rect",
          "pos": {"left": 508, "top": 330, "width": 428, "height": 20},
          "fill": CITI, "stroke": None, "text": "Revenue by Segment",
          "font_color": "#FFFFFF", "font_size": 11, "font_bold": True,
          "h_align": "center", "v_align": "middle"})
A.append({"type": "add_text_box", "slide": 1, "text": "($ in B)",
          "pos": {"left": 508, "top": 352, "width": 60, "height": 14},
          "font_color": SUB, "font_size": 8})
cats = ["1Q25", "4Q25", "1Q26"]
seg = [
    ("Services", [5.2, 6.3, 6.1], NAVY),
    ("Markets", [6.1, 4.6, 7.2], CITI),
    ("Banking", [2.8, 2.9, 3.1], MIDBLUE),
    ("Wealth", [1.5, 1.8, 1.8], CYAN),
    ("U.S. Consumer Cards", [4.6, 4.6, 4.8], LTBLUE),
    ("All Other (Managed Basis)", [1.5, -0.2, 1.7], GRAY),
]
totals = [round(sum(v[i] for _, v, _ in seg), 1) for i in range(3)]  # ~21.6/19.9/24.6
src_totals = ["21.6", "19.9", "24.6"]
series = [{"name": n, "values": v, "color": c} for n, v, c in seg]
series.append({"name": "Total", "values": totals, "color": "#FFFFFF"})
A.append({"type": "add_chart", "slide": 1, "chart_type": "columnstacked",
          "pos": {"left": 508, "top": 366, "width": 428, "height": 162},
          "categories": cats, "series": series,
          "show_legend": True, "value_format": "0.0", "ref_name": "revchart"})
A.append({"type": "set_chart_legend", "slide": 1, "shape_id": "revchart",
          "props": {"visible": True, "position": "bottom", "font_size": 7}})
A.append({"type": "set_chart_axis", "slide": 1, "shape_id": "revchart", "axis": "y",
          "props": {"visible": False}})
# segment data labels (white, inside)
for i in range(1, 7):
    A.append({"type": "set_chart_series", "slide": 1, "shape_id": "revchart",
              "series_index": i, "props": {"show_labels": True, "label_position": "center",
              "label_color": "#FFFFFF", "label_size": 7}})
# Total series -> invisible line on primary axis, labels above = stacked totals
A.append({"type": "set_chart_series", "slide": 1, "shape_id": "revchart", "series_index": 7,
          "props": {"chart_type": "line", "line_color": "#FFFFFF", "line_weight": 0.1,
                    "marker_style": "none", "show_labels": True, "label_position": "above",
                    "label_color": INK, "label_size": 9, "label_bold": True,
                    "custom_labels": src_totals, "hide_from_legend": True}})

OUT.write_text(json.dumps({"actions": A}, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"wrote {OUT}  ({len(A)} actions)")

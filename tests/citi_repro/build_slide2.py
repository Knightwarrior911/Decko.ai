"""Emit slide2.actions.json — Citi 'Quarterly expense trend and year-over-year
expense drivers'. Combo chart (stacked columns + efficiency line on secondary
axis + invisible total-label line) + small data table + 4 driver callout cards.
Slide = 960x540pt. Built as deck slide 1 (blank).
"""
import json
from pathlib import Path

OUT = Path(__file__).resolve().parent / "slide2.actions.json"
NAVY = "#0F1632"; CITI = "#255BE3"; MIDBLUE = "#1F4E79"
LTBLUE = "#73C2FC"; GRAY = "#A6A6A6"; INK = "#0F1632"; SUB = "#7C7C7C"
CARDBG = "#EEF2F8"; CARDLINE = "#C9D3E3"

A = []
A.append({"type": "add_slide", "position": 1, "layout_index": 6})

# title + chrome
A.append({"type": "add_text_box", "slide": 1,
          "text": "Quarterly expense trend and year-over-year expense drivers",
          "pos": {"left": 24, "top": 14, "width": 800, "height": 30},
          "font_color": CITI, "font_size": 22, "font_bold": True})
A.append({"type": "add_line", "slide": 1, "x1": 24, "y1": 50, "x2": 936, "y2": 50,
          "color": CITI, "weight_pt": 1.5})
A.append({"type": "add_text_box", "slide": 1, "text": "7",
          "pos": {"left": 916, "top": 518, "width": 28, "height": 16},
          "font_color": SUB, "font_size": 9, "h_align": "right"})
A.append({"type": "add_text_box", "slide": 1,
          "text": "Note: Totals may not sum due to rounding. All footnotes are presented starting on Slide 28.",
          "pos": {"left": 24, "top": 520, "width": 640, "height": 14},
          "font_color": SUB, "font_size": 7})

# ---- LEFT: Expense Overview combo chart ----
A.append({"type": "add_shape", "slide": 1, "kind": "rect",
          "pos": {"left": 24, "top": 58, "width": 560, "height": 20},
          "fill": CITI, "stroke": None, "text": "Expense Overview",
          "font_color": "#FFFFFF", "font_size": 11, "font_bold": True,
          "h_align": "center", "v_align": "middle"})
A.append({"type": "add_text_box", "slide": 1, "text": "($ in B)",
          "pos": {"left": 24, "top": 80, "width": 60, "height": 14},
          "font_color": SUB, "font_size": 8})

cats = ["1Q25", "2Q25", "3Q25", "4Q25", "1Q26"]
# stack order bottom -> top matches source
cols = [
    ("Compensation and Benefits & Restructuring", [7.5, 7.6, 7.5, 7.1, 8.4], NAVY),
    ("Transactional and Product Servicing", [1.1, 1.2, 1.1, 1.2, 1.2], LTBLUE),
    ("Technology / Communication", [2.4, 2.3, 2.3, 2.4, 2.3], MIDBLUE),
    ("Other Expenses ex-notable item", [2.5, 2.5, 2.7, 3.2, 2.4], CITI),
    ("Goodwill Impairment Charge", [0, 0, 0.7, 0, 0], GRAY),
]
eff = [62.2, 62.7, 64.7, 69.6, 58.1]
totals = ["$13.4", "$13.6", "$14.3", "$13.8", "$14.3"]
tot_vals = [13.4, 13.6, 14.3, 13.8, 14.3]

series = [{"name": n, "values": v, "color": c} for n, v, c in cols]
series.append({"name": "Reported Efficiency Ratio", "values": eff, "color": CITI})

# P4: ONE add_chart builds the whole combo — efficiency as a secondary-axis
# line (combo), and totals_label auto-adds the invisible total-line + "$13.4"
# labels excluded from the legend. (Was 4 set_chart_series + manual Total series.)
A.append({"type": "add_chart", "slide": 1, "chart_type": "columnstacked",
          "pos": {"left": 24, "top": 96, "width": 560, "height": 348},
          "categories": cats, "series": series,
          "show_legend": True, "value_format": '"$"0.0', "ref_name": "ec",
          "combo": [{"name": "Reported Efficiency Ratio",
                     "chart_type": "line", "axis_group": "secondary"}],
          "totals_label": True})
A.append({"type": "set_chart_legend", "slide": 1, "shape_id": "ec",
          "props": {"visible": True, "position": "bottom", "font_size": 7}})
# hide primary axis but KEEP its scale (visible:false collapses columns on a combo chart)
A.append({"type": "set_chart_axis", "slide": 1, "shape_id": "ec", "axis": "y",
          "props": {"min": 0, "max": 16, "tick_label_position": "none",
                    "line_visible": False}})
# column data labels (white, centered); P3: Goodwill (series 5) hides 0-quarters
for i in range(1, 6):
    p = {"show_labels": True, "label_position": "center",
         "label_color": "#FFFFFF", "label_size": 7}
    if i == 5:
        p["suppress_zero_labels"] = True
    A.append({"type": "set_chart_series", "slide": 1, "shape_id": "ec",
              "series_index": i, "props": p})
# series 6 = efficiency line styling + % labels
A.append({"type": "set_chart_series", "slide": 1, "shape_id": "ec", "series_index": 6,
          "props": {"line_color": CITI, "line_weight": 2.0,
                    "marker_style": "circle", "marker_size": 6, "marker_fill": CITI,
                    "show_labels": True, "label_position": "above",
                    "label_color": CITI, "label_size": 8, "label_bold": True,
                    "custom_labels": ["62.2%", "62.7%", "64.7%", "69.6%", "58.1%"]}})
# secondary axis: keep scale, hide visually
A.append({"type": "set_chart_axis", "slide": 1, "shape_id": "ec", "axis": "y2",
          "props": {"min": 50, "max": 72, "tick_label_position": "none",
                    "line_visible": False}})

# ---- small data table under chart ----
A.append({"type": "add_table", "slide": 1, "rows": 2, "cols": 6,
          "pos": {"left": 24, "top": 450, "width": 560, "height": 56}, "ref_name": "dt"})
A.append({"type": "apply_table_style", "slide": 1, "shape_id": "dt",
          "style_id": "no_style_no_grid"})
dt_rows = [
    ("Direct Staff (thousands)", ["229", "230", "227", "226", "224"]),
    ("Severance(3) ($B)", ["0.1", "0.4", "0.2", "0.1", "0.5"]),
]
for r, (lbl, vals) in enumerate(dt_rows, 1):
    A.append({"type": "set_cell_text", "slide": 1, "shape_id": "dt", "row": r, "col": 1, "value": lbl})
    A.append({"type": "set_cell_text_align", "slide": 1, "shape_id": "dt", "row": r, "col": 1,
              "h_align": "left", "v_align": "middle"})
    for c, v in enumerate(vals, 2):
        A.append({"type": "set_cell_text", "slide": 1, "shape_id": "dt", "row": r, "col": c, "value": v})
        A.append({"type": "set_cell_text_align", "slide": 1, "shape_id": "dt", "row": r, "col": c,
                  "h_align": "center", "v_align": "middle"})
for r in (1, 2):
    A.append({"type": "set_row_font_size", "slide": 1, "shape_id": "dt", "row": r, "value": 8})
A.append({"type": "set_cell_font_color", "slide": 1, "shape_id": "dt", "row": 1, "col": 1, "value": CITI})
A.append({"type": "set_cell_font_color", "slide": 1, "shape_id": "dt", "row": 2, "col": 1, "value": CITI})
A.append({"type": "set_table_col_width", "slide": 1, "shape_id": "dt", "col": 1, "width_pt": 160})
for c in range(2, 7):
    A.append({"type": "set_table_col_width", "slide": 1, "shape_id": "dt", "col": c, "width_pt": 80})

# ---- RIGHT: 1Q26 Expense Drivers callout cards ----
A.append({"type": "add_text_box", "slide": 1, "text": "1Q26 Expense Drivers",
          "pos": {"left": 600, "top": 58, "width": 336, "height": 22},
          "font_color": CITI, "font_size": 15, "font_bold": True})
A.append({"type": "add_line", "slide": 1, "x1": 600, "y1": 80, "x2": 936, "y2": 80,
          "color": CITI, "weight_pt": 1.0})

# body = list of bullet strings (NO glyph chars — real PowerPoint bullets applied)
cards = [
    ("Transactional and Product Servicing", "Up 11% YoY",
     ["Higher volumes in Markets, Services, U.S. Consumer Cards and Banking"]),
    ("Technology / Communication", "Down (2)% YoY",
     ["Reduction in technology contractors as a result of productivity",
      "Largely offset by technology charges and continued investments in technology and in the businesses to drive additional efficiencies and revenue growth"]),
    ("Other Expenses ex-notable item", "Down (5)% YoY",
     ["Lower legal expenses", "Lower professional services fees",
      "Partially offset by higher tax charges and deposit insurance costs"]),
    ("Compensation and Benefits & Restructuring", "Up 12% YoY",
     ["Higher severance charges",
      "Higher compensation associated with investments in Banking and Services",
      "Higher performance-related compensation",
      "Partially offset by productivity savings, stranded cost reduction and lower transformation expenses in Corporate/Other"]),
]
top = 84
H = [64, 100, 92, 140]
for (title, yoy, body), h in zip(cards, H):
    A.append({"type": "add_shape", "slide": 1, "kind": "round_rect",
              "pos": {"left": 600, "top": top, "width": 336, "height": h},
              "fill": CARDBG, "stroke": CARDLINE, "stroke_weight_pt": 0.75})
    A.append({"type": "add_text_box", "slide": 1, "text": f"{title}  —  {yoy}",
              "pos": {"left": 610, "top": top + 6, "width": 316, "height": 26},
              "font_color": NAVY, "font_size": 9, "font_bold": True})
    cb = f"cb_{top}"
    A.append({"type": "add_text_box", "slide": 1, "text": "\n".join(body),
              "pos": {"left": 610, "top": top + 30, "width": 316, "height": h - 34},
              "font_color": INK, "font_size": 8, "ref_name": cb})
    for i in range(len(body)):
        A.append({"type": "set_indent_level", "slide": 1, "shape_id": cb, "paragraph_index": i, "value": 0})
        A.append({"type": "set_bullet_style", "slide": 1, "shape_id": cb, "paragraph_index": i, "value": "disc"})
        A.append({"type": "set_paragraph_font_size", "slide": 1, "shape_id": cb, "paragraph_index": i, "value": 8})
    A.append({"type": "set_text_autofit", "slide": 1, "shape_id": cb, "mode": "shrink"})
    top += h + 8

OUT.write_text(json.dumps({"actions": A}, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"wrote {OUT}  ({len(A)} actions)")

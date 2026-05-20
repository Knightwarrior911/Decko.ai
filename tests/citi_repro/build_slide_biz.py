"""Slide A — "Five interconnected businesses driving strong 1Q26 performance"
(Citi Q1'26 earnings PDF, slide 4 / page index 3).

VP intent -> JSON action batch the LLM brain would emit. Emits
tests/citi_repro/slide_biz.actions.json. Data traces to SRC_bizmodel.txt.

Decorative rules/brackets are thin autoshape rects (not add_line) so the verify
loop does not flag orphan/zero-size connectors. Min font 8pt (verify gate).
Card bodies are round_rect callout cards. The 5 cards are linked by real
add_connector elements (the "interconnected" motif) anchored bottom->top so
their bounding box is non-degenerate. Title block is grouped (group_shapes).
"""
import json
from pathlib import Path

W = Path(__file__).resolve().parent
OUT = W / "slide_biz.actions.json"

BLUE = "#255BE3"
NAVY = "#0F1632"
GRAY = "#7C7C7C"
BAND = "#0F1632"
GREEN = "#2E7D32"

SLIDE_W = 960
M = 24
NCARDS = 5
GAP = 10
CARD_W = (SLIDE_W - 2 * M - (NCARDS - 1) * GAP) // NCARDS  # 174
CARD_TOP = 90
HDR_H = 22
BODY_BOTTOM = 456

CARDS = [
    dict(name="Services", hdr="#0F1632",
         stat="TTS: #1 Rank(1)\nTTS: gained ~100 bps of market share YoY(1)\n"
              "Securities Services: #1 in Direct Custody(2)",
         bullets="", yoy="17% YoY",
         series=[("TTS", [3.9, 4.6], "#0F1632"),
                 ("Securities Services", [1.3, 1.5], "#5B6B8C")],
         totals=["$5.2", "$6.1"], pol=True),
    dict(name="Markets", hdr="#1F4E79",
         stat="#3 Overall Rank (tied)(3)\nFixed Income: #2 Rank(3)\n"
              "Equities: #6 Rank (tied)(3)",
         bullets="Record prime balances(4), up more than 50% YoY",
         yoy="19% YoY",
         series=[("Fixed Income Markets", [4.6, 5.2], "#1F4E79"),
                 ("Equity Markets", [1.5, 2.1], "#6FA8DC")],
         totals=["$6.1", "$7.2"], pol=True),
    dict(name="Banking", hdr="#6B5D3E",
         stat="Investment Banking: #5 Rank(5)\nIB Fees: up 12% YoY",
         bullets="Record 1Q in Advisory(6)", yoy="11% YoY",
         series=[("Investment Banking", [1.8, 2.1], "#6B5D3E"),
                 ("Corporate Lending", [0.7, 0.8], "#A89368"),
                 ("Gain/(loss) on loan hedges", [0.3, 0.2], "#D8CBA8")],
         totals=["$2.8", "$3.1"], pol=False),
    dict(name="Wealth", hdr="#255BE3",
         stat="Private Bank (Global): #6 Rank, with APAC and MEA Private Bank: "
              "#3 Rank (tied)(7)",
         bullets="18% EBT Margin\nNet new investment asset(f) flows of ~$15 billion",
         yoy="20% YoY",
         series=[("Private Bank", [0.4, 0.4], "#255BE3"),
                 ("Wealth at Work / Citigold & Retail Banking", [1.1, 1.4],
                  "#7FA8E8")],
         totals=["$1.5", "$1.8"], pol=True),
    dict(name="U.S. Consumer Cards", hdr="#5B2C6F",
         stat="#3 Rank in U.S. Cards(8)",
         bullets="Spend volume up 6% YoY in General Purpose Credit Cards",
         yoy="4% YoY",
         series=[("U.S. Consumer Cards", [4.6, 4.8], "#5B2C6F")],
         totals=["$4.6", "$4.8"], pol=True),
]


def card_x(i):
    return M + i * (CARD_W + GAP)


def hrule(A, x1, x2, y, color, h=2):
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": x1, "top": y, "width": x2 - x1, "height": h},
              "fill": color, "stroke": None})


def vrule(A, x, y1, y2, color, w=2):
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": x, "top": y1, "width": w, "height": y2 - y1},
              "fill": color, "stroke": None})


def build():
    A = []
    A.append({"type": "add_slide", "position": 1, "layout_index": 6})
    A.append({"type": "add_text_box", "slide": 1,
              "text": "Five interconnected businesses driving strong 1Q26 "
                      "performance",
              "pos": {"left": M, "top": 12, "width": 820, "height": 28},
              "font_color": BLUE, "font_size": 22, "font_bold": True,
              "ref_name": "ttl"})
    A.append({"type": "add_text_box", "slide": 1, "text": "citi",
              "pos": {"left": 902, "top": 14, "width": 36, "height": 20},
              "font_color": BLUE, "font_size": 15, "font_bold": True,
              "h_align": "right", "ref_name": "wm"})
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": M, "top": 45, "width": SLIDE_W - 2 * M,
                      "height": 2}, "fill": BLUE, "stroke": None,
              "ref_name": "rule"})
    A.append({"type": "group_shapes", "slide": 1,
              "shape_ids": ["ttl", "wm", "rule"], "ref_name": "hdr_grp"})

    # Top grouping labels + bracket rules (thin rects, not connectors)
    spans = [("GLOBAL NETWORK", 0, 1), ("INTERCONNECTED", 2, 2),
             ("DIVERSIFIED", 3, 4)]
    for label, a, b in spans:
        x0 = card_x(a)
        x1 = card_x(b) + CARD_W
        cx = (x0 + x1) / 2
        A.append({"type": "add_text_box", "slide": 1, "text": label,
                  "pos": {"left": x0, "top": 50, "width": x1 - x0,
                          "height": 16},
                  "font_color": GRAY, "font_size": 10, "font_bold": True,
                  "h_align": "center"})
        hrule(A, x0 + 14, x1 - 14, 68, GRAY)
        vrule(A, cx - 1, 68, 75, GRAY)

    # Cards
    for i, c in enumerate(CARDS):
        x = card_x(i)
        A.append({"type": "add_shape", "slide": 1, "kind": "rect",
                  "ref_name": f"hdr{i}",
                  "pos": {"left": x, "top": CARD_TOP, "width": CARD_W,
                          "height": HDR_H},
                  "fill": c["hdr"], "stroke": None, "text": c["name"],
                  "font_color": "#FFFFFF", "font_size": 11, "font_bold": True,
                  "h_align": "center", "v_align": "middle"})
        A.append({"type": "add_shape", "slide": 1, "kind": "round_rect",
                  "pos": {"left": x, "top": CARD_TOP + HDR_H, "width": CARD_W,
                          "height": BODY_BOTTOM - (CARD_TOP + HDR_H)},
                  "fill": "#F5F7FA", "stroke": "#D7DEE8",
                  "stroke_weight_pt": 0.5})

        sref = f"stat{i}"
        A.append({"type": "add_text_box", "slide": 1, "text": c["stat"],
                  "pos": {"left": x + 6, "top": CARD_TOP + HDR_H + 6,
                          "width": CARD_W - 12, "height": 64},
                  "font_color": NAVY, "font_size": 8, "font_bold": True,
                  "font_italic": True, "ref_name": sref})
        A.append({"type": "set_text_autofit", "slide": 1, "shape_id": sref,
                  "mode": "shrink"})

        A.append({"type": "add_text_box", "slide": 1,
                  "text": c["yoy"] + "  ▲",
                  "pos": {"left": x + 6, "top": CARD_TOP + HDR_H + 70,
                          "width": CARD_W - 12, "height": 14},
                  "font_color": GREEN, "font_size": 9, "font_bold": True,
                  "h_align": "center"})

        chart_top = CARD_TOP + HDR_H + 86
        chart_h = 200
        cref = f"ch{i}"
        series = [{"name": n, "values": v, "color": col}
                  for (n, v, col) in c["series"]]
        t0 = round(sum(s[1][0] for s in c["series"]), 1)
        t1 = round(sum(s[1][1] for s in c["series"]), 1)
        series.append({"name": "Total", "values": [t0, t1],
                       "color": "#FFFFFF"})
        A.append({"type": "add_chart", "slide": 1,
                  "chart_type": "columnstacked",
                  "pos": {"left": x + 6, "top": chart_top,
                          "width": CARD_W - 12, "height": chart_h},
                  "categories": ["1Q25", "1Q26"], "series": series,
                  "show_legend": False, "value_format": "0.0",
                  "ref_name": cref})
        A.append({"type": "set_chart_axis", "slide": 1, "shape_id": cref,
                  "axis": "y", "props": {"visible": False}})
        for si in range(1, len(c["series"]) + 1):
            A.append({"type": "set_chart_series", "slide": 1,
                      "shape_id": cref, "series_index": si,
                      "props": {"show_labels": True,
                                "label_position": "center",
                                "label_color": "#FFFFFF", "label_size": 7}})
        A.append({"type": "set_chart_series", "slide": 1, "shape_id": cref,
                  "series_index": len(c["series"]) + 1,
                  "props": {"chart_type": "line", "line_color": "#FFFFFF",
                            "line_weight": 0.1, "marker_style": "none",
                            "show_labels": True, "label_position": "above",
                            "label_color": NAVY, "label_size": 9,
                            "label_bold": True, "custom_labels": c["totals"]}})

        if c["bullets"]:
            bref = f"bul{i}"
            A.append({"type": "add_text_box", "slide": 1, "text": c["bullets"],
                      "pos": {"left": x + 6, "top": chart_top + chart_h + 2,
                              "width": CARD_W - 12,
                              "height": BODY_BOTTOM - (chart_top + chart_h)
                              - 4},
                      "font_color": NAVY, "font_size": 8, "ref_name": bref})
            for p in range(c["bullets"].count("\n") + 1):
                A.append({"type": "set_indent_level", "slide": 1,
                          "shape_id": bref, "paragraph_index": p, "value": 0})
                A.append({"type": "set_bullet_style", "slide": 1,
                          "shape_id": bref, "paragraph_index": p,
                          "value": "disc"})
                A.append({"type": "set_paragraph_font_size", "slide": 1,
                          "shape_id": bref, "paragraph_index": p, "value": 8})
            A.append({"type": "set_text_autofit", "slide": 1,
                      "shape_id": bref, "mode": "shrink"})

        if c["pol"]:
            A.append({"type": "add_text_box", "slide": 1,
                      "text": "Positive operating leverage",
                      "pos": {"left": x, "top": BODY_BOTTOM + 2,
                              "width": CARD_W, "height": 14},
                      "font_color": GREEN, "font_size": 8, "font_bold": True,
                      "h_align": "center"})

    # interconnection connectors (bottom -> top so bbox is non-degenerate)
    for i in range(NCARDS - 1):
        A.append({"type": "add_connector", "slide": 1, "kind": "elbow",
                  "from_shape_name": f"hdr{i}", "to_shape_name": f"hdr{i+1}",
                  "from_point": "bottom", "to_point": "top",
                  "color": GRAY, "weight_pt": 1.0,
                  "arrow_start": "filled", "arrow_end": "filled"})

    # closing dark band
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": M, "top": 470, "width": SLIDE_W - 2 * M,
                      "height": 38},
              "fill": BAND, "stroke": None,
              "text": "Best quarterly revenue in a decade for the firm and "
                      "Markets, Wealth and U.S. Consumer Cards; Highest 1Q "
                      "revenues in Services in a decade(6)",
              "font_color": "#FFFFFF", "font_size": 11, "font_bold": True,
              "h_align": "center", "v_align": "middle"})

    A.append({"type": "add_footnote", "slide": 1,
              "text": "Note: Totals may not sum due to rounding. All "
                      "footnotes are presented starting on Slide 28.",
              "page_number": "4"})
    return A


def main():
    actions = build()
    OUT.write_text(json.dumps({"actions": actions}, ensure_ascii=False,
                              indent=1), encoding="utf-8")
    print(f"wrote {OUT}  ({len(actions)} actions)")


if __name__ == "__main__":
    main()

"""Slide B — "Banamex 24% stake - estimated financial impacts at closing(1)"
(Citi Q1'26 earnings PDF, slide 22 / page index 21).

VP intent -> JSON action batch. Emits tests/citi_repro/slide_banamex.actions.json.
Data traces to SRC_banamex.txt.

Layout: title + rule + citi wordmark; LEFT a column of lettered (a-e) navy
circle callouts each with a wrapped paragraph plus a closing unlettered note;
RIGHT a blue banner over a real native table (1 header + 13 data rows, 4 cols)
with light-blue banded bold subtotal rows and small lettered circle markers
linking rows back to the left notes; a full-width dark closing band; footnote
and page number.
"""
import json
from pathlib import Path

W = Path(__file__).resolve().parent
OUT = W / "slide_banamex.actions.json"

BLUE = "#255BE3"
NAVY = "#1F4E79"
DARK = "#0F1632"
GRAY = "#7C7C7C"
BANDFILL = "#DCE6F7"

SLIDE_W = 960
M = 24

# Left lettered callouts: (letter or "", text, est_height_pt)
LEFT = [
    ("a", "At closing, assets will increase reflecting the consideration "
          "received for the 24% stake based on a fixed price of ~MXN 43 "
          "billion or ~USD 2.5 billion(2)", 50),
    ("b", "The net loss on sale is recorded primarily in Additional Paid-in "
          "Capital (APIC) and is made up of the net of the sale consideration "
          "received and 24% of the Banamex U.S. GAAP Book Value", 50),
    ("c", "24% of the approximately $8.6B of total Banamex CTA (inclusive of "
          "amounts already reclassified to NCI), as of February 27, 2026, "
          "moves from AOCI to NCI - this causes a benefit to stockholders' "
          "equity since an unrealized loss is moving out of stockholders' "
          "equity", 66),
    ("d", "Therefore, Stockholder's equity increases by ~$1.7B\n"
          "As a reminder, the CTA to NCI reclass results in a temporary "
          "increase to stockholders' equity and CET1 until it reverses at "
          "deconsolidation", 56),
    ("e", "NCI increases by 24% of the Banamex U.S. GAAP Book Value largely "
          "offset by 24% of the CTA losses reclassed from AOCI", 42),
    ("", "Upon closing of all committed purchases, Citi will have sold 49% of "
         "Banamex. Therefore, 49% of all financial impacts of Banamex will "
         "impact the NCI(3)", 42),
]

# Table: header + 13 data rows. bold = banded subtotal row.
HEADER = ["$USD in B", "25% Stake(4)\nClose", "24% Stake", "Total"]
ROWS = [
    ("Total Estimated Sale Consideration", "2.3", "2.5", "4.8", False),
    ("Impact to Total Assets (Debit)", "2.3", "2.5", "4.8", True),
    ("Total Estimated Sale Consideration", "2.3", "2.5", "4.8", False),
    ("Less:  % of Banamex U.S. GAAP Book Value", "(2.9)", "(2.9)", "(5.8)", False),
    ("Loss on Sale", "(0.6)", "(0.4)", "(1.0)", True),
    ("Reclassification of Negative CTA(5)", "2.3", "2.1", "4.4", False),
    ("Impact to Stockholders' Equity (Credit)", "1.7", "1.7", "3.4", True),
    ("% of Banamex U.S. GAAP Book Value", "2.9", "2.9", "5.8", False),
    ("Reclassification of Negative CTA(5)", "(2.3)", "(2.1)", "(4.4)", False),
    ("Impact to NCI (Credit)", "0.6", "0.8", "1.4", True),
    ("Impact to Stockholders' Equity", "1.7", "1.7", "3.4", False),
    ("Impact to NCI", "0.6", "0.8", "1.4", False),
    ("Impact to Total Equity (Credit)", "2.3", "2.5", "4.8", True),
]

TBL_LEFT = 500
TBL_TOP = 96
TBL_W = SLIDE_W - M - TBL_LEFT  # 436
NROWS = len(ROWS) + 1
ROW_H = 26
TBL_H = NROWS * ROW_H

# which data-row (1-based, excludes header) gets a lettered marker on the table
ROW_MARKERS = {2: "a", 5: "b", 7: "c", 10: "d", 13: "e"}


def build():
    A = []
    A.append({"type": "add_slide", "position": 1, "layout_index": 6})
    A.append({"type": "add_text_box", "slide": 1,
              "text": "Banamex 24% stake – estimated financial impacts at "
                      "closing(1)",
              "pos": {"left": M, "top": 12, "width": 860, "height": 28},
              "font_color": BLUE, "font_size": 22, "font_bold": True})
    A.append({"type": "add_text_box", "slide": 1, "text": "citi",
              "pos": {"left": 902, "top": 14, "width": 36, "height": 20},
              "font_color": BLUE, "font_size": 15, "font_bold": True,
              "h_align": "right"})
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": M, "top": 45, "width": SLIDE_W - 2 * M,
                      "height": 2}, "fill": BLUE, "stroke": None})

    # LEFT lettered callouts
    y = 58
    LX = M
    LW = TBL_LEFT - M - 16
    for letter, text, h in LEFT:
        if letter:
            A.append({"type": "add_shape", "slide": 1, "kind": "oval",
                      "pos": {"left": LX, "top": y, "width": 16, "height": 16},
                      "fill": NAVY, "stroke": None, "text": letter,
                      "font_color": "#FFFFFF", "font_size": 9,
                      "font_bold": True, "h_align": "center",
                      "v_align": "middle"})
            tx, tw = LX + 22, LW - 22
        else:
            A.append({"type": "add_shape", "slide": 1, "kind": "oval",
                      "pos": {"left": LX + 2, "top": y + 4, "width": 5,
                              "height": 5}, "fill": DARK, "stroke": None})
            tx, tw = LX + 22, LW - 22
        ref = f"lc_{letter or 'z'}"
        A.append({"type": "add_text_box", "slide": 1, "text": text,
                  "pos": {"left": tx, "top": y, "width": tw, "height": h},
                  "font_color": DARK, "font_size": 9, "ref_name": ref})
        npar = text.count("\n") + 1
        for p in range(npar):
            A.append({"type": "set_indent_level", "slide": 1, "shape_id": ref,
                      "paragraph_index": p, "value": 0 if p == 0 else 1})
            A.append({"type": "set_bullet_style", "slide": 1, "shape_id": ref,
                      "paragraph_index": p,
                      "value": "none" if p == 0 else "dash"})
            A.append({"type": "set_paragraph_font_size", "slide": 1,
                      "shape_id": ref, "paragraph_index": p,
                      "value": 9 if p == 0 else 8})
        A.append({"type": "set_text_autofit", "slide": 1, "shape_id": ref,
                  "mode": "shrink"})
        y += h + 10

    # RIGHT banner
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": TBL_LEFT, "top": 58, "width": TBL_W,
                      "height": 32},
              "fill": BLUE, "stroke": None,
              "text": "Estimated Balance Sheet Impacts at Closing of the 24% "
                      "stake sale - subject to changes including FX",
              "font_color": "#FFFFFF", "font_size": 9, "font_bold": True,
              "h_align": "center", "v_align": "middle"})

    # RIGHT native table
    A.append({"type": "add_table", "slide": 1, "rows": NROWS, "cols": 4,
              "pos": {"left": TBL_LEFT, "top": TBL_TOP, "width": TBL_W,
                      "height": TBL_H}, "ref_name": "tbl"})
    A.append({"type": "apply_table_style", "slide": 1, "shape_id": "tbl",
              "style_id": "no_style_no_grid"})

    # header
    for ci, h in enumerate(HEADER, 1):
        A.append({"type": "set_cell_text", "slide": 1, "shape_id": "tbl",
                  "row": 1, "col": ci, "value": h})
        A.append({"type": "set_cell_text_align", "slide": 1,
                  "shape_id": "tbl", "row": 1, "col": ci,
                  "h_align": "left" if ci == 1 else "center",
                  "v_align": "middle"})
        A.append({"type": "set_cell_fill", "slide": 1, "shape_id": "tbl",
                  "row": 1, "col": ci, "color": BLUE})
        A.append({"type": "set_cell_font_color", "slide": 1,
                  "shape_id": "tbl", "row": 1, "col": ci, "value": "#FFFFFF"})
        A.append({"type": "set_cell_font_bold", "slide": 1,
                  "shape_id": "tbl", "row": 1, "col": ci, "value": True})

    # data rows
    for ri, (label, c2, c3, c4, bold) in enumerate(ROWS, start=2):
        vals = [label, c2, c3, c4]
        for ci, v in enumerate(vals, 1):
            A.append({"type": "set_cell_text", "slide": 1, "shape_id": "tbl",
                      "row": ri, "col": ci, "value": v})
            A.append({"type": "set_cell_text_align", "slide": 1,
                      "shape_id": "tbl", "row": ri, "col": ci,
                      "h_align": "left" if ci == 1 else "center",
                      "v_align": "middle"})
            if bold:
                A.append({"type": "set_cell_fill", "slide": 1,
                          "shape_id": "tbl", "row": ri, "col": ci,
                          "color": BANDFILL})
                A.append({"type": "set_cell_font_bold", "slide": 1,
                          "shape_id": "tbl", "row": ri, "col": ci,
                          "value": True})

    for r in range(1, NROWS + 1):
        A.append({"type": "set_row_font_size", "slide": 1, "shape_id": "tbl",
                  "row": r, "value": 8})
    # faint row separators (match source grid)
    for r in range(2, NROWS + 1):
        for ci in range(1, 5):
            A.append({"type": "set_cell_border", "slide": 1,
                      "shape_id": "tbl", "row": r, "col": ci,
                      "side": "bottom", "color": "#D7DEE8",
                      "weight_pt": 0.5, "visible": True})
    A.append({"type": "set_table_col_width", "slide": 1, "shape_id": "tbl",
              "col": 1, "width_pt": 214})
    for ci in (2, 3, 4):
        A.append({"type": "set_table_col_width", "slide": 1,
                  "shape_id": "tbl", "col": ci,
                  "width_pt": (TBL_W - 214) // 3})

    # lettered markers next to specific table rows
    mk_refs = []
    for drow, letter in ROW_MARKERS.items():
        cy = TBL_TOP + drow * ROW_H + (ROW_H - 14) / 2
        mref = f"mk_{letter}"
        mk_refs.append(mref)
        A.append({"type": "add_shape", "slide": 1, "kind": "oval",
                  "ref_name": mref,
                  "pos": {"left": TBL_LEFT - 18, "top": cy, "width": 14,
                          "height": 14}, "fill": NAVY, "stroke": None,
                  "text": letter, "font_color": "#FFFFFF", "font_size": 8,
                  "font_bold": True, "h_align": "center",
                  "v_align": "middle"})
    A.append({"type": "group_shapes", "slide": 1, "shape_ids": mk_refs,
              "ref_name": "mk_grp"})

    # closing dark band
    A.append({"type": "add_shape", "slide": 1, "kind": "rect",
              "pos": {"left": M, "top": 486, "width": SLIDE_W - 2 * M,
                      "height": 34},
              "fill": DARK, "stroke": None,
              "text": "Once Citi's voting stock ownership in Banamex is below "
                      "50%, Citi will deconsolidate the entity(6)",
              "font_color": "#FFFFFF", "font_size": 12, "font_bold": True,
              "h_align": "center", "v_align": "middle"})

    # footnote + page number
    A.append({"type": "add_text_box", "slide": 1,
              "text": "Note: As of February 27, 2026. All footnotes are "
                      "presented starting on Slide 28.",
              "pos": {"left": M, "top": 522, "width": 700, "height": 14},
              "font_color": GRAY, "font_size": 8})
    A.append({"type": "add_text_box", "slide": 1, "text": "22",
              "pos": {"left": 908, "top": 521, "width": 28, "height": 16},
              "font_color": GRAY, "font_size": 9, "h_align": "right"})
    return A


def main():
    actions = build()
    OUT.write_text(json.dumps({"actions": actions}, ensure_ascii=False,
                              indent=1), encoding="utf-8")
    print(f"wrote {OUT}  ({len(actions)} actions)")


if __name__ == "__main__":
    main()

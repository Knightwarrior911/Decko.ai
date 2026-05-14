"""Build think-cell feature overview — 2 slides.

Slide 1: 3-col × 2-row grid, 6 sections, 2×2 bullet layout per cell.
Slide 2: Left 67% blank | Right 33% condensed sidebar reference.

Output: test_decks/thinkcell_overview.pptx
"""

import json
import sys
import tempfile
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER   = REPO_ROOT / "PPT_AI_Editor.pptm"
OUT_DECK  = REPO_ROOT / "test_decks" / "thinkcell_overview.pptx"
PNG_DIR   = REPO_ROOT / "test_decks" / "thinkcell_overview_pngs"

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY   = "#152540"
TC_RED = "#D94F2B"
WHITE  = "#FFFFFF"
BODY   = "#1F2937"
GRAY   = "#6B7280"
LGRAY  = "#D1D5DB"
XGRAY  = "#F3F4F6"   # near-white for left blank area

H1, B1 = "#1A3A6A", "#EEF3FB"
H2, B2 = "#17553A", "#EBF5EF"
H3, B3 = "#8B4000", "#FBF0E8"
H4, B4 = "#3D1B70", "#F2EEF9"
H5, B5 = "#1A5C7A", "#E8F4F8"
H6, B6 = "#6B1818", "#F9EEEE"

W, H = 960, 540

# ── Pre-downloaded image paths ────────────────────────────────────────────────
_A = r"C:\Users\vinit\AppData\Local\Temp\tc_overview_d_add4b0\assets\page_httpswww.think-cell.comenproductthink-ce_20260514_122225"
_C = r"C:\Users\vinit\AppData\Local\Temp\tc_overview_d_add4b0\assets\page_httpswww.think-cell.comenessentialsthink_20260514_122249"
_L = r"C:\Users\vinit\AppData\Local\Temp\tc_overview_0gobl36g\assets\page_httpswww.think-cell.comenessentialsthink_20260514_124742"
_AI= r"C:\Users\vinit\AppData\Local\Temp\tc_overview_0gobl36g\assets\page_httpswww.think-cell.comenproductthink-ce_20260514_124807"

# ── Content ───────────────────────────────────────────────────────────────────
TITLE    = "Think-cell — Feature Overview"
PAGE_MSG = (
    "think-cell is a PowerPoint add-in used by all top 10 global consulting firms — "
    "covering chart building, Excel data links, timelines, and AI-powered editing."
)

SECTIONS = [
    {
        "title_short": "CHARTS & VISUALIZATIONS",
        "hdr": H1, "bg": B1,
        "img": str(Path(_A) / "img_013.jpg"),
        "sub": "Waterfall · Column · Scatter · Mekko · Combo",
        # 4 bullets as 2×2: [top-left, top-right, bottom-left, bottom-right]
        "bullets": [
            "Waterfall/bridge charts for EBITDA\n& P&L — totals auto-calculate",
            "40+ types: stacked/grouped bars,\nMekko, scatter, combo & more",
            "CAGR arrows, difference arrows\n& value lines auto-placed",
            "Labels & legends reposition for\nbest-practice chart layouts",
        ],
    },
    {
        "title_short": "EXCEL INTEGRATION",
        "hdr": H2, "bg": B2,
        "img": str(Path(_A) / "img_026.jpg"),
        "sub": "Live Links · Slide Workbooks · Data Tables",
        "bullets": [
            "Live Excel links: charts update\nwhen your model changes",
            "Slide workbooks: full Excel inside\na slide — one file, no attachment",
            "Data tables with formulas, Harvey\nballs & checkboxes in PowerPoint",
            "One file feeds multiple charts\nacross multiple slides",
        ],
    },
    {
        "title_short": "DEAL TIMELINE CHARTS",
        "hdr": H3, "bg": B3,
        "img": str(Path(_A) / "img_020.jpg"),
        "sub": "Gantt · Milestones · Status Columns",
        "bullets": [
            "Build deal timelines directly in\nPowerPoint — no external tool",
            "Scale: day / week / month /\nquarter / fiscal year",
            "Activity bars, milestones & Harvey\nball status in one chart",
            "Rows self-adjust when added or\nremoved — no reformatting",
        ],
    },
    {
        "title_short": "PRODUCTIVITY TOOLS",
        "hdr": H4, "bg": B4,
        "img": str(Path(_C) / "img_008.jpg"),
        "sub": "Layout · Consistency · Workflow",
        "bullets": [
            "Align, size & scale multiple elements\nacross slides in one action",
            "Swap two objects; replace fonts\nacross entire deck in one go",
            "Style files enforce brand colors &\nfonts — automatic compliance",
            "Sanitize: strip metadata before\nsharing; email slides from PPT",
        ],
    },
    {
        "title_short": "CONTENT LIBRARY",
        "hdr": H5, "bg": B5,
        "img": str(Path(_L) / "img_032.jpg"),
        "sub": "Templates · Icons · Chart Scanner",
        "bullets": [
            "250+ slide templates that adapt\nto your brand color palette",
            "10M+ icons & stock photos\nsearchable inside PowerPoint",
            "Slide-level keyword search across\nyour entire saved deck library",
            "Chart Scanner: turn any PDF chart\ninto editable think-cell chart",
        ],
    },
    {
        "title_short": "AI ASSIST  (EARLY ACCESS)",
        "hdr": H6, "bg": B6,
        "img": str(Path(_AI) / "img_007.jpg"),
        "sub": "Data Research · Text Editing · Translation",
        "bullets": [
            "Pull Morningstar/Statista data;\nauto-generates a chart in PPT",
            "Rewrite, shorten or summarize\nslide text via natural language",
            "AI edits run across one slide or\nentire deck in a single command",
            "Translate 100+ languages —\nformatting fully preserved",
        ],
    },
]

# Condensed sidebar content for slide 2 (section title + 2 short bullets)
SIDEBAR_SECTIONS = [
    {
        "title": "Charts & Visualizations", "hdr": H1,
        "lines": [
            "40+ types incl. waterfall/bridge, Mekko, scatter, Gantt & combo",
            "CAGR arrows, diff. arrows & annotations auto-calculated & placed",
        ],
    },
    {
        "title": "Excel Integration", "hdr": H2,
        "lines": [
            "Live Excel links — charts update on model change; slide workbooks",
            "Data tables with formulas, Harvey balls & checkboxes inside PPT",
        ],
    },
    {
        "title": "Deal Timeline Charts", "hdr": H3,
        "lines": [
            "Gantt charts with milestones, Harvey ball status; day-to-FY scale",
            "Self-adjusting layouts — rows add/remove without reformatting",
        ],
    },
    {
        "title": "Productivity Tools", "hdr": H4,
        "lines": [
            "1-click align / scale / swap; replace fonts across full deck",
            "Style files, auto-agenda, sanitize — full brand + workflow control",
        ],
    },
    {
        "title": "Content Library", "hdr": H5,
        "lines": [
            "250+ templates + 10M icons; slide-level search across saved decks",
            "Chart Scanner: replicate any PDF/image chart as editable think-cell",
        ],
    },
    {
        "title": "AI Assist  (Early Access)", "hdr": H6,
        "lines": [
            "Morningstar/Statista data → chart in 1 click, no subscription",
            "Rewrite / translate / summarize via natural language; 100+ languages",
        ],
    },
]

# ── Layout: Slide 1 ──────────────────────────────────────────────────────────
LM, RM, GUTTER, ROW_G = 20, 20, 10, 6
COL_W = (W - LM - RM - 2 * GUTTER) // 3    # 300 pts
COL_X = [LM, LM + COL_W + GUTTER, LM + 2 * (COL_W + GUTTER)]
TITLE_H = 62
ROW_Y   = [TITLE_H + 6, TITLE_H + 6 + 230 + ROW_G]
ROW_H   = 230
HDR_H   = 22
IMG_H   = 80
SHDR_H  = 14
# Remaining vertical space for 2×2 bullet grid
BUL_AREA_H = ROW_H - HDR_H - IMG_H - SHDR_H - 10   # ~104 pt
BUL_ROW_H  = BUL_AREA_H // 2                         # ~52 pt per row
BUL_COL_W  = (COL_W - 6) // 2                        # ~147 pt per col

# ── Layout: Slide 2 sidebar ───────────────────────────────────────────────────
SB_X  = int(W * 0.667)   # 640 — start of right sidebar
SB_W  = W - SB_X         # 320
_SB_START = 48            # y where first block starts (after header)
_SB_FOOT  = 26            # footer bar height
SB_BLOCK_H = (H - _SB_START - _SB_FOOT) // 6   # 77 pt — fits exactly


def run(app, actions):
    return app.Run(
        "PPT_AI_Editor!ExecuteFromString",
        json.dumps({"actions": actions}),
    )


# ═════════════════════════════ SLIDE 1 ACTIONS ═══════════════════════════════
def build_slide1(sl=1):
    a = []
    a.append({"type": "clear_slide", "slide": sl})

    # Top accent bar
    a += [
        {"type": "add_shape", "slide": sl, "ref_name": "s1_accent",
         "kind": "rect", "pos": {"left": 0, "top": 0, "width": W, "height": 6},
         "fill": TC_RED, "no_outline": True},
        {"type": "add_shape", "slide": sl, "ref_name": "s1_stripe",
         "kind": "rect", "pos": {"left": 0, "top": 6, "width": 4, "height": H - 6},
         "fill": NAVY, "no_outline": True},
        {"type": "add_text_box", "slide": sl, "ref_name": "s1_title",
         "pos": {"left": 14, "top": 10, "width": 700, "height": 32},
         "text": TITLE,
         "font_size": 20, "font_bold": True, "font_color": NAVY,
         "h_align": "left", "v_align": "middle"},
        {"type": "add_text_box", "slide": sl, "ref_name": "s1_msg",
         "pos": {"left": 14, "top": 44, "width": 880, "height": 20},
         "text": PAGE_MSG,
         "font_size": 8.5, "font_italic": True, "font_color": GRAY,
         "h_align": "left", "v_align": "middle"},
        {"type": "add_line", "slide": sl,
         "x1": 14, "y1": 66, "x2": W - 14, "y2": 66,
         "color": LGRAY, "weight_pt": 0.5},
    ]

    for idx, sec in enumerate(SECTIONS):
        row = idx // 3
        col = idx % 3
        cx  = COL_X[col]
        cy  = ROW_Y[row]
        hc  = sec["hdr"]
        bc  = sec["bg"]
        r   = f"s1_{idx}"

        # Cell background + header bar
        a += [
            {"type": "add_shape", "slide": sl, "ref_name": f"{r}_bg",
             "kind": "rect",
             "pos": {"left": cx, "top": cy, "width": COL_W, "height": ROW_H},
             "fill": bc, "no_outline": True},
            {"type": "add_shape", "slide": sl, "ref_name": f"{r}_hdr",
             "kind": "rect",
             "pos": {"left": cx, "top": cy, "width": COL_W, "height": HDR_H},
             "fill": hc, "no_outline": True},
            {"type": "add_text_box", "slide": sl, "ref_name": f"{r}_ht",
             "pos": {"left": cx + 6, "top": cy + 2,
                     "width": COL_W - 12, "height": HDR_H - 2},
             "text": sec["title_short"],
             "font_size": 7.5, "font_bold": True, "font_color": WHITE,
             "h_align": "left", "v_align": "middle"},
        ]

        # Image (full width)
        img_y = cy + HDR_H + 2
        img_p = sec.get("img")
        if img_p and Path(img_p).exists():
            a.append({"type": "insert_picture", "slide": sl, "path": img_p,
                       "pos": {"left": cx + 2, "top": img_y,
                                "width": COL_W - 4, "height": IMG_H}})
        else:
            a.append({"type": "add_shape", "slide": sl, "ref_name": f"{r}_ph",
                       "kind": "rect",
                       "pos": {"left": cx + 2, "top": img_y,
                                "width": COL_W - 4, "height": IMG_H},
                       "fill": hc, "no_outline": True})

        # Sub-header (full width, italic tag line)
        shdr_y = img_y + IMG_H + 2
        a.append({"type": "add_text_box", "slide": sl, "ref_name": f"{r}_sh",
                   "pos": {"left": cx + 4, "top": shdr_y,
                            "width": COL_W - 8, "height": SHDR_H},
                   "text": sec["sub"],
                   "font_size": 8, "font_bold": True, "font_italic": True,
                   "font_color": hc, "h_align": "left", "v_align": "middle"})

        # 2×2 bullet grid
        bul_start = shdr_y + SHDR_H + 2
        for j, bullet in enumerate(sec["bullets"]):
            brow = j // 2   # 0 or 1
            bcol = j % 2    # 0 or 1
            bx = cx + 4 + bcol * BUL_COL_W
            by = bul_start + brow * BUL_ROW_H

            # Colored dot
            a.append({"type": "add_shape", "slide": sl, "ref_name": f"{r}_d{j}",
                       "kind": "circle",
                       "pos": {"left": bx + 2, "top": by + 5, "width": 3, "height": 3},
                       "fill": hc, "no_outline": True})
            # Bullet text
            a.append({"type": "add_text_box", "slide": sl, "ref_name": f"{r}_b{j}",
                       "pos": {"left": bx + 8, "top": by,
                                "width": BUL_COL_W - 12, "height": BUL_ROW_H - 2},
                       "text": bullet,
                       "font_size": 7.5, "font_color": BODY,
                       "h_align": "left", "v_align": "top"})

        # Vertical divider between left and right bullet columns
        div_x = cx + 4 + BUL_COL_W - 1
        a.append({"type": "add_line", "slide": sl,
                   "x1": div_x, "y1": bul_start,
                   "x2": div_x, "y2": bul_start + BUL_AREA_H,
                   "color": LGRAY, "weight_pt": 0.4})

    return a


# ═════════════════════════════ SLIDE 2 ACTIONS ═══════════════════════════════
def build_slide2(sl=2):
    a = []
    a.append({"type": "clear_slide", "slide": sl})

    # ── Left 67%: clean blank area ────────────────────────────────────────────
    # Very subtle background + placeholder cue
    a += [
        {"type": "add_shape", "slide": sl, "ref_name": "s2_left_bg",
         "kind": "rect",
         "pos": {"left": 0, "top": 0, "width": SB_X, "height": H},
         "fill": WHITE, "no_outline": True},
        # Subtle left edge accent
        {"type": "add_shape", "slide": sl, "ref_name": "s2_left_accent",
         "kind": "rect",
         "pos": {"left": 0, "top": 0, "width": 4, "height": H},
         "fill": NAVY, "no_outline": True},
        # Placeholder cue text (very light gray)
        {"type": "add_text_box", "slide": sl, "ref_name": "s2_placeholder",
         "pos": {"left": 40, "top": H // 2 - 20, "width": SB_X - 60, "height": 40},
         "text": "[ Your content here ]",
         "font_size": 18, "font_color": "#D1D5DB",
         "h_align": "center", "v_align": "middle"},
    ]

    # ── Vertical divider ─────────────────────────────────────────────────────
    a.append({"type": "add_line", "slide": sl,
               "x1": SB_X, "y1": 0, "x2": SB_X, "y2": H,
               "color": LGRAY, "weight_pt": 1.0})

    # ── Right 33%: think-cell sidebar ────────────────────────────────────────
    # Sidebar background — very dark navy
    a.append({"type": "add_shape", "slide": sl, "ref_name": "s2_sb_bg",
               "kind": "rect",
               "pos": {"left": SB_X, "top": 0, "width": SB_W, "height": H},
               "fill": "#0F1E33", "no_outline": True})

    # Sidebar header — think-cell name + tagline
    a += [
        {"type": "add_text_box", "slide": sl, "ref_name": "s2_sb_title",
         "pos": {"left": SB_X + 10, "top": 6, "width": SB_W - 16, "height": 22},
         "text": "think-cell",
         "font_size": 13, "font_bold": True, "font_color": WHITE,
         "h_align": "left", "v_align": "middle"},
        {"type": "add_shape", "slide": sl, "ref_name": "s2_sb_dot",
         "kind": "circle",
         "pos": {"left": SB_X + 10 + 73, "top": 12, "width": 5, "height": 5},
         "fill": TC_RED, "no_outline": True},
        {"type": "add_text_box", "slide": sl, "ref_name": "s2_sb_sub",
         "pos": {"left": SB_X + 10, "top": 28, "width": SB_W - 16, "height": 14},
         "text": "PowerPoint add-in  ·  used by all top 10 consulting firms",
         "font_size": 7, "font_color": "#8899AA",
         "h_align": "left", "v_align": "middle"},
        # Orange accent line under header
        {"type": "add_line", "slide": sl,
         "x1": SB_X + 10, "y1": 44, "x2": SB_X + SB_W - 10, "y2": 44,
         "color": TC_RED, "weight_pt": 1.0},
    ]

    # ── 6 section blocks ──────────────────────────────────────────────────────
    for i, sec in enumerate(SIDEBAR_SECTIONS):
        sy = 48 + i * SB_BLOCK_H
        hc = sec["hdr"]

        # Left color bar
        a.append({"type": "add_shape", "slide": sl, "ref_name": f"s2_bar{i}",
                   "kind": "rect",
                   "pos": {"left": SB_X, "top": sy + 2,
                            "width": 3, "height": SB_BLOCK_H - 4},
                   "fill": hc, "no_outline": True})

        # Section title
        a.append({"type": "add_text_box", "slide": sl, "ref_name": f"s2_t{i}",
                   "pos": {"left": SB_X + 8, "top": sy + 3,
                            "width": SB_W - 14, "height": 16},
                   "text": sec["title"].upper(),
                   "font_size": 7.5, "font_bold": True, "font_color": WHITE,
                   "h_align": "left", "v_align": "middle"})

        # 2 content lines  (SB_BLOCK_H=77: title at +2, lines at +18 and +46)
        for k, line in enumerate(sec["lines"]):
            ly = sy + 18 + k * 28
            a += [
                {"type": "add_shape", "slide": sl, "ref_name": f"s2_d{i}_{k}",
                 "kind": "circle",
                 "pos": {"left": SB_X + 10, "top": ly + 4, "width": 3, "height": 3},
                 "fill": hc, "no_outline": True},
                {"type": "add_text_box", "slide": sl, "ref_name": f"s2_l{i}_{k}",
                 "pos": {"left": SB_X + 17, "top": ly,
                          "width": SB_W - 23, "height": 26},
                 "text": line,
                 "font_size": 7, "font_color": "#AABBCC",
                 "h_align": "left", "v_align": "top"},
            ]

        # Separator line between blocks (not after last)
        if i < len(SIDEBAR_SECTIONS) - 1:
            sep_y = sy + SB_BLOCK_H - 1
            a.append({"type": "add_line", "slide": sl,
                       "x1": SB_X + 6, "y1": sep_y,
                       "x2": SB_X + SB_W - 6, "y2": sep_y,
                       "color": "#1E3050", "weight_pt": 0.4})

    # Stats footer at bottom of sidebar
    stats = [("70%", "faster"), ("40+", "chart types"), ("35k+", "companies")]
    sw = SB_W // 3
    footer_y = _SB_START + 6 * SB_BLOCK_H + 2
    a.append({"type": "add_shape", "slide": sl, "ref_name": "s2_footer_bg",
               "kind": "rect",
               "pos": {"left": SB_X, "top": footer_y - 4, "width": SB_W, "height": 26},
               "fill": TC_RED, "no_outline": True})
    for k, (num, lbl) in enumerate(stats):
        fx = SB_X + k * sw
        a += [
            {"type": "add_text_box", "slide": sl, "ref_name": f"s2_sn{k}",
             "pos": {"left": fx + 4, "top": footer_y - 2, "width": 28, "height": 22},
             "text": num,
             "font_size": 9, "font_bold": True, "font_color": WHITE,
             "h_align": "right", "v_align": "middle"},
            {"type": "add_text_box", "slide": sl, "ref_name": f"s2_sl{k}",
             "pos": {"left": fx + 34, "top": footer_y - 2, "width": sw - 36, "height": 22},
             "text": lbl,
             "font_size": 7, "font_color": "#FFD0C0",
             "h_align": "left", "v_align": "middle"},
        ]

    return a


def main() -> int:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    tmpdir = Path(tempfile.mkdtemp(prefix="tc_overview_"))
    try:
        try:
            app.Visible = True
        except Exception:
            pass

        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Add()
        # Add a second blank slide
        layout = deck.SlideMaster.CustomLayouts(1)
        if deck.Slides.Count == 0:
            deck.Slides.AddSlide(1, layout)
        deck.Slides.AddSlide(2, layout)

        tmp_path = tmpdir / "thinkcell_overview.pptx"
        deck.SaveAs(str(tmp_path))
        deck.Windows(1).Activate()
        time.sleep(0.3)

        # ── Build slide 1 ──────────────────────────────────────────────────
        print("[build] slide 1 — 6-section grid...")
        r1 = run(app, build_slide1(sl=1))
        print(f"        {r1}")

        # ── Build slide 2 ──────────────────────────────────────────────────
        print("[build] slide 2 — sidebar layout...")
        r2 = run(app, build_slide2(sl=2))
        print(f"        {r2}")

        # ── Save + export ─────────────────────────────────────────────────
        OUT_DECK.parent.mkdir(parents=True, exist_ok=True)
        deck.SaveAs(str(OUT_DECK))
        print(f"[saved] {OUT_DECK}")

        PNG_DIR.mkdir(parents=True, exist_ok=True)
        for sl in deck.Slides:
            sl.Export(
                str(PNG_DIR / f"slide_{sl.SlideNumber:02d}.png"),
                "PNG", 1280, 720,
            )
        print(f"[png]   {PNG_DIR}")

    finally:
        try:
            for p in list(app.Presentations):
                try:
                    p.Saved = True
                    p.Close()
                except Exception:
                    pass
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(0.5)
    return 0


if __name__ == "__main__":
    sys.exit(main() or 0)

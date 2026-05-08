"""Regenerate test_decks/ from scratch.

Run: python tests/make_test_decks.py
"""
import sys
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

REPO_ROOT = Path(__file__).resolve().parent.parent
DECKS_DIR = REPO_ROOT / "test_decks"


def make_smoke_3slide(path: Path) -> None:
    pres = Presentation()
    layout_title = pres.slide_layouts[0]
    layout_content = pres.slide_layouts[1]

    s1 = pres.slides.add_slide(layout_title)
    s1.shapes.title.text = "Q3 Results"
    s1.placeholders[1].text = "Subtitle text"

    s2 = pres.slides.add_slide(layout_content)
    s2.shapes.title.text = "Bullets"
    s2.placeholders[1].text = "Revenue up 12%\nMargins improved\nHeadcount stable"

    s3 = pres.slides.add_slide(layout_content)
    s3.shapes.title.text = "Next Steps"
    s3.placeholders[1].text = "Plan Q4\nHire 5 engineers"

    pres.save(str(path))


def make_full_visual(path: Path) -> None:
    pres = Presentation()
    layout_blank = pres.slide_layouts[6]

    slide = pres.slides.add_slide(layout_blank)

    # Title textbox
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(1))
    tf = tb.text_frame
    tf.text = "Q3 Visual Slide"
    run = tf.paragraphs[0].runs[0]
    run.font.size = Pt(32)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)

    # Table
    rows, cols = 3, 3
    table_shape = slide.shapes.add_table(
        rows, cols, Inches(0.5), Inches(1.5), Inches(9), Inches(2)
    )
    tbl = table_shape.table
    headers = ["Metric", "Q2", "Q3"]
    for c, h in enumerate(headers):
        tbl.cell(0, c).text = h
    data = [
        ["Revenue", "100", "112"],
        ["Margin", "26%", "28%"],
    ]
    for r, row in enumerate(data, start=1):
        for c, v in enumerate(row):
            tbl.cell(r, c).text = v

    # Plain rectangle with fill
    from pptx.enum.shapes import MSO_SHAPE
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4), Inches(2), Inches(1)
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    rect.text_frame.text = "Box"

    pres.save(str(path))


def make_phase2(path: Path) -> None:
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN

    pres = Presentation()
    layout_blank = pres.slide_layouts[6]

    # --- Slide 1: multi-paragraph bullet body + speaker notes
    s1 = pres.slides.add_slide(layout_blank)

    title = s1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title.text_frame.text = "Bullet Body Slide"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
    title.text_frame.paragraphs[0].runs[0].font.bold = True

    body = s1.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(4))
    tf = body.text_frame
    tf.text = "First point about revenue"
    tf.paragraphs[0].runs[0].font.size = Pt(20)

    p2 = tf.add_paragraph()
    p2.text = "Second point about margins"
    p2.runs[0].font.size = Pt(20)

    p3 = tf.add_paragraph()
    p3.text = "Third point about headcount"
    p3.runs[0].font.size = Pt(20)

    s1.notes_slide.notes_text_frame.text = "Highlight Q3 outperformance and Y/Y growth"

    # --- Slide 2: native chart
    s2 = pres.slides.add_slide(layout_blank)
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("FY24", (100, 110, 120, 130))
    s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1), Inches(1.5), Inches(8), Inches(4.5),
        chart_data,
    )

    # --- Slide 3: three ungrouped boxes (group_shapes test will group them)
    s3 = pres.slides.add_slide(layout_blank)
    box1 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1.5), Inches(2), Inches(1))
    box1.text_frame.text = "A"
    box2 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4), Inches(1.5), Inches(2), Inches(1))
    box2.text_frame.text = "B"
    box3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7), Inches(1.5), Inches(2), Inches(1))
    box3.text_frame.text = "C"

    # --- Slide 4: 4x3 table + plain rectangle
    s4 = pres.slides.add_slide(layout_blank)
    rows, cols = 4, 3
    tbl_shape = s4.shapes.add_table(rows, cols, Inches(0.5), Inches(0.5), Inches(9), Inches(3))
    tbl = tbl_shape.table
    headers = ["Metric", "FY24", "FY25"]
    for c, h in enumerate(headers):
        tbl.cell(0, c).text = h
    body_data = [
        ["Revenue", "100", "112"],
        ["Margin", "26%", "28%"],
        ["Headcount", "1,200", "1,250"],
    ]
    for r, row in enumerate(body_data, start=1):
        for c, v in enumerate(row):
            tbl.cell(r, c).text = v

    rect = s4.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4.5), Inches(2), Inches(0.8)
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    rect.text_frame.text = "Plain"

    pres.save(str(path))


def main() -> int:
    DECKS_DIR.mkdir(parents=True, exist_ok=True)
    make_smoke_3slide(DECKS_DIR / "smoke_3slide.pptx")
    make_full_visual(DECKS_DIR / "full_visual.pptx")
    make_phase2(DECKS_DIR / "phase2.pptx")
    print(f"[done] wrote 3 decks to {DECKS_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

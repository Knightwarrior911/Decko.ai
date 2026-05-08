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


def main() -> int:
    DECKS_DIR.mkdir(parents=True, exist_ok=True)
    make_smoke_3slide(DECKS_DIR / "smoke_3slide.pptx")
    make_full_visual(DECKS_DIR / "full_visual.pptx")
    print(f"[done] wrote 2 decks to {DECKS_DIR}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

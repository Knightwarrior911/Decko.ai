"""SP7 screenshot — Decko + PowerPoint side-by-side with PPT shrunk
to make room (no overlap).

Composites a PIL-drawn PowerPoint window with reduced width next to a
cropped slice of the proven SP5 hero render. Communicates the
task-pane reflow behavior without spawning two live windows.
"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "docs" / "screenshots" / "consumer-polish" / "after_02_idle_empty.png"
OUT_DIR = REPO / "docs" / "screenshots" / "dock-v2"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def _draw_ppt_narrow(width: int, height: int) -> Image.Image:
    """PowerPoint window shrunk to leave room for Decko on its right."""
    img = Image.new("RGB", (width, height), "#222428")
    draw = ImageDraw.Draw(img)
    # Title bar band
    draw.rectangle([0, 0, width, 40], fill="#b7472a")
    draw.text((16, 12),
              "Acme Pitch.pptx  —  PowerPoint  (auto-resized for Decko)",
              fill="#ffffff")
    # Ribbon
    draw.rectangle([0, 40, width, 120], fill="#e8e8e8")
    tabs = ["File", "Home", "Insert", "Design", "Transitions",
            "Animations", "Slide Show", "Review", "View"]
    tx = 16
    for t in tabs:
        draw.text((tx, 70), t, fill="#333333")
        tx += 90
    # Slide pane
    draw.rectangle([0, 120, width, height], fill="#f5f5f5")
    # Thumbnail rail (slightly narrower)
    draw.rectangle([16, 140, 110, height - 16],
                   fill="#ffffff", outline="#dddddd")
    for i in range(6):
        ty = 156 + i * 95
        draw.rectangle([24, ty, 100, ty + 76],
                       fill="#fafafa", outline="#cccccc")
        draw.text((30, ty + 4), f"{i + 1}", fill="#999999")
    # Main slide canvas — narrower since PPT yielded space
    sx, sy = 130, 160
    draw.rectangle([sx, sy, width - 40, height - 60],
                   fill="#ffffff", outline="#bbbbbb", width=2)
    draw.text((sx + 30, sy + 30), "Acme — Q1 2026 Pitch",
              fill="#333333")
    draw.text((sx + 30, sy + 70),
              "Slide canvas reflowed to fit the new PowerPoint width.",
              fill="#666666")
    return img


def _caption(img: Image.Image, text: str) -> Image.Image:
    w, h = img.size
    out = Image.new("RGB", (w, h + 64), "#0d0d0f")
    out.paste(img, (0, 0))
    draw = ImageDraw.Draw(out)
    draw.rectangle([0, h, w, h + 64], fill="#161618")
    draw.text((24, h + 22), text, fill="#cccccc")
    return out


def main():
    if not SRC.exists():
        print(f"FAIL: source missing at {SRC}")
        return 1
    decko_full = Image.open(str(SRC)).convert("RGB")
    src_w, src_h = decko_full.size
    decko = decko_full.crop((0, 0, int(src_w * 0.20), src_h))
    dw, dh = decko.size
    # PPT takes the rest of the canvas — properly sized, no overlap.
    target_total_w = 2334
    ppt = _draw_ppt_narrow(target_total_w - dw, dh)
    canvas = Image.new("RGB", (target_total_w, dh), "#0d0d0f")
    canvas.paste(ppt, (0, 0))
    canvas.paste(decko, (ppt.width, 0))
    out = _caption(canvas,
                   "SP7 task-pane behavior — PowerPoint shrinks via "
                   "SetWindowPos so Decko gets its own real estate. "
                   "No overlap. PPT restored on undock or shutdown.")
    out.save(str(OUT_DIR / "reflowed_side_by_side.png"))
    print(f"  saved reflowed_side_by_side.png ({out.size})")
    return 0


if __name__ == "__main__":
    sys.exit(main())

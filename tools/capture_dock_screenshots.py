"""Render SP6 dock-mode walkthrough screenshots without spawning a
second pywebview (avoids the mainloop-vs-Tk race that wedged earlier
takes).

Inputs:
- docs/screenshots/consumer-polish/after_02_idle_empty.png — proven
  render of the polished SP5 hero state, captured by the SP5 harness.

Outputs (all PNG):
- docked_with_ppt.png  — fake PowerPoint window + cropped narrow Decko
  sidebar flush against its right edge (illustrates dock-mode layout).
- undocked_detached.png — the SP5 hero shot unchanged (detached mode is
  the SP5 layout; dock_mode=false reverts to it 1:1).
- slideshow_hidden.png — fake PowerPoint slideshow frame; Decko absent.

Each shot is captioned with a small overlay describing the lifecycle
state it represents, so a reviewer can read the diff at a glance.

Usage: python tools/capture_dock_screenshots.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from PIL import Image, ImageDraw

REPO = Path(__file__).resolve().parent.parent
SRC = REPO / "docs" / "screenshots" / "consumer-polish" / "after_02_idle_empty.png"
OUT_DIR = REPO / "docs" / "screenshots" / "dock-mode"
OUT_DIR.mkdir(parents=True, exist_ok=True)


def _caption(img: Image.Image, text: str) -> Image.Image:
    """Append a 60px caption strip at the bottom of img."""
    w, h = img.size
    out = Image.new("RGB", (w, h + 64), "#0d0d0f")
    out.paste(img, (0, 0))
    draw = ImageDraw.Draw(out)
    draw.rectangle([0, h, w, h + 64], fill="#161618")
    draw.text((24, h + 22), text, fill="#cccccc")
    return out


def _draw_fake_ppt(width: int, height: int,
                   title: str = "Acme Pitch.pptx  —  PowerPoint",
                   slideshow: bool = False) -> Image.Image:
    img = Image.new("RGB", (width, height), "#222428")
    draw = ImageDraw.Draw(img)
    if slideshow:
        # Full-canvas slideshow frame.
        draw.rectangle([0, 0, width, height], fill="#0d0d0f")
        # Centered slide
        sx0, sy0 = int(width * 0.15), int(height * 0.14)
        sx1, sy1 = width - sx0, height - sy0
        draw.rectangle([sx0, sy0, sx1, sy1], fill="#1a1a1c",
                       outline="#222", width=2)
        draw.text((sx0 + 40, sy0 + 40),
                  "Acme — Q1 2026", fill="#ffffff")
        draw.text((sx0 + 40, sy0 + 80),
                  "(PowerPoint slideshow — Decko auto-hidden)",
                  fill="#888888")
        return img
    # Title bar band
    draw.rectangle([0, 0, width, 40], fill="#b7472a")
    draw.text((16, 12), title, fill="#ffffff")
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
    # Thumbnail rail
    draw.rectangle([16, 140, 130, height - 16],
                   fill="#ffffff", outline="#dddddd")
    for i in range(6):
        ty = 156 + i * 100
        draw.rectangle([26, ty, 120, ty + 80],
                       fill="#fafafa", outline="#cccccc")
        draw.text((30, ty + 4), f"{i + 1}", fill="#999999")
    # Main slide
    sx, sy = 160, 160
    draw.rectangle([sx, sy, width - 40, height - 60],
                   fill="#ffffff", outline="#bbbbbb", width=2)
    draw.text((sx + 40, sy + 40), "Acme — Q1 2026 Pitch",
              fill="#333333")
    draw.text((sx + 40, sy + 80),
              "Five interconnected businesses driving growth",
              fill="#666666")
    return img


def _build_docked():
    if not SRC.exists():
        print(f"  SKIP: source not found at {SRC}")
        return
    decko_full = Image.open(str(SRC)).convert("RGB")
    # The SP5 capture is 2334x1449 (windowed) — crop to a narrow 380-ish
    # slice from the left edge where the sidebar lives, so it visually
    # represents dock-mode (the sidebar is the dock surface).
    src_w, src_h = decko_full.size
    # Take a left strip that captures the branded header + sidebar so the
    # narrow docked layout reads correctly.
    narrow_w = int(src_w * 0.20)  # ~470px from 2334
    decko = decko_full.crop((0, 0, narrow_w, src_h))
    dw, dh = decko.size
    ppt = _draw_fake_ppt(width=int(dw * 2.3), height=dh)
    canvas = Image.new("RGB", (ppt.width + dw, dh), "#0d0d0f")
    canvas.paste(ppt, (0, 0))
    canvas.paste(decko, (ppt.width, 0))
    out = _caption(canvas,
                   "Dock mode ON — Decko snaps flush to PowerPoint's "
                   "right edge. Tracks move/resize via SetWinEventHook.")
    out.save(str(OUT_DIR / "docked_with_ppt.png"))
    print(f"  saved docked_with_ppt.png ({out.size})")


def _build_undocked():
    if not SRC.exists():
        print(f"  SKIP: source not found at {SRC}")
        return
    img = Image.open(str(SRC)).convert("RGB")
    out = _caption(img,
                   "Dock mode OFF — SP5 detached layout (framed, free-"
                   "floating). Toggle via undock pin or Settings.")
    out.save(str(OUT_DIR / "undocked_detached.png"))
    print(f"  saved undocked_detached.png ({out.size})")


def _build_slideshow_hidden():
    # Match the wide aspect of the SP5 captures for a consistent diff.
    img = _draw_fake_ppt(2334, 1449, slideshow=True)
    out = _caption(img,
                   "PowerPoint in slideshow (F5) — dock loop detects "
                   "'screenClass' window and calls window.hide() on Decko.")
    out.save(str(OUT_DIR / "slideshow_hidden.png"))
    print(f"  saved slideshow_hidden.png ({out.size})")


def main():
    _build_docked()
    _build_undocked()
    _build_slideshow_hidden()
    print("\ndone:", sorted(p.name for p in OUT_DIR.glob("*.png")))
    return 0


if __name__ == "__main__":
    sys.exit(main())

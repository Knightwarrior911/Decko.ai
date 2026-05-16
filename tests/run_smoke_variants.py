"""generate_variants principled-archetype smoke test, via PowerPoint COM.

Run: python tests/run_smoke_variants.py   (exit 0 only at 100%)

generate_variants template+n=5 -> 5 slides, each a STRUCTURALLY
different principled archetype, each obeying the design rules:
  - every shape inside a >=36pt margin
  - exactly ONE strictly-dominant (largest-font) element
  - <=2 distinct fill colours
  - all carry the content text
Plus: the exported prompt contains the DECK DESIGN PRINCIPLES block.

COM-resilient. No AI/network.
"""
import json
import os
import sys
import tempfile
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

CARRIER = Path(__file__).resolve().parent.parent / "PPT_AI_Editor.pptm"
SW, SH, M, EPS = 960.0, 540.0, 36.0, 1.5
N = 5
CONTENT = {"heading": "VARH", "body": "VB1 VB2 VB3 VB4"}
SENTINELS = ("VARH", "VB1", "VB4")


def open_app():
    last = None
    for _ in range(15):
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = True
            return app
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2.0)
    raise RuntimeError(f"PowerPoint COM bring-up failed: {last!r}")


def run_session():
    pythoncom.CoInitialize()
    app = open_app()
    carrier = pres = None
    fails = []
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
        pres = app.Presentations.Add()
        pres.PageSetup.SlideWidth = SW
        pres.PageSetup.SlideHeight = SH
        pres.Windows(1).Activate()
        time.sleep(1.0)

        summ = app.Run("PPT_AI_Editor!ExecuteFromString", json.dumps(
            {"actions": [{"type": "generate_variants", "template": "title",
                          "content": CONTENT, "n": N}],
             "verify_after": False}))
        if "1 applied" not in summ:
            fails.append(("apply", "summary", summ))

        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        vs = snap["slides"][-N:]
        if len(vs) != N:
            fails.append(("count", "slides", f"{len(vs)} new, want {N}"))

        sigs = []
        for i, sl in enumerate(vs):
            shapes = sl["shapes"]
            txt = " || ".join((s.get("text") or "") for s in shapes)
            for sen in SENTINELS:
                if sen not in txt:
                    fails.append((f"v{i}", "content", f"{sen!r} missing"))

            # margin conformance
            for s in shapes:
                p = s.get("pos") or {}
                L, T = p.get("left"), p.get("top")
                W, Hh = p.get("width"), p.get("height")
                if None in (L, T, W, Hh):
                    continue
                if (L < M - EPS or T < M - EPS or L + W > SW - M + EPS
                        or T + Hh > SH - M + EPS):
                    fails.append((f"v{i}", "margin",
                                  f"shape {L:.0f},{T:.0f},{W:.0f},{Hh:.0f} "
                                  f"breaks >= {M} margin"))
                    break

            # exactly one strictly-dominant font
            fsizes = [s["font"]["size"] for s in shapes
                      if isinstance(s.get("font"), dict)
                      and isinstance(s["font"].get("size"), (int, float))]
            if fsizes:
                mx = max(fsizes)
                if fsizes.count(mx) != 1:
                    fails.append((f"v{i}", "dominant",
                                  f"{fsizes.count(mx)} shapes share max font {mx}"))
            else:
                fails.append((f"v{i}", "dominant", "no font sizes found"))

            # <=2 distinct fill colours
            fills = {(s.get("fill") or "").upper() for s in shapes
                     if (s.get("fill") or "").strip()}
            if len(fills) > 2:
                fails.append((f"v{i}", "fills",
                              f"{len(fills)} distinct fills: {sorted(fills)}"))

            sigs.append((len(shapes), tuple(sorted(round(x) for x in fsizes))))

        if len(set(sigs)) != len(sigs):
            fails.append(("distinct", "archetypes",
                          f"signatures not all unique: {sigs}"))

        # design principles present in the exported prompt
        ppath = os.path.join(tempfile.mkdtemp(prefix="vp_"), "prompt.txt")
        app.Run("PPT_AI_Editor!DumpPromptToFile", ppath)
        prm = Path(ppath).read_text(encoding="utf-8", errors="replace")
        for marker in ("DECK DESIGN PRINCIPLES", "HIERARCHY", "CLUSTERING",
                       "WHITESPACE", "accent"):
            if marker not in prm:
                fails.append(("prompt", "principles", f"{marker!r} missing"))
        return fails
    finally:
        for p in (pres, carrier):
            try:
                if p is not None:
                    p.Saved = True
                    p.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
        time.sleep(2.0)


def main() -> int:
    transient = (pywintypes.com_error, AttributeError, RuntimeError)
    result = None
    last = None
    for attempt in range(1, 4):
        try:
            result = run_session()
            break
        except transient as e:  # noqa: PERF203
            last = e
            print(f"  retry transient COM error (attempt {attempt}): {e!r}")
            os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
            time.sleep(5.0)
    if result is None:
        print(f"FAIL: variants run failed after retries: {last!r}")
        return 1
    fails = result
    print("checks: 5 archetypes - distinct + margin + one-dominant + <=2 fills "
          "+ content + design principles in prompt")
    if fails:
        for stage, chk, why in fails:
            print(f"  FAIL [{stage}] {chk}: {why}")
    else:
        print("  5 structurally-distinct conformant archetypes; "
              "design principles present in prompt")
    ok = not fails
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

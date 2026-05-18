"""Citi-repro harness: drive Decko carrier via COM, run an actions JSON file,
export PNGs, assert chart-reality.

Usage:
  python tests/citi_repro/run_citi.py <actions.json> [--deck NAME]

Reads {"actions":[...]}, opens PPT_AI_Editor.pptm so its VBA loads, creates a
fresh deck, sends actions to ExecuteFromString in chunks (direct COM string arg
bypasses the MSForms textbox corruption), prints the action-log skip/err
summary, exports each slide to PNG, and prints a chart-reality report
(real ChartObject vs autoshape) per slide.
"""
import json, sys, time, os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
CARRIER = ROOT / "PPT_AI_Editor.pptm"
WORK = ROOT / "tests" / "citi_repro"
MSO_CHART = 3
CHUNK = 30  # actions per ExecuteFromString call


def log_lines(log: Path):
    if not log.exists():
        return []
    return [l for l in log.read_text(encoding="utf-8", errors="replace").splitlines() if l.strip()]


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: run_citi.py <actions.json> [--deck NAME]")
        return 2
    actions_path = Path(sys.argv[1])
    deck_name = "citi_repro"
    if "--deck" in sys.argv:
        deck_name = sys.argv[sys.argv.index("--deck") + 1]
    deck_path = WORK / f"{deck_name}.pptx"
    log = Path(str(deck_path) + ".action_log.jsonl")
    png_dir = WORK / f"{deck_name}_pngs"
    png_dir.mkdir(parents=True, exist_ok=True)

    actions = json.loads(actions_path.read_text(encoding="utf-8"))["actions"]
    print(f"[in] {actions_path}  ({len(actions)} actions)")

    import win32com.client as win32
    if log.exists():
        log.unlink()

    # kill stray PowerPoint from a prior iteration (no concurrent COM harnesses)
    os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
    time.sleep(1.5)

    app = win32.DispatchEx("PowerPoint.Application")
    app.Visible = True
    try:
        app.AutomationSecurity = 1
    except Exception as e:
        print(f"warn: AutomationSecurity: {e}")

    print(f"[open] {CARRIER}")
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)

    if deck_path.exists():
        try: deck_path.unlink()
        except Exception: pass
    deck = app.Presentations.Add()
    deck.SaveAs(str(deck_path))
    print(f"[deck] {deck_path}  (starts with {deck.Slides.Count} slide)")

    def run(payload: str) -> str:
        try:
            deck.Windows(1).Activate()
        except Exception:
            pass
        try:
            return app.Run("ExecuteFromString", payload)
        except Exception:
            return app.Run("PPT_AI_Editor.pptm!ExecuteFromString", payload)

    # send in chunks; engine is stateless + slide-targeted so order-preserving chunking is safe
    t0 = time.time()
    for i in range(0, len(actions), CHUNK):
        chunk = actions[i:i + CHUNK]
        before = len(log_lines(log))
        run(json.dumps({"actions": chunk}))
        after = len(log_lines(log))
        print(f"  chunk {i//CHUNK+1}: actions {i}..{i+len(chunk)-1}  log+{after-before}")
        time.sleep(0.05)
    deck.Save()
    print(f"[exec] done in {time.time()-t0:.1f}s")

    # action-log skip/error report
    skips = []
    for l in log_lines(log):
        try:
            e = json.loads(l)
        except Exception:
            continue
        st = e.get("status")
        if st and st != "ok":
            skips.append((e.get("op"), st, e.get("reason", "")))
    print(f"\n=== ACTION LOG: {len(log_lines(log))} entries, {len(skips)} non-ok ===")
    for op, st, rs in skips[:60]:
        print(f"  {st:9s} {op:24s} {rs}")

    # chart-reality assertion
    print("\n=== CHART REALITY ===")
    for i in range(1, deck.Slides.Count + 1):
        sl = deck.Slides(i)
        charts = autoshapes = tables = 0
        for sh in sl.Shapes:
            try:
                if sh.HasChart:
                    charts += 1
                    continue
            except Exception:
                pass
            try:
                if sh.HasTable:
                    tables += 1
                    continue
            except Exception:
                pass
            try:
                if sh.Type == 1:  # msoAutoShape
                    autoshapes += 1
            except Exception:
                pass
        ctypes = []
        for sh in sl.Shapes:
            try:
                if sh.HasChart:
                    ctypes.append(int(sh.Chart.ChartType))
            except Exception:
                pass
        print(f"  slide {i}: charts={charts} tables={tables} autoshapes={autoshapes} chart_types={ctypes}")

    # export PNGs
    print("\n=== PNG EXPORT ===")
    for sl in deck.Slides:
        out = png_dir / f"slide_{sl.SlideNumber:02d}.png"
        try:
            sl.Export(str(out), "PNG", 1920, 1080)
            print(f"  {out}")
        except Exception as e:
            print(f"  EXPORT FAIL slide {sl.SlideNumber}: {e}")

    deck.Save()
    print(f"\ndeck: {deck_path}\nlog:  {log}\npng:  {png_dir}")
    print("(PowerPoint left open)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

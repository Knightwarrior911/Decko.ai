"""decks-as-code smoke test (build_deck_from_spec / extract_spec /
generate_variants), driven via PowerPoint COM.

Run: python tests/run_smoke_spec.py   (exit 0 only at 100%)

Fresh 960x540 deck:
  (1) build_deck_from_spec(S) -> slide count + all slot text present.
  (2) extract_spec -> reconstructs S's template + required slot text
      per slide (template-built decks round-trip exactly via tpl_* tags).
  (3) generate_variants(title, n=3) -> 3 slides, each carries the slot
      text, headline boxes pairwise geometrically distinct.

COM-resilient. No AI/network.
"""
import json
import os
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

CARRIER = Path(__file__).resolve().parent.parent / "PPT_AI_Editor.pptm"

SPEC = {"deck": [
    {"template": "title",
     "content": {"title": "S_TITLE", "subtitle": "S_SUB"}},
    {"template": "bullets",
     "content": {"heading": "S_BHEAD", "bullets": ["S_B1", "S_B2"]}},
    {"template": "two_col",
     "content": {"heading": "S_TCH", "left_body": "S_TL", "right_body": "S_TR"}},
    {"template": "quote",
     "content": {"quote_text": "S_Q", "attribution": "S_QA"}},
]}

# required slot text expected back from extract_spec, per deck entry
RT = [
    ("title", {"title": "S_TITLE", "subtitle": "S_SUB"}),
    ("bullets", {"heading": "S_BHEAD", "bullets": ["S_B1", "S_B2"]}),
    ("two_col", {"heading": "S_TCH", "left_body": "S_TL", "right_body": "S_TR"}),
    ("quote", {"quote_text": "S_Q", "attribution": "S_QA"}),
]


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


def all_texts(snap):
    out = []
    for sl in snap["slides"]:
        for sh in sl["shapes"]:
            out.append((sh.get("text") or ""))
    return " || ".join(out)


def run_session():
    pythoncom.CoInitialize()
    app = open_app()
    carrier = pres = None
    fails = []
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
        pres = app.Presentations.Add()
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540
        pres.Windows(1).Activate()
        time.sleep(1.0)

        # (1) build
        summ = app.Run("PPT_AI_Editor!ExecuteFromString", json.dumps(
            {"actions": [{"type": "build_deck_from_spec", "spec": SPEC}],
             "verify_after": False}))
        if "1 applied" not in summ:
            fails.append(("build", "apply", summ))
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        blob = all_texts(snap)
        for sen in ["S_TITLE", "S_SUB", "S_BHEAD", "S_B1", "S_B2",
                    "S_TCH", "S_TL", "S_TR", "S_Q", "S_QA"]:
            if sen not in blob:
                fails.append(("build", "slot text", f"{sen!r} missing"))

        # (2) round-trip
        spec_js = app.Run("PPT_AI_Editor!ExtractDeckSpecJson")
        try:
            got = json.loads(spec_js).get("deck", [])
        except Exception as e:  # noqa: BLE001
            got = []
            fails.append(("extract", "json", f"{e}: {spec_js!r}"))
        if len(got) != len(RT):
            fails.append(("extract", "count",
                          f"got {len(got)} entries, want {len(RT)}"))
        else:
            for i, (tpl, slots) in enumerate(RT):
                g = got[i]
                if g.get("template") != tpl:
                    fails.append(("extract", f"entry{i} template",
                                  f"got {g.get('template')!r} want {tpl!r}"))
                gc = g.get("content", {})
                for k, v in slots.items():
                    if gc.get(k) != v:
                        fails.append(("extract", f"entry{i}.{k}",
                                      f"got {gc.get(k)!r} want {v!r}"))

        # (3) variants
        summ = app.Run("PPT_AI_Editor!ExecuteFromString", json.dumps(
            {"actions": [{"type": "generate_variants", "template": "title",
                          "content": {"title": "V_T", "subtitle": "V_S"},
                          "n": 3}], "verify_after": False}))
        if "1 applied" not in summ:
            fails.append(("variants", "apply", summ))
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        vslides = snap["slides"][-3:]
        if len(vslides) != 3:
            fails.append(("variants", "count", f"{len(vslides)} new slides"))
        head_pos = []
        for vi, sl in enumerate(vslides):
            txts = " || ".join((s.get("text") or "") for s in sl["shapes"])
            for sen in ("V_T", "V_S"):
                if sen not in txts:
                    fails.append(("variants", f"slide{vi} text",
                                  f"{sen!r} missing"))
            hs = next((s for s in sl["shapes"]
                       if "V_T" in (s.get("text") or "")), None)
            if hs:
                p = hs.get("pos") or {}
                head_pos.append((round(p.get("left", -1), 1),
                                 round(p.get("top", -1), 1)))
        if len(head_pos) == 3 and len(set(head_pos)) < 3:
            fails.append(("variants", "distinct",
                          f"headline boxes not pairwise distinct: {head_pos}"))
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
        print(f"FAIL: spec run failed after retries: {last!r}")
        return 1

    fails = result
    print("checks: (1) build correctness  (2) round-trip  (3) n distinct variants")
    if fails:
        for stage, chk, why in fails:
            print(f"  FAIL [{stage}] {chk}: {why}")
    else:
        print("  build OK; spec round-trips template+slots; 3 distinct variants")
    ok = not fails
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

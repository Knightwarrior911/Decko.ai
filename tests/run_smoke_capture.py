"""Captured-templates ("Deck DNA") smoke test, driven via PowerPoint COM.

Run: python tests/run_smoke_capture.py   (exit 0 only at 100%)

Proves the three make-or-break properties on a temp registry file:
  (a) data registry  — capture writes the JSON sidecar with auto-slots
  (b) generic render  — apply_template <captured> reproduces captured
                        geometry + substitutes new slot text
  + generate_variants templates:[captured, builtin] -> distinct slides
  + (c) live manifest — BuildCapturedManifest contains the captured
                        name/slots/usage
  + safe delete       — gone from registry & manifest, a previously
                        built slide is UNCHANGED

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
TOL = 2.5


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


def texts_of(slide):
    return [(s.get("text") or "") for s in slide["shapes"]]


def find_pos(slide, needle):
    for s in slide["shapes"]:
        if needle in (s.get("text") or ""):
            return s.get("pos") or {}
    return None


def run_session():
    pythoncom.CoInitialize()
    tmp = tempfile.mkdtemp(prefix="decko_cap_")
    reg = os.path.join(tmp, "templates.json")
    app = open_app()
    carrier = pres = None
    fails = []

    def ex(action):
        return app.Run("PPT_AI_Editor!ExecuteFromString",
                       json.dumps({"actions": [action], "verify_after": False}))

    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
        pres = app.Presentations.Add()
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540
        pres.Windows(1).Activate()
        time.sleep(1.0)

        # Presentations.Add() starts with 0 slides; templates append, so
        # derive indices from the live slide count rather than assuming.
        ex({"type": "apply_template", "template": "title",
            "content": {"title": "CAP_T", "subtitle": "CAP_S"}})
        src_idx = pres.Slides.Count

        # (a) capture -> registry file written with auto-slots
        ex({"type": "capture_template", "name": "t_cap", "slide": src_idx,
            "registry_path": reg})
        if not os.path.exists(reg):
            fails.append(("a", "registry file", f"missing {reg}"))
            return fails
        regj = json.loads(Path(reg).read_text(encoding="utf-8"))
        tpl = regj.get("templates", {}).get("t_cap")
        if not tpl:
            fails.append(("a", "registry entry", "t_cap absent"))
            return fails
        if tpl.get("slots") != ["heading", "body"]:
            fails.append(("a", "auto-slots", f"got {tpl.get('slots')!r}"))
        head_shape = next((s for s in tpl["shapes"]
                           if s.get("slot") == "heading"), None)

        # (b) generic render of the captured template w/ new content
        ex({"type": "apply_template", "template": "t_cap",
            "content": {"heading": "NEW_H", "body": "NEW_B"},
            "registry_path": reg})
        cap_idx = pres.Slides.Count
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        cap_slide = snap["slides"][cap_idx - 1]
        ts = " || ".join(texts_of(cap_slide))
        for sen in ("NEW_H", "NEW_B"):
            if sen not in ts:
                fails.append(("b", "slot text", f"{sen!r} missing"))
        if head_shape:
            p = find_pos(cap_slide, "NEW_H")
            if not p:
                fails.append(("b", "geometry", "no NEW_H shape"))
            else:
                for k in ("left", "top", "width"):
                    if abs(p.get(k, -999) - head_shape[k]) > TOL:
                        fails.append(("b", f"geom {k}",
                                      f"got {p.get(k)} want ~{head_shape[k]}"))

        # variants across [captured, builtin]
        # t_cap is title-derived; pair it with a structurally DIFFERENT
        # builtin (quote) so the two variants are genuinely distinct.
        ex({"type": "generate_variants",
            "templates": ["t_cap", "quote"],
            "content": {"heading": "VAR_X", "body": "VAR_Y",
                        "quote_text": "VAR_X", "attribution": "VAR_Y"},
            "registry_path": reg})
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        v = snap["slides"][-2:]
        if len(v) != 2:
            fails.append(("variants", "count", f"{len(v)} new"))
        else:
            for i, sl in enumerate(v):
                if "VAR_X" not in " || ".join(texts_of(sl)):
                    fails.append(("variants", f"slide{i}", "VAR_X missing"))
            if texts_of(v[0]) == texts_of(v[1]) and \
               [s.get("pos") for s in v[0]["shapes"]] == \
               [s.get("pos") for s in v[1]["shapes"]]:
                fails.append(("variants", "distinct", "two variants identical"))

        # (c) live manifest
        man = app.Run("PPT_AI_Editor!BuildCapturedManifest", reg)
        for needle in ("t_cap", "slots:", "heading", "apply_template"):
            if needle not in (man or ""):
                fails.append(("c", "manifest", f"{needle!r} not in manifest"))

        # safe delete: gone from registry+manifest, built slide UNCHANGED
        before = texts_of(json.loads(
            app.Run("PPT_AI_Editor!BuildSnapshotJson"))["slides"][cap_idx - 1])
        ex({"type": "delete_template", "name": "t_cap", "registry_path": reg})
        regj2 = json.loads(Path(reg).read_text(encoding="utf-8"))
        if "t_cap" in regj2.get("templates", {}):
            fails.append(("delete", "registry", "t_cap still present"))
        man2 = app.Run("PPT_AI_Editor!BuildCapturedManifest", reg)
        if "t_cap" in (man2 or ""):
            fails.append(("delete", "manifest", "t_cap still advertised"))
        after = texts_of(json.loads(
            app.Run("PPT_AI_Editor!BuildSnapshotJson"))["slides"][cap_idx - 1])
        if before != after:
            fails.append(("delete", "built slide changed",
                          f"{before!r} -> {after!r}"))
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
        print(f"FAIL: capture run failed after retries: {last!r}")
        return 1
    fails = result
    print("checks: (a) data registry  (b) generic render+variants  "
          "(c) live manifest  + safe delete")
    if fails:
        for stage, chk, why in fails:
            print(f"  FAIL [{stage}] {chk}: {why}")
    else:
        print("  capture->registry; generic render round-trips geometry+slots; "
              "variants distinct; manifest live; delete safe")
    ok = not fails
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

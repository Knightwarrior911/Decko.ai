"""Dialogs + icon-trim smoke test (deterministic), via PowerPoint COM.

Run: python tests/run_smoke_dialogs.py   (exit 0 only at 100%)

(a) icon trim: the exported prompt keeps concise CDN-sourcing guidance +
    curated names and does NOT inject the exhaustive allow-list.
(b) static src/modUI.bas: ManageTemplates is numbered + says DELETE /
    Cancel; CaptureTemplate reworded; both still call modActionsCapture.
(c) capture + delete still work end-to-end via the action path
    (numbered list reflects add then remove).

Visual black/robot theming is intentionally NOT covered here (deferred
goal 3; a VBA InputBox cannot be themed). No AI/network.
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

REPO = Path(__file__).resolve().parent.parent
CARRIER = REPO / "PPT_AI_Editor.pptm"


def static_modui_checks():
    src = (REPO / "src" / "modUI.bas").read_text(encoding="utf-8",
                                                 errors="replace")
    f = []
    need = [
        ("ManageTemplates DELETE wording", "To DELETE one" in src),
        ("ManageTemplates Cancel wording", "press Cancel" in src),
        ("numbered list helper used", "NumberedTemplateList" in src),
        ("capture prompt reworded",
         "Name this captured template" in src),
        ("old capture wording gone",
         'InputBox("Template name (saved' not in src),
        ("calls capture action",
         "modActionsCapture.Do_capture_template_act" in src),
        ("calls delete action",
         "modActionsCapture.Do_delete_template_act" in src),
    ]
    for label, ok in need:
        if not ok:
            f.append(("b-static", label, "failed"))
    return f


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
    fails = list(static_modui_checks())
    big = ""
    at = REPO / "data" / "icons_allowed.txt"
    if at.exists():
        big = at.read_text(encoding="utf-8", errors="replace").strip()
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)

        # (a) icon trim
        pp = os.path.join(tempfile.mkdtemp(prefix="dlg_"), "p.txt")
        app.Run("PPT_AI_Editor!DumpPromptToFile", pp)
        prm = Path(pp).read_text(encoding="utf-8", errors="replace")
        a = [
            ("ICON ACTION present", "ICON ACTION" in prm),
            ("HARD RULE exact-name", "HARD RULE" in prm
             and "EXACT Microsoft Fluent UI" in prm),
            ("CDN browse URL",
             "unpkg.com/browse/@fluentui/svg-icons/icons/" in prm),
            ("filename pattern", "{name}_{size}_{style}.svg" in prm),
            ("curated names",
             "Common valid names" in prm and "building_factory" in prm
             and "people" in prm and "globe" in prm),
            ("semantic fallback", "semantic" in prm.lower()),
            ("exhaustive list NOT injected",
             "FLUENT UI ICON ALLOW-LIST" not in prm
             and "All 830 names" not in prm
             and (len(big) < 200 or big not in prm)),
        ]
        for label, ok in a:
            if not ok:
                fails.append(("a-icon", label, "failed"))

        # (c) capture + delete still work via the action path
        tmpreg = os.path.join(tempfile.mkdtemp(prefix="dlgreg_"),
                              "templates.json")
        pres = app.Presentations.Add()
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540
        pres.Windows(1).Activate()
        time.sleep(1.0)

        def ex(a_):
            return app.Run("PPT_AI_Editor!ExecuteFromString",
                           json.dumps({"actions": [a_],
                                       "verify_after": False}))
        ex({"type": "apply_template", "template": "title",
            "content": {"title": "DLG_T", "subtitle": "DLG_S"}})
        sidx = pres.Slides.Count
        ex({"type": "capture_template", "name": "dlg_cap",
            "slide": sidx, "registry_path": tmpreg})
        lst1 = app.Run("PPT_AI_Editor!NumberedTemplateList", tmpreg)
        if "dlg_cap" not in (lst1 or ""):
            fails.append(("c-flow", "after capture",
                          f"dlg_cap not listed: {lst1!r}"))
        if "1. dlg_cap" not in (lst1 or ""):
            fails.append(("c-flow", "numbered format",
                          f"expected '1. dlg_cap' in {lst1!r}"))
        ex({"type": "delete_template", "name": "dlg_cap",
            "registry_path": tmpreg})
        lst2 = app.Run("PPT_AI_Editor!NumberedTemplateList", tmpreg)
        if "dlg_cap" in (lst2 or ""):
            fails.append(("c-flow", "after delete",
                          f"dlg_cap still listed: {lst2!r}"))
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
        print(f"FAIL: dialogs run failed after retries: {last!r}")
        return 1
    fails = result
    print("checks: (a) icon trim  (b) modUI wording  (c) capture/delete flow")
    if fails:
        for stage, label, why in fails:
            print(f"  FAIL [{stage}] {label}: {why}")
    else:
        print("  icon list trimmed; dialogs numbered + DELETE/Cancel wording; "
              "capture/delete still work")
    ok = not fails
    print("\nRESULT:", "PASS" if ok else "FAIL")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())

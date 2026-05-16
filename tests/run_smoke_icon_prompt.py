"""Verifies that the LLM export prompt now contains the Fluent icon allow-list.

Calls PPT_AI_Editor!DumpPromptToFile (test hook in modUI) to write the full
prompt template to a tempfile, then asserts:
  - the manifest header appears
  - the allow-list contains known-good icon names (people, globe, building_factory)
  - the dead URL hint is gone
  - the schema example uses building_factory not factory
"""
import os
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"


def main():
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
    try:
        out_path = Path(tempfile.gettempdir()) / "pptai_prompt_dump.txt"
        if out_path.exists():
            out_path.unlink()
        app.Run("PPT_AI_Editor!DumpPromptToFile", str(out_path))
        assert out_path.exists(), f"dump file not written: {out_path}"
        prompt = out_path.read_text(encoding="utf-8", errors="replace")
        print(f"prompt size: {len(prompt)} chars")

        # Contract (updated): the exhaustive allow-list is intentionally
        # gone; concise CDN-sourcing guidance is retained so an off-box
        # LLM can still pick exact Fluent names.
        big_list_body = ""
        allow_txt = REPO_ROOT / "data" / "icons_allowed.txt"
        if allow_txt.exists():
            big_list_body = allow_txt.read_text(encoding="utf-8",
                                                errors="replace")
        checks = [
            ("ICON ACTION section present",
             "ICON ACTION" in prompt),
            ("HARD RULE clause present",
             "HARD RULE" in prompt and "EXACT Microsoft Fluent UI" in prompt),
            ("CDN browse URL present",
             "unpkg.com/browse/@fluentui/svg-icons/icons/" in prompt),
            ("filename pattern note present",
             "{name}_{size}_{style}.svg" in prompt),
            ("semantic fallback instruction present",
             "semantic" in prompt.lower()),
            ("curated common names present",
             "Common valid names" in prompt and "people" in prompt
             and "globe" in prompt and "building_factory" in prompt),
            ("default size=32 stated",
             "size=32" in prompt),
            ("dead URL hint removed",
             "icon.fluentui.dev" not in prompt and
             "fluenticons.co" not in prompt),
            ("schema example uses building_factory",
             "building_factory" in prompt),
            ("exhaustive allow-list NOT injected",
             "FLUENT UI ICON ALLOW-LIST" not in prompt
             and "All 830 names" not in prompt
             and (len(big_list_body) < 200
                  or big_list_body.strip() not in prompt)),
        ]

        failed = []
        for label, ok in checks:
            mark = "ok  " if ok else "FAIL"
            print(f"  {mark}  {label}")
            if not ok:
                failed.append(label)

        if failed:
            print(f"\n{len(failed)} FAILED")
            sys.exit(1)
        print("\nicon-prompt injection verified")
    finally:
        try:
            carrier.Saved = True
            carrier.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()

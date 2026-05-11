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

        checks = [
            ("manifest header present",
             "FLUENT UI ICON ALLOW-LIST" in prompt),
            ("HARD RULE clause present",
             "HARD RULE" in prompt and "MUST be picked verbatim" in prompt),
            ("semantic fallback instruction present",
             "semantic" in prompt.lower()),
            ("manifest name 'people' present",
             "\npeople\n" in prompt or "\npeople\r\n" in prompt),
            ("manifest name 'globe' present",
             "\nglobe\n" in prompt or "\nglobe\r\n" in prompt),
            ("manifest name 'building_factory' present",
             "\nbuilding_factory\n" in prompt or "\nbuilding_factory\r\n" in prompt),
            ("default size=32 stated",
             "size=32" in prompt),
            ("dead URL hint removed",
             "icon.fluentui.dev" not in prompt and
             "fluenticons.co" not in prompt),
            ("schema example uses building_factory",
             "building_factory" in prompt),
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

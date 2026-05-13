"""
test_fix_errors_button.py — exercise the 'Fix Errors' button flow.

Build a JSON batch with deliberate validation errors, feed it through
BuildErrorFixPrompt, dump the clipboard-ready prompt that the LLM would see.
"""

import json
import sys
import time
from pathlib import Path

import win32com.client
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"


def main() -> int:
    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    carrier = app.Presentations.Open(str(CARRIER))
    # Need an active deck for ActivePresentation context inside validator
    test_pres = app.Presentations.Add()
    test_pres.Slides.Add(1, 12)  # ppLayoutBlank
    test_pres.Windows(1).Activate()

    # Deliberately malformed actions covering common LLM mistakes
    bad_batch = {
        "actions": [
            # 1: set_paragraph_text uses 'text' instead of 'value'
            {"type": "set_paragraph_text", "slide": 1, "shape_id": 5,
             "paragraph_index": 0, "text": "Hello"},
            # 2: add_paragraph missing after_paragraph_index
            {"type": "add_paragraph", "slide": 1, "shape_id": 5, "value": "New"},
            # 3: set_font_color missing value
            {"type": "set_font_color", "slide": 1, "shape_id": 5},
            # 4: add_shape missing pos
            {"type": "add_shape", "slide": 1, "kind": "rrect"},
            # 5: insert_icon with pos instead of flat coords
            {"type": "insert_icon", "slide": 1, "icon": "people",
             "pos": {"left": 60, "top": 120, "width": 48, "height": 48}},
            # 6: set_text on missing slide (out of range)
            {"type": "set_text", "slide": 99, "shape_id": 5, "value": "X"},
            # 7: VALID action — should NOT appear in fix prompt
            {"type": "set_text", "slide": 1, "shape_id": 5, "value": "Valid"},
        ]
    }
    batch_json = json.dumps(bad_batch)

    print("Calling BuildErrorFixPrompt with 6 invalid + 1 valid action...")
    t0 = time.perf_counter()
    prompt = app.Run(
        "PPT_AI_Editor.pptm!modExecuteInstructions.BuildErrorFixPrompt",
        batch_json,
    )
    elapsed_ms = (time.perf_counter() - t0) * 1000
    print(f"  built in {elapsed_ms:.1f} ms; prompt length = {len(prompt)} chars\n")

    print("=== Clipboard-ready prompt (what 'Fix Errors' would copy) ===\n")
    print(prompt)

    test_pres.Close()
    carrier.Close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

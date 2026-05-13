"""
test_fix_button.py — end-to-end test of the Fix This button flow.

1. Build deck with deliberate quality issues
2. Run verify via ExecuteFromString (writes warnings.json)
3. Trigger modVerify.CopyWarningsPromptToClipboard (what the button does)
4. Read clipboard back and print
"""

import json
import sys
import time
from pathlib import Path

import win32com.client
import win32clipboard
import pythoncom

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
TEST_DECK = REPO_ROOT / "tests" / "test_verify_deck.pptx"


def read_clipboard_text() -> str:
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT):
            return win32clipboard.GetClipboardData(win32clipboard.CF_TEXT).decode("mbcs", errors="replace")
        return ""
    finally:
        win32clipboard.CloseClipboard()


def main() -> int:
    if not TEST_DECK.exists():
        print(f"ERROR: {TEST_DECK} not found — run tests/test_verify_loop.py first.")
        return 1

    pythoncom.CoInitialize()
    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    while app.Presentations.Count > 0:
        try:
            app.Presentations(1).Close()
        except Exception:
            break

    carrier = app.Presentations.Open(str(CARRIER), WithWindow=False)
    test_pres = app.Presentations.Open(str(TEST_DECK))
    test_pres.Windows(1).Activate()

    print("Running verification to populate warnings.json...")
    instructions = json.dumps({"actions": [], "verify_after": True, "verify_scope": "deck"})
    result = app.Run("PPT_AI_Editor.pptm!modExecuteInstructions.ExecuteFromString", instructions)
    print(f"  {result}")

    print("\nTriggering modVerify.CopyWarningsPromptToClipboard (what 'Fix This' button calls)...")
    t0 = time.perf_counter()
    n_copied = app.Run("PPT_AI_Editor.pptm!modVerify.CopyWarningsPromptToClipboard")
    elapsed_ms = (time.perf_counter() - t0) * 1000
    print(f"  reported {n_copied} warning(s) copied in {elapsed_ms:.1f} ms")

    print("\n=== Clipboard contents (what gets pasted into LLM) ===\n")
    clip = read_clipboard_text()
    # Truncate for readable output; full thing is on the clipboard
    if len(clip) > 4000:
        print(clip[:2000])
        print(f"\n... [{len(clip)-2000} more chars on clipboard] ...\n")
    else:
        print(clip)
    print(f"\nClipboard total length: {len(clip)} chars")

    test_pres.Close()
    carrier.Close()
    return 0


if __name__ == "__main__":
    sys.exit(main())

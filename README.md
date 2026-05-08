# PPT AI Editor

VBA-based PowerPoint editor driven by natural language via an external LLM.

See `docs/specs/2026-05-08-ppt-ai-editor-design.md` for design.

## Setup (developer machine, one-time)

1. Install Python 3.10+ and run: `pip install -r requirements.txt`
2. In PowerPoint: File → Options → Trust Center → Trust Center Settings → Macro Settings → check **Trust access to the VBA project object model**.
3. Run `python tools/build_carrier.py` to bootstrap `PPT_AI_Editor.pptm`.
4. Run `python update_macros.py` to sync `src/` modules into the carrier.

## Usage

1. Open the deck you want to edit, plus `PPT_AI_Editor.pptm`.
2. Alt+F8 → `ExportSnapshot` → click **Copy snapshot + prompt template**.
3. Paste into your LLM tool, describe the change you want.
4. Copy the LLM's instructions JSON.
5. Alt+F8 → `ExecuteInstructions` → paste, click **Parse**, review, click **Apply**.

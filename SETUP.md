# Decko.ai — Fresh-Clone Setup

If you cloned this repo on a new machine (or fresh Claude session), follow
this exactly. Every step is needed; skipping one leaves the carrier in a
half-built state.

## Prerequisites (one-time per machine)

1. **Windows** with **Microsoft PowerPoint** installed (Decko drives
   PowerPoint via COM; no PowerPoint = no Decko). Tested on Windows 11 +
   PowerPoint 365 / 2021 / 2019.
2. **Python 3.10+** on PATH.
3. In PowerPoint:
   - File → Options → Trust Center → Trust Center Settings → Macro
     Settings → check **Trust access to the VBA project object model**.
   - Without this, `update_macros.py` and `tools/add_fix_button.py` cannot
     write into the carrier.

## Install steps

From the repo root:

```bash
# 1. Python deps (pywin32 ships pythoncom + win32clipboard automatically)
pip install -r requirements.txt

# 2. Sync VBA modules from src/ into the carrier .pptm
python update_macros.py

# 3. Install the two clipboard buttons on frmExecute.
#    Idempotent — re-running just repositions / refreshes handler code.
python tools/add_fix_button.py
```

After step 3 the carrier `PPT_AI_Editor.pptm` is fully functional. Open it
in PowerPoint and press Alt+F8 to see the three macros:

- `ExportSnapshot` — copy deck JSON snapshot + prompt template to clipboard
- `ExecuteInstructions` — paste LLM-returned actions and Apply
- `ImportSlides` — pull slides from another deck

## Verify the install worked

```bash
# Builds a deck with deliberate quality issues, runs verification, dumps
# warnings to console. Should report ~47 warnings in <500 ms.
python tests/test_verify_loop.py

# Builds a bad-actions batch, runs error-fix prompt builder, prints the
# clipboard-ready LLM prompt that Fix Errors would copy.
python tests/test_fix_errors_button.py

# Same flow for Fix This button (depends on test_verify_loop.py having run
# first to populate test_verify_deck.pptx).
python tests/test_fix_button.py
```

All three should run end-to-end without errors. If `test_verify_loop.py`
prints `25-50 warnings detected`, the install is good.

## What lives where

```
PPT_AI_Editor.pptm       carrier deck — opens in PowerPoint, holds macros
                         (regenerated from src/ by update_macros.py)

src/
  modUI.bas              Alt+F8 entry points
  modExportSnapshot.bas  Snapshot JSON builder
  modExecuteInstructions Parse / Validate / Dispatch / verify pipeline
  modBackup.bas          Auto-backup helper (runs on every Apply)
  modJSON.bas            JSON parser / encoder
  modVerify.bas          Post-Apply quality-check loop (32 checks)
  modActions*.bas        Action implementations (~165 types across 14 files)
  frmExecute.frm/.frx    "Execute Instructions" UserForm (Parse/Apply/
                         Fix Errors/Fix This buttons)
  frmExport.frm/.frx     "Export Snapshot" UserForm
  frmImportSlides.frm    "Import Slides" UserForm

docs/
  ACTIONS_REFERENCE.md   Complete schema for every action (165+)
  PROMPTING_GUIDE.md     How to write VP requests; LLM-facing
  EXAMPLES.md            Worked VP-prompt → actions-JSON pairs
  LAYOUT_RECIPES.md      Slide-layout cookbook (quad / 67-33 / etc.)
  USER_GUIDE.md          End-user walkthrough
  VERIFICATION.md        Quality-check loop + Fix buttons explained

tools/
  build_carrier.py       Bootstrap an empty .pptm carrier
  build_forms.py         Author the three UserForms programmatically
  add_fix_button.py      ONE-TIME installer for Fix Errors / Fix This
                         buttons on frmExecute (idempotent)
  precheck_carrier.py    Carrier sanity check

tests/
  test_verify_loop.py    End-to-end verify loop (builds problem deck)
  test_fix_button.py     End-to-end Fix This button flow
  test_fix_errors_button End-to-end Fix Errors button flow
  run_smoke*.py          Action-by-action smoke harnesses
  make_test_decks.py     Regenerate deterministic test decks

update_macros.py         Sync src/*.bas + *.frm into the carrier
```

## Common gotchas

- **"Sub or Function not defined: ParseJson"** on first run — you forgot
  `python update_macros.py`. modJSON.bas isn't in the carrier yet.
- **`Cannot edit Macro on a hidden workbook`** when running
  `add_fix_button.py` — close all PowerPoint windows first; the script
  expects a clean state.
- **Buttons missing from Execute form** — run `python tools/add_fix_button.py`
  once. After it runs, the .frm/.frx in src/ contains the buttons
  permanently; future `update_macros.py` runs preserve them.
- **Verify loop logs `warnings.json` next to a read-only deck** — Decko
  writes the sidecar to `<deck>.warnings.json`. If the directory is
  read-only or sync-locked (OneDrive), verify will silently skip the
  sidecar write but the warnings still appear in the form status bar.
- **`Trust access to the VBA project object model` not enabled** — every
  install script aborts with COM error 0x800A03EC. Enable it in PowerPoint
  Trust Center and retry.

## Updating after a `git pull`

```bash
git pull
python update_macros.py          # always safe; pulls fresh .bas/.frm
# tools/add_fix_button.py is idempotent — re-run only if frmExecute changed
```

## After adding / renaming an action in the dispatcher

Two files must stay in sync with `modExecuteInstructions.DispatchAction`:

1. **`GetActionGuidance` Case table** — canonical signature + example for
   the new action. Coverage test `tests/test_guidance_coverage.py` fails if
   you forget.
2. **`GetAllActionTypes` master list** — used by `FindSimilarActions` for
   "did you mean" suggestions on `unknown_type` errors.
3. **`docs/ACTIONS_REFERENCE.md` auto-appendix** — auto-regenerated from
   GetActionGuidance. Run:

   ```bash
   python tools/sync_actions_guidance.py
   ```

   Drift test `tests/test_guidance_doc_sync.py` fails if you forget.

Both tests run end-to-end via PowerPoint COM (~10-20 s each).

## Cleaning up

To wipe the carrier and rebuild from scratch:

```bash
rm PPT_AI_Editor.pptm
python tools/build_carrier.py
python tools/build_forms.py
python update_macros.py
python tools/add_fix_button.py
```

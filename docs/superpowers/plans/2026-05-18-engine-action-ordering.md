# Engine Action-Ordering & Guidance-Contract — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. NOTE: if subagent dispatch is model-usage-limited, the controller executes tasks directly with the SAME deterministic-gate verification + two-stage controller review.

**Goal:** Make the weak LLM sequence multi-step actions right first try by (A) eliminating every `GetActionGuidance`↔`ValidateAction` drift (root cause of the `add_run`/superscript dead-end) behind a permanent static contract test, and (B) adding run-model + ordering steering to the prompt and per-action guidance.

**Architecture:** Pure-static Python parse of `src/modExecuteInstructions.bas` — no COM, no engine-execution change. The contract test is the gate; guidance/prompt are corrected to match the validator (validator = source of truth). Then `update_macros.py` + `sync_actions_guidance.py` keep the carrier and `ACTIONS_REFERENCE.md` in lockstep.

**Tech Stack:** Python 3.11 (regex, json, pytest-free plain script like the other `run_smoke_*`), VBA (string-only edits to `GetActionGuidance`/`PromptTemplate`), `update_macros.py`, `tools/sync_actions_guidance.py`. Spec: `docs/superpowers/specs/2026-05-18-engine-action-ordering-design.md` (EAC1–EAC5 locked — never edit `ValidateAction`/`DispatchAction` logic).

---

## File Structure

```
tests/test_guidance_contract.py   NEW  static drift gate (parser + assertions + KNOWN_UNPARSEABLE)
src/modExecuteInstructions.bas    MOD  GetActionGuidance REQUIRED/EXAMPLE corrections + 4 ordering one-liners. NO Validate/Dispatch change.
src/frmExport.frm                 MOD  PromptTemplate: + MISTAKE 6 (run model) + ORDERING RECIPES block
docs/ACTIONS_REFERENCE.md         REGEN via tools/sync_actions_guidance.py
```

VBA invariant: only `GetActionGuidance` string literals + `PromptTemplate` concatenation change. `ValidateAction`, `RequireFields`, `DispatchAction`, every `Do_*` are untouched (verified by `git diff` per task).

---

## Task 1: The static contract test (gate; first run enumerates drift)

**Files:**
- Create: `tests/test_guidance_contract.py`

- [ ] **Step 1: Write the test**

`tests/test_guidance_contract.py`:
```python
"""Static guidance<->validator contract. No COM. For every action in
modExecuteInstructions.GetActionGuidance, assert its REQUIRED list +
EXAMPLE keys cover everything its ValidateAction RequireFields requires
(shape_id<->shape_name and explicit `act.Exists(a)/Exists(b)` groups
are alternatives). Exit 1 on any drift, misleading extra, or an
unparseable-case-set change. Source of truth = the validator."""
import json
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
BAS = (REPO / "src" / "modExecuteInstructions.bas").read_text(
    encoding="utf-8", errors="replace")

# Cases whose required fields are NOT a literal RequireFields(Array(...))
# — validated dynamically. Each must have a one-line reason. The test
# FAILS if the actually-unparseable set differs from this allowlist.
KNOWN_UNPARSEABLE = {
    "add_connector": "from/to validated via explicit act.Exists pairs",
}


def _case_blocks(func_src: str):
    """Yield (tuple_of_action_names, block_text) for each Case in a
    Select Case body. Handles grouped + line-continued Case labels."""
    # join VBA line continuations
    joined = re.sub(r"_\r?\n\s*", " ", func_src)
    parts = re.split(r"\n\s*Case ", joined)
    for p in parts[1:]:
        head, _, body = p.partition("\n")
        if head.strip().startswith("Else"):
            continue
        names = re.findall(r'"([^"]+)"', head)
        if names:
            yield tuple(names), body


def _slice_func(name: str) -> str:
    m = re.search(r"(?:Public |Private )?Function " + re.escape(name)
                  + r"\b.*?\nEnd Function", BAS, re.S)
    assert m, f"function {name} not found"
    return m.group(0)


def _guidance_map():
    """action -> (set(required_tokens), set(example_keys))."""
    out = {}
    src = "".join(_slice_func(f) for f in
                  ("GetActionGuidance", "GetActionGuidance_Part2",
                   "GetActionGuidance_Part3")
                  if re.search(r"Function " + f + r"\b", BAS))
    for names, body in _case_blocks(src):
        text = body
        req = set()
        m = re.search(r"REQUIRED:\s*([^\"]*)", text)
        if m:
            for tok in m.group(1).split(","):
                t = re.sub(r"\(.*?\)", "", tok).strip()
                t = re.split(r"\s", t)[0].strip()
                if t and t.lower() not in ("none", "n/a", "-"):
                    req.add(t)
        ex = set()
        em = re.search(r"EXAMPLE:\s*(\{.*?\})", text)
        if em:
            raw = em.group(1).replace('""', '"')
            try:
                ex = set(json.loads(raw).keys())
            except Exception:
                ex = set(re.findall(r'""(\w+)""\s*:', body))
        for a in names:
            out[a] = (req, ex)
    return out


def _validator_map():
    """action -> set(required) or None (unparseable)."""
    out = {}
    va = _slice_func("ValidateAction")
    for names, body in _case_blocks(va):
        block = body.split("\n        Case ")[0]
        req = None
        rf = re.search(r"RequireFields\(act,\s*Array\(([^)]*)\)\)",
                       block, re.S)
        if rf:
            req = set(re.findall(r'"([^"]+)"', rf.group(1)))
        else:
            pairs = re.findall(
                r'Not act\.Exists\("([^"]+)"\)\s+And\s+Not '
                r'act\.Exists\("([^"]+)"\)', block)
            if pairs:
                req = set()
                for a, b in pairs:
                    req.add(a + "|" + b)
        for a in names:
            out[a] = req
    return out


def _satisfied(field: str, have: set) -> bool:
    if "|" in field:                       # explicit alt group
        return any(p in have for p in field.split("|"))
    if field in have:
        return True
    if "shape_ids" in field:
        return field.replace("shape_ids", "shape_names") in have
    if "shape_id" in field:
        return field.replace("shape_id", "shape_name") in have
    return False


def main() -> int:
    g = _guidance_map()
    v = _validator_map()
    fails, unparseable = [], set()
    for act, req in v.items():
        if req is None:
            unparseable.add(act)
            continue
        if act not in g:
            fails.append(f"{act}: no GetActionGuidance case")
            continue
        greq, gex = g[act]
        have = greq | {x for f in greq for x in f.split("|")}
        for f in req:
            if not _satisfied(f, have):
                fails.append(f"{act}: validator requires '{f}' but "
                             f"guidance REQUIRED={sorted(greq)}")
            if not _satisfied(f, gex):
                fails.append(f"{act}: validator requires '{f}' but "
                             f"EXAMPLE keys={sorted(gex)}")
        # misleading extra: a guidance-required token the validator
        # does not require and is not an alternative of one
        valt = set()
        for f in req:
            valt.add(f)
            if "shape_id" in f:
                valt.add(f.replace("shape_id", "shape_name"))
            if "shape_ids" in f:
                valt.add(f.replace("shape_ids", "shape_names"))
            valt |= set(f.split("|"))
        for t in greq:
            if t not in valt and not any(
                    t in f.split("|") or
                    t == f.replace("shape_id", "shape_name") or
                    t == f.replace("shape_ids", "shape_names")
                    for f in req):
                fails.append(f"{act}: guidance REQUIRED lists '{t}' "
                             f"which the validator does NOT require")

    expected_unp = set(KNOWN_UNPARSEABLE)
    if unparseable != expected_unp:
        fails.append(
            f"unparseable-case set changed: got {sorted(unparseable)}, "
            f"allowlist {sorted(expected_unp)} — review & update "
            f"KNOWN_UNPARSEABLE")

    if fails:
        print(f"GUIDANCE CONTRACT: {len(fails)} issue(s)")
        for f in fails:
            print("  FAIL", f)
        print("\nRESULT: FAIL")
        return 1
    print(f"GUIDANCE CONTRACT: {len(v)} actions, "
          f"{len(KNOWN_UNPARSEABLE)} allowlisted-unparseable, 0 drift")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run it — enumerate the drift (expected FAIL first)**

Run: `python tests/test_guidance_contract.py ; echo "EXIT=$?"`
Expected: `RESULT: FAIL`, EXIT=1, with a printed list that INCLUDES
`add_run: validator requires 'run_index' but guidance REQUIRED=...`
plus any other drifted actions. **Copy the full FAIL list into the
task report — it is the audit output Task 2 fixes.**
Also confirm the `unparseable-case set` line: if it reports actions
beyond `add_connector`, add each to `KNOWN_UNPARSEABLE` WITH a real
one-line reason (read that case in the .bas to confirm it is truly
dynamic, not a parser miss), re-run until the only `unparseable`
discrepancy is gone. Do NOT allowlist a case that is really a
`RequireFields` the parser failed to read — fix the parser instead.

- [ ] **Step 3: Commit the test (red is expected here)**

```bash
git add tests/test_guidance_contract.py
git commit -m "$(cat <<'EOF'
test(engine): static GetActionGuidance<->ValidateAction contract gate

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: Fix every guidance↔validator drift

**Files:**
- Modify: `src/modExecuteInstructions.bas` (ONLY inside `GetActionGuidance`/`GetActionGuidance_Part*` string literals — never `ValidateAction`/`RequireFields`/`DispatchAction`)

- [ ] **Step 1: Re-run the gate to get the live list**

Run: `python tests/test_guidance_contract.py`
For EACH `FAIL <action>: validator requires '<f>' but guidance
REQUIRED=...` / `EXAMPLE keys=...`:

- [ ] **Step 2: For each drifted action, correct its guidance**

In `src/modExecuteInstructions.bas`, find `Case "<action>"` inside
`GetActionGuidance*`. Edit its `REQUIRED:` line to include every
validator-required token (use the validator's exact field name; for a
`shape_id` field you may write `shape_id`), and edit its `EXAMPLE:`
JSON to include a key for every required token with a realistic value.
Concrete first fix (`add_run`, lines ~2814-2818) — validator requires
`slide, shape_id, paragraph_index, run_index, value`:
```vba
        Case "add_run"
            GetActionGuidance = _
                "  REQUIRED: slide, shape_id, paragraph_index, run_index, value(string)" & vbCrLf & _
                "  ORDER: set the paragraph text FIRST; add_run INSERTS a new run at run_index (0-based) without destroying existing runs. Then target that run_index with set_run_* actions." & vbCrLf & _
                "  OPTIONAL: bold(bool), italic(bool), underline(bool), color(#RRGGBB), font_name(string), font_size(int)" & vbCrLf & _
                "  EXAMPLE:  {""type"":""add_run"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""run_index"":1,""value"":""3""}"
```
Apply the same discipline (REQUIRED ⊇ validator, EXAMPLE has every
required key) to every other action the gate flagged. Keep each edit
inside the existing `_ &` concatenation; add no module-level
declarations (compile-wedge risk).

- [ ] **Step 3: Gate green**

Run: `python tests/test_guidance_contract.py ; echo EXIT=$?`
Expected: `RESULT: PASS`, EXIT=0.

- [ ] **Step 4: Confirm no logic touched**

Run: `git diff -U0 src/modExecuteInstructions.bas | grep -E "^\+|^-" | grep -iE "RequireFields|Function ValidateAction|DispatchAction|Do_[a-z]" || echo "guidance-only OK"`
Expected: `guidance-only OK` (no validator/dispatch lines changed).

- [ ] **Step 5: Commit**

```bash
git add src/modExecuteInstructions.bas
git commit -m "$(cat <<'EOF'
fix(engine): correct all GetActionGuidance<->validator drift

add_run (+ every other drifted action): REQUIRED/EXAMPLE now match
ValidateAction RequireFields. Guidance strings only; no Validate/
Dispatch/Do_* change. Contract gate green.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: PromptTemplate — MISTAKE 6 + ORDERING RECIPES

**Files:**
- Modify: `src/frmExport.frm` (inside `PromptTemplate` `s = s & "..."` concatenation only)

- [ ] **Step 1: Locate the insertion point**

Run: `grep -n 'MISTAKE 5 -\|=============================' src/frmExport.frm | head`
Insert MISTAKE 6 immediately AFTER the MISTAKE 5 block and BEFORE the
closing `=============================` separator of the CRITICAL
MISTAKES section (same `s = s & "..." & vbCrLf` style).

- [ ] **Step 2: Add MISTAKE 6 + ORDERING RECIPES**

Add (exact lines, matching surrounding escaping — `""` for quotes):
```vba
    s = s & "MISTAKE 6 - Formatting PART of a text without a separate run." & vbCrLf
    s = s & "  Superscript/subscript/bold-word/color-word apply to ONE run." & vbCrLf
    s = s & "  set_text/set_paragraph_text make a SINGLE run, so there is no" & vbCrLf
    s = s & "  second run to target. To format only part of the text:" & vbCrLf
    s = s & "  1) set_text (or set_paragraph_text) with the BASE text only," & vbCrLf
    s = s & "  2) add_run to append the new run (REQUIRES paragraph_index AND" & vbCrLf
    s = s & "     run_index, both 0-based; run_index = the new run's index)," & vbCrLf
    s = s & "  3) set_run_superscript/set_run_* on that paragraph_index+run_index." & vbCrLf
    s = s & "  WRONG: set_text value:""53"" then set_run_superscript run_index:1 (only run 0 exists)" & vbCrLf
    s = s & "  RIGHT: set_text value:""5""  ->  add_run paragraph_index:0 run_index:1 value:""3""" & vbCrLf
    s = s & "         ->  set_run_superscript slide:S shape_id:I paragraph_index:0 run_index:1 value:true" & vbCrLf & vbCrLf
    s = s & "ORDERING RECIPES (emit actions in THIS order or they fail):" & vbCrLf
    s = s & "  - Partial formatting: set_text/set_paragraph_text -> add_run -> set_run_*  (see MISTAKE 6)" & vbCrLf
    s = s & "  - Same shape: ALL add_paragraph FIRST, then any set_bullet_style /" & vbCrLf
    s = s & "    set_indent_level / set_run_* / add_run (add_paragraph rebuilds the" & vbCrLf
    s = s & "    text frame and destroys earlier run/paragraph formatting)." & vbCrLf
    s = s & "  - Create before reference: add_shape(ref_name) before add_connector;" & vbCrLf
    s = s & "    add_table before set_cell_text; add_slide before slide-scoped ops." & vbCrLf
    s = s & "  - Deleting paragraphs: delete from HIGHEST index DOWN to lowest." & vbCrLf
    s = s & "  - Multiple find_replace_text on overlapping text: longest string first." & vbCrLf & vbCrLf
```

- [ ] **Step 3: Contract gate unaffected + smoke**

Run: `python tests/test_guidance_contract.py` → `RESULT: PASS`
(PromptTemplate is not parsed by the contract test; this just confirms
no accidental cross-file breakage).

- [ ] **Step 4: Commit**

```bash
git add src/frmExport.frm
git commit -m "$(cat <<'EOF'
feat(engine): PromptTemplate MISTAKE 6 (run model) + ORDERING RECIPES

Steers the weak LLM to set_text -> add_run -> set_run_superscript and
the other order-critical sequences, first try.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: Per-action ordering one-liners in GetActionGuidance

**Files:**
- Modify: `src/modExecuteInstructions.bas` (`GetActionGuidance*` strings only)

- [ ] **Step 1: Add an `ORDER:` line to four entries**

For `set_run_superscript`, `set_run_subscript`, `add_paragraph` (and
verify `add_run` already got its `ORDER:` line in Task 2), add ONE
`& vbCrLf & "  ORDER: ..."` line inside that case's concatenation:
- `set_run_superscript` / `set_run_subscript`:
  `"  ORDER: the target run must already exist. If only one run exists, add_run first, then set this on the new run_index."`
- `add_paragraph`:
  `"  ORDER: emit ALL add_paragraph for a shape BEFORE any set_run_*/set_bullet_style/set_indent_level/add_run on it; add_paragraph rebuilds the text frame."`
Keep inside the existing `_ &` chain; do not alter REQUIRED/EXAMPLE
text (the contract gate must stay green).

- [ ] **Step 2: Gate + logic-untouched check**

Run: `python tests/test_guidance_contract.py` → `RESULT: PASS`
Run: `git diff -U0 src/modExecuteInstructions.bas | grep -iE "RequireFields|Function ValidateAction|DispatchAction|Do_[a-z]" || echo "guidance-only OK"` → `guidance-only OK`

- [ ] **Step 3: Commit**

```bash
git add src/modExecuteInstructions.bas
git commit -m "$(cat <<'EOF'
feat(engine): per-action ORDER hints (run/superscript/add_paragraph)

So the in-app Fix Errors path also carries the sequencing rule.

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: Sync carrier + regenerate ACTIONS_REFERENCE + regression

**Files:**
- Modify: `docs/ACTIONS_REFERENCE.md` (regenerated by the tool)
- (carrier `PPT_AI_Editor.pptm` rebaked by `update_macros.py`)

- [ ] **Step 1: Bake src into the carrier**

Run: `taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2; python update_macros.py`
Expected: `[done]`, exit 0.

- [ ] **Step 2: Carrier-open compile smoke (catches a VBA wedge)**

Run:
```bash
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2
python -c "import win32com.client,pythoncom; pythoncom.CoInitialize(); a=win32com.client.DispatchEx('PowerPoint.Application'); a.Visible=True; p=a.Presentations.Open(r'C:\Users\vinit\Documents\PPT_AI_Editor\PPT_AI_Editor.pptm',WithWindow=True); n=len(a.Run('PPT_AI_Editor!GetAllActionTypes')); p.Saved=True; p.Close(); a.Quit(); print('GAT len',n)"
```
Expected: `GAT len` > 1000 (project compiles; no modal). If it wedges:
a string edit broke a line-continuation — fix the offending
`GetActionGuidance`/`PromptTemplate` line, re-run Step 1-2.

- [ ] **Step 3: Regenerate the public reference**

Run: `taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2; python tools/sync_actions_guidance.py`
Expected: rewrites `docs/ACTIONS_REFERENCE.md` AUTO-GUIDANCE block, exit 0.

- [ ] **Step 4: Regression gate (sequential, single PowerPoint at a time)**

```bash
python tests/test_guidance_contract.py            # RESULT: PASS
python tests/test_guidance_coverage.py            # all covered
python tests/test_example_validity.py             # all examples valid
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/test_guidance_doc_sync.py   # in sync
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke_guidance.py       # PASS
```
All must pass. (`taskkill` between the COM ones — single instance.)

- [ ] **Step 5: Restore carrier if only rebaked (no src delta) / commit reference**

`PPT_AI_Editor.pptm` legitimately changed (src guidance changed) — keep
it. Commit the rebaked carrier + regenerated reference together:
```bash
git add docs/ACTIONS_REFERENCE.md PPT_AI_Editor.pptm
git commit -m "$(cat <<'EOF'
chore(engine): rebake carrier + regen ACTIONS_REFERENCE after guidance fixes

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: Final gate + manual acceptance + README note

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Full deterministic gate (one shot, sequential)**

```bash
python tests/test_guidance_contract.py            # RESULT: PASS, EXIT 0
python tests/test_guidance_coverage.py
python tests/test_example_validity.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/test_guidance_doc_sync.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke_guidance.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke.py
```
All green. Confirm engine logic untouched across the branch:
`git diff origin/main..HEAD -- src/ | grep -iE '^\+.*\b(RequireFields|DispatchAction)\b|Function ValidateAction' || echo "no validator/dispatch logic changed"` → `no validator/dispatch logic changed`.

- [ ] **Step 2: Add README note**

Under the Verification / docs section of `README.md`, add:
```markdown
### Guidance contract (action ordering)

`python tests/test_guidance_contract.py` statically proves every
action's `GetActionGuidance` REQUIRED + EXAMPLE matches its
`ValidateAction` required fields (so a weak LLM is never told to omit
a field the validator demands). The prompt also carries MISTAKE 6
(run model: set_text → add_run → set_run_superscript) and an ORDERING
RECIPES block. Validator stays the source of truth; guidance is never
allowed to drift from it again.
```

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "$(cat <<'EOF'
docs: README — guidance contract + ordering steering

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

- [ ] **Step 4: Manual acceptance (honest scope — not automatable here)**

Document for the user to run once (paste-JSON or desktop chat, real
LLM): prompt = "in a box that says '5', add '3' then make the '3'
superscript". Expected correct batch the steered LLM should now emit
(order matters):
```json
{"actions":[
  {"type":"set_text","slide":1,"shape_id":<id>,"value":"5"},
  {"type":"add_run","slide":1,"shape_id":<id>,"paragraph_index":0,"run_index":1,"value":"3"},
  {"type":"set_run_superscript","slide":1,"shape_id":<id>,"paragraph_index":0,"run_index":1,"value":true}
]}
```
Record the actual LLM output in the task report; pass = it sequences
set_text → add_run(run_index) → set_run_superscript without dead-ending.

---

## Self-Review

**Spec coverage:** EAC1 (A+B, no execution change) — T1/T2 (A: contract + fixes, guidance-only diff check), T3/T4 (B: prompt + per-action steering). EAC2 (static, no COM, no refactor) — T1 parser is pure file parse; COM only in T5 compile-smoke/regression (not the contract gate). EAC3 (validator = truth) — T2 Step 2 fixes guidance to the validator, Step 4 asserts no validator/dispatch change. EAC4 (auto-mirror + tests green) — T5 sync + regression. EAC5 (one sub-project, A+B together) — single plan. Spec §3 precise rules (required/alternatives/extras/unparseable allowlist) — implemented in T1 `_satisfied`, extra-token check, `KNOWN_UNPARSEABLE` + set-change failure. §5 risks (compile wedge) — T5 Step 2 carrier smoke. §6 metric — T6 Step 1. Honest-scope manual test — T6 Step 4. No gaps.

**Placeholder scan:** No TBD/TODO. The T1 test code is complete and runnable; the dead `if False` scaffolding line was removed inline from `_validator_map` (it now starts at `va = _slice_func("ValidateAction")`).

**Type consistency:** `_guidance_map`→`{action:(set,set)}`, `_validator_map`→`{action:set|None}`, `_satisfied(field,have)`→bool, `KNOWN_UNPARSEABLE` dict — all consistent across T1 and referenced identically in later tasks (T2/T4 just re-run `tests/test_guidance_contract.py`). Commands/paths consistent (`python tests/test_guidance_contract.py`, `update_macros.py`, `tools/sync_actions_guidance.py`). No drift.

Fixed inline: the dead `if False` scaffolding line is explicitly ordered removed in T1.

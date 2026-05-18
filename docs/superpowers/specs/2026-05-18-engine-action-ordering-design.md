# Engine Action-Ordering & Guidance-Contract — Design Spec

Date: 2026-05-18
Status: Approved (design); pending implementation plan
Track: `.pptm` engine (Decko original) — separate from the SP-series desktop app.

## 1. Background

A weak external LLM (e.g. `deepseek-v4-*`, the work LLM) predictably
botches multi-step action ORDER and dead-ends on capabilities it does
not realize exist. Reproducible user case: "add text '3' into a box
that already says '5', then make the '3' superscript." The LLM produced
a single run `"53"` (no `run_index:1` to target), believed no `add_run`
action existed, and gave up.

**Root cause found during context exploration:** `add_run`'s
`ValidateAction` (`src/modExecuteInstructions.bas:641`) REQUIRES
`slide, shape_id, paragraph_index, run_index, value`, but its
`GetActionGuidance` REQUIRED line + EXAMPLE
(`src/modExecuteInstructions.bas:2814-2818`) OMIT `run_index`. A model
that follows the guidance emits `add_run` without `run_index`, the
validator rejects it, and the model concludes `add_run` is unusable.
This is a **guidance↔validator drift bug**, not a missing feature.
The prompt's MISTAKE blocks (1-5) also contain nothing about the run
model, `add_run`, superscript-needs-its-own-run, or run-op ordering.

## 2. Locked decisions

| # | Decision |
|---|----------|
| EAC1 | Scope = **A** (guidance↔validator drift audit + fix + regression test) **+ B** (prompt/guidance ordering & run steering). **No engine execution-path change** — only guidance *strings* + one new static test. |
| EAC2 | Approach 1 — a pure-static Python parse of `src/modExecuteInstructions.bas`; deterministic contract test; no COM; no refactor of `ValidateAction`/`DispatchAction` logic. |
| EAC3 | `ValidateAction` is the source of truth. Fix drift by correcting `GetActionGuidance`/EXAMPLE to match it — never by changing the validator. |
| EAC4 | Guidance text lives in `src/frmExport.frm` (`PromptTemplate`) + `src/modExecuteInstructions.bas` (`GetActionGuidance`); `docs/ACTIONS_REFERENCE.md` stays auto-mirrored via `tools/sync_actions_guidance.py`; existing drift/coverage/example tests must stay green. |
| EAC5 | A + B ship together as ONE cohesive sub-project (B's steering references actions whose guidance A corrects). Single spec/plan. |

## 3. Architecture

### A — Guidance↔validator contract

New deterministic test `tests/test_guidance_contract.py` (pure static,
no COM). It parses `src/modExecuteInstructions.bas` and, for every
dispatched action discovered in `DispatchAction`:

1. Resolves the action's validator-required field set from the
   matching `ValidateAction` `Case` — the `RequireFields(act,
   Array("a","b",...))` tokens, including multi-line `Array(...)`,
   grouped `Case "x","y", _` blocks, and alternative groups the engine
   accepts (e.g. `shape_id` OR `shape_name`, `path` OR `picture_path`).
2. Resolves the action's `GetActionGuidance` `REQUIRED:` token list and
   the JSON keys in its `EXAMPLE:` line.
3. Asserts: every validator-required field (or one member of an
   alternative group) appears in BOTH the guidance `REQUIRED:` list and
   the `EXAMPLE`. Reports, separately, guidance `REQUIRED` tokens that
   the validator does not require (misleading extras).
4. Emits a non-zero exit on any mismatch. **Explicitly lists every
   action `Case` it could not statically parse** (validated by a helper
   instead of `RequireFields`, computed dynamically, etc.) so those are
   human-reviewed, never silently passed.

Audit & fix: run the contract test → it enumerates current mismatches
(`add_run` is the first known) → correct each action's
`GetActionGuidance` REQUIRED line + EXAMPLE to agree with its
validator. No validator edits.

### B — Ordering & run steering

Add to `src/frmExport.frm` `PromptTemplate`:

- **MISTAKE 6 — Run model / partial formatting.** To format only PART
  of a shape's text (superscript/subscript/bold-word/color-word) that
  part must be its OWN run. Correct sequence: set the text first
  (`set_text`/`set_paragraph_text`), THEN `add_run` to append the new
  run (requires `paragraph_index` AND `run_index`, both 0-based), THEN
  apply `set_run_superscript`/`set_run_*` to that run's `run_index`.
  Include the exact "5" + "3"-superscript worked example.
- **ORDERING RECIPES block** — the crucial-order sequences, populated
  from the audit + known constraints: (a) all `add_paragraph` before
  any `set_bullet_style`/`set_indent_level`/`set_run_*`/`add_run` on the
  same shape (add_paragraph rebuilds `tr.Text`, destroying run/para
  formatting); (b) create-before-reference (add_shape with `ref_name`
  before `add_connector`; add_table before `set_cell_text`; add_slide
  before slide-scoped ops); (c) paragraph deletes high→low index;
  (d) longest-first `find_replace_text`; (e) the run/superscript recipe.

Add a one-line ordering note to the `GetActionGuidance` entries for
`add_run`, `set_run_superscript`, `set_run_subscript`, `add_paragraph`
so the in-app **Fix Errors** path also carries the sequencing hint.

### Sync

After edits: `python update_macros.py` (bake `src/` → carrier) and
`python tools/sync_actions_guidance.py` (regenerate the
`docs/ACTIONS_REFERENCE.md` AUTO-GUIDANCE appendix). The carrier must
still open and `GetAllActionTypes` return its full string (compile
sanity — large VBA string edits can wedge the project; see Risks).

## 4. Files

- CREATE `tests/test_guidance_contract.py` — the static drift gate.
- MODIFY `src/modExecuteInstructions.bas` — `GetActionGuidance`
  REQUIRED/EXAMPLE corrections for every drifted action + ordering
  one-liners on the four run/paragraph entries. **No change to
  `ValidateAction` or `DispatchAction` logic.**
- MODIFY `src/frmExport.frm` — `PromptTemplate`: MISTAKE 6 + ORDERING
  RECIPES blocks.
- REGEN `docs/ACTIONS_REFERENCE.md` via `tools/sync_actions_guidance.py`.
- `update_macros.py` (run, not modified).

## 5. Risks & handling

- **Zero execution-path change** → engine behavior is unchanged; only
  guidance strings + a new test. Lowest-risk class of change.
- **VBA compile wedge.** `GetActionGuidance` / `PromptTemplate` are
  large `s = s & "..." & vbCrLf` concatenations; a broken line
  continuation or a mid-module declaration compile-errors the project
  and modal-wedges COM. Mitigation: after every edit batch run
  `update_macros.py` then a carrier-open smoke asserting
  `len(GetAllActionTypes) > 1000`; keep all edits inside existing
  concatenation/`Select Case` structure; no new module-level
  declarations.
- **Contract-test false confidence.** If the parser silently skips
  unparseable cases it gives false green. Mitigation: the test prints
  and counts unparseable cases and fails if that set changes
  unexpectedly / is non-empty without an explicit allowlist.
- **Doc drift.** `tools/sync_actions_guidance.py` + the existing
  `test_guidance_doc_sync.py` keep `ACTIONS_REFERENCE.md` in lockstep.

## 6. Testing & success metric

Done when:
1. `python tests/test_guidance_contract.py` = 0 mismatches across all
   parseable dispatched actions, exit 0, and its unparseable-case list
   is empty or an explicitly reviewed allowlist.
2. Regression all green: `python tests/test_guidance_coverage.py`,
   `python tests/test_example_validity.py`,
   `python tests/test_guidance_doc_sync.py`,
   `python tests/run_smoke_guidance.py`, and a carrier-open compile
   smoke (`GetAllActionTypes` length > 1000) after `update_macros.py`.
3. `docs/ACTIONS_REFERENCE.md` regenerated and `test_guidance_doc_sync`
   green.

**Honest scope (not gated):** actual weak-LLM "right first try"
behavior is NOT in the deterministic gate (no LLM in CI). Manual
acceptance: re-run the exact failing scenario — "add text '3' into a
box that already says '5', then make the '3' superscript" — once
through the app/paste-JSON and confirm the LLM now emits a correctly
ordered batch (set text → `add_run` with `run_index` →
`set_run_superscript` on that `run_index`). Consistent with the
SP1/SP2 honest-scope precedent.

## 7. Out of scope

C (engine pre-Apply auto-sequencer/validator), any
`ValidateAction`/`DispatchAction`/action-execution change, the
single-source-of-truth refactor (Approach 3), desktop app changes,
SP5. No new actions (the needed actions already exist; the problem is
discoverability + ordering steering + guidance correctness).

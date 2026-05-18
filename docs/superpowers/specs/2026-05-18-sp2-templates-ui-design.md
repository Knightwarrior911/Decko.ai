# SP2 — Templates & Deck-DNA Visual Layer — Design Spec

Date: 2026-05-18
Status: Approved (design); pending implementation plan
Scope: SP2 only. Builds on SP1 (merged to `main`, app at HEAD ~`ccbeb64`+).

## 1. Background

SP1 shipped a Windows pywebview+Python desktop app wrapping the FROZEN
VBA/COM engine: BYO-key LLM, 100% local, persistent settings, per-session
chat history, selectable text, Save-PowerPoint / New-session / working
indicator. Today the engine's template + Deck-DNA + decks-as-code surface
is only reachable by typing a natural-language chat request that the LLM
must translate. SP2 exposes that surface as **direct visual controls** —
no LLM round-trip for deterministic ops.

Engine `src/` is FROZEN and reused via `app.Run("PPT_AI_Editor!…")`.

## 2. Locked decisions

| # | Decision | Consequence |
|---|----------|-------------|
| DC1 | Slide-over RIGHT panel, button-triggered ("Templates"); chat stays primary; panel wide + scrollable with collapsible sections | Chat layout untouched; panel holds the full feature set |
| DC2 | Hybrid content entry: Apply inserts the template with app-supplied **placeholder** content instantly (no LLM); optional **Fill-with-AI** (one-line brief → LLM drafts slot text → form populates → user reviews → Apply) | Deterministic apply is LLM-free; AI is opt-in |
| DC3 | Full surface in the panel: apply (builtin + captured), capture (active slide + name), manage (list/rename/delete), generate_variants, build_deck_from_spec, extract_spec | One cohesive visual layer |
| DC4 | Apply/variants target control: **append** OR **replace slide N** | Safe default append; explicit replace |
| DC5 | Approach A: new `Api` methods build the canonical actions JSON and run via the existing single-COM-thread `DeckController`/`ExecuteFromString`; **no LLM** for deterministic ops; each visual op recorded as a turn in the current session | Thinnest; reuses SP1; unified history; gate-testable |
| DC6 | Capture = the **active** PowerPoint slide + a name prompt | Matches the existing macro semantics |
| DC7 | Engine `src/` and the carrier remain UNCHANGED | SP2 wraps, never rewrites |

## 3. Architecture

Reuse the entire SP1 app. New code is additive:

- `app/template_slots.py` — pure data: `BUILTIN_SLOTS` (the 7 builtin
  templates → their literal required slots) + per-slot placeholder
  defaults + a `default_content(template)` helper. Authoritative slot
  map (from `modActionsTemplate.ValidateTemplateSlots`):
  - `title`: `title`, `subtitle`
  - `section`: `section_number`, `section_title`
  - `bullets`: `heading`, `bullets` (list[str])
  - `two_col`: `heading`, `left_body`, `right_body`
  - `comparison`: `heading`, `left_label`, `left_body`, `right_label`,
    `right_body`
  - `kpi_dashboard`: `heading`, `tiles` (list[{stat,label}])
  - `quote`: `quote_text`, `attribution`
- `app/deck_controller.py` — add small helpers used by the new Api
  methods: run a one-off action batch (reusing `_run`/ExecuteFromString),
  fetch captured-template list, fetch extract-spec JSON. (`get_prompt_
  template`, `get_snapshot`, `save_deck_now` already exist.) No new COM
  apartment concerns — everything still runs on the single `_com` worker.
- `app/main.py` `Api` — new methods (all routed through `self._com`,
  all require an active session/`self.orch`, each logs a turn via
  `self.store.add_turn(..., session_id=self.session_id)`):
  - `list_builtin_templates()` → `BUILTIN_SLOTS` metadata (pure, no COM)
  - `list_captured_templates()` → captured registry (names + slots)
  - `apply_template(template, content, target)` —
    `target = {"mode":"append"}` or `{"mode":"replace","slide":N}`
  - `fill_with_ai(template, brief)` → existing `LLMClient`, focused
    "produce JSON for exactly these slots" prompt → returns a content
    dict (does NOT apply; the panel populates the form for review)
  - `capture_template(name)` — capture active slide
  - `rename_template(from_name, to_name)` (engine params `from`/`to`)
  - `delete_template(name)`
  - `generate_variants(payload, target)` — `payload` = `{template,n}`
    or `{templates:[…]}` + `content`
  - `build_deck_from_spec(spec)` — spec = list of `{template,content}`;
    optional `clear_existing`
  - `extract_spec()` → returns the spec JSON string for the panel
- `app/web/{index.html,app.css,app.js}` — a right slide-over panel
  toggled by a "Templates" button; collapsible sections: Builtins,
  Deck-DNA, Variants, Decks-as-code. Reuses the existing
  working-indicator, selectable text, and Copy-button patterns.

## 4. Data flow

**Apply (deterministic, no LLM):** pick template → slot form
prefilled with placeholders → choose target (append / replace N) →
`api.apply_template(template, content, target)` → `_com` →
`DeckController` builds `{"actions":[{"type":"apply_template",
"template":…,"content":…(,"slide":N)}]}` → `ExecuteFromString` →
summary → `store.add_turn(session_id, request="Apply template:
<name> (<target>)", result_summary=summary)` → panel shows result,
sessions list refreshes.

**Fill-with-AI:** `api.fill_with_ai(template, brief)` → `LLMClient.call`
with a focused prompt naming the exact slots → returns a content dict
→ panel fills the slot form → user edits/reviews → normal Apply.

**Capture:** `api.capture_template(name)` → action
`{"type":"capture_template","name":…}` (active slide) →
ExecuteFromString → log turn → refresh Deck-DNA list.

**Manage:** rename → `{"type":"rename_template","from":…,"to":…}`;
delete → `{"type":"delete_template","name":…}`; list →
`list_templates` action / registry read.

**Variants / build-from-spec / extract:** corresponding single actions;
`extract_spec` returns JSON to the panel (shown + Copy button).

## 5. Error handling

Reuse SP1 mechanisms: `EmptyDeckError`, `NoPowerPointError`, COM
transient retry, `httpx.HTTPStatusError` surfacing (for Fill-with-AI).
Engine `FAILURES (N)` and `ValidateTemplateSlots` skip reasons are shown
verbatim in the panel result line. Panel controls are disabled with a
note until a session is started (same gate as chat). `extract_spec` on a
deck with no template-tagged slides returns a clear "no template slides"
message, not an error. No silent deck save — the SP1 Save-PowerPoint
button remains the only save path.

## 6. Testing

Same discipline as SP1 (deterministic COM harness; LLM stubbed/mocked;
UI manual).

- `tests/app/test_template_slots.py` — pure unit: `BUILTIN_SLOTS`
  completeness vs the authoritative map; `default_content()` produces
  every required slot for each of the 7 builtins (correct shapes:
  `bullets`→list, `tiles`→list of {stat,label}).
- `tests/app/test_llm_client.py` (extend) — `fill_with_ai`'s prompt
  shaping + JSON parse via mocked `httpx` transport (no network).
- `tests/run_smoke_app_templates.py` (new, COM, no LLM): on a built
  test deck via the `Api`/`_com` path — apply each of the 7 builtins
  (assert engine "applied", slide count grows / replaces); capture an
  active slide → `list_captured_templates` shows it → rename → delete →
  gone; `generate_variants` adds N slides; `build_deck_from_spec`
  builds the spec'd slides; `extract_spec` returns JSON containing the
  built template entries. Single isolated PowerPoint run, transient-COM
  retry, kill orphan POWERPNT — SP1 pattern.
- `tests/run_smoke_app.py` aggregator gains a `templates` gate. The
  existing unit suite + SP1 gates (`store_unit`, `llmclient_unit`,
  `core_loop`, `packaging_smoke`) stay green.
- UI slide-over panel: NOT in the deterministic gate — manual
  screenshot review (honest scope, per SP1 §8 precedent).

## 7. Success metric

SP2 is done when:
1. `tests/run_smoke_app.py` = 100% (now including the `templates`
   gate) + the unit suite green.
2. The full engine regression guard stays green (src frozen:
   `git diff origin/main -- src/` empty).
3. PyInstaller still builds a launchable `Decko.exe`
   (`packaging_smoke` green) — the panel ships in the exe.

UI visuals and real-LLM Fill-with-AI quality are out of the
deterministic gate (manual), consistent with project history.

## 8. Out of scope for SP2

SP3 licensing, SP4 cloud/accounts, SP5 branding/auto-update/polish.
No engine `src/` changes. No new LLM providers. No undo (unchanged).

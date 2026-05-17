# Decko Desktop App — Design Spec

Date: 2026-05-17
Status: Approved (design); pending implementation plan
Scope of this spec: **SP1 only** (Core local app MVP). SP2–SP5 listed for context, not designed here.

## 1. Background

Decko today is a VBA engine living in a `.pptm` carrier, driven from Python
via PowerPoint COM (win32com). The "brain" is an external LLM; the user
pastes a deck snapshot into the LLM, pastes the returned actions JSON back,
and the VBA "hands" (246-action engine + 32-check verify loop + builtin
templates + Deck DNA captured templates) execute it against the live deck.

Goal: ship Decko as a **distributable Windows desktop app** the owner AND
end users can install and run on their own computers with **no Python or
dev tooling**. The valuable IP is the action engine + verify loop +
prompt/templates — NOT the VBA delivery mechanism. The app wraps the
existing engine; it does not rewrite it.

## 2. Locked decisions (premises resolved during brainstorming)

| # | Decision | Consequence |
|---|----------|-------------|
| D1 | **All end users already have Microsoft PowerPoint desktop** | Keep the COM engine. Full fidelity. No python-pptx port. Far smaller build. |
| D2 | Platform is **Windows-only** | Forced by COM. Mac/iOS impossible on this engine; not in scope ever for this path. |
| D3 | **BYO key** for the LLM brain | App calls the LLM itself using each user's own API key, stored locally. No billing/usage infra for the owner. |
| D4 | **Local now, cloud later** | SP1 is 100% local: no server, no accounts. Local DB schema designed sync-ready so a future cloud layer adds without rework. |
| D5 | Deck interaction: **both modes** | Attach to the currently-open presentation AND open a picked `.pptx` file. |
| D6 | Providers: **Anthropic + OpenAI + generic OpenAI-compatible (base-URL)** | Generic field covers DeepSeek / MiniMax / local / proxies. |
| D7 | UI paradigm: **chat + side panel** | Conversational main thread; fixed side panel for deck target, settings, history. |
| D8 | Shell tech: **Approach B — pywebview + Python core, PyInstaller `.exe`** | Web-tech chat UI in a native window; Python backend reuses the entire COM engine unchanged; one `.exe`, no deps for users. |

## 3. Decomposition (platform → sub-projects)

The original request is a platform, not one spec. Each sub-project gets its
own spec → plan → build cycle.

| # | Sub-project | Status |
|---|-------------|--------|
| **SP1** | **Core local app MVP** — shell, deck picker (both modes), prompt → BYO-key LLM → run engine → verify/Fix display, settings (key/provider), local SQLite (history, sync-ready), one-click installer, no Python for users | **designed here** |
| SP2 | Templates / Deck DNA visual UI (browse / apply / capture / manage) | later |
| SP3 | Licensing / activation (if selling) | later |
| SP4 | Cloud accounts + cross-device sync (the "later" in D4) | later |
| SP5 | Branding, auto-update, polish | later |

## 4. SP1 Architecture

One Windows `.exe` produced by PyInstaller. Bundles: Python runtime,
pywebview, app code, the **prebuilt** `PPT_AI_Editor.pptm` carrier, provider
HTTP clients.

Layers:

- **UI** — HTML/CSS/JS rendered in a pywebview native window. Chat thread +
  side panel.
- **Bridge** — pywebview `js_api`: JS calls Python methods, Python returns
  results/events.
- **Core (Python; reuses the existing toolchain):**
  - `DeckController` — PowerPoint COM. Attach mode (Dispatch →
    `ActivePresentation`) or file mode (DispatchEx → open picked `.pptx`,
    operate, `Save`). Exposes `get_snapshot()`, `run_actions(json)`,
    `run_verify()`, `fix_prompt(kind)`.
  - `LLMClient` — provider abstraction (Anthropic native / OpenAI / generic
    OpenAI-compatible base-URL). Builds prompt = carrier `PromptTemplate` +
    snapshot + user message; returns sanitized actions JSON.
  - Engine calls — invoke the carrier's existing
    `BuildSnapshotJson` / `ExecuteFromString` / verify entry points
    **unchanged**.
  - `Store` — SQLite at `%APPDATA%\Decko\decko.db`. Sync-ready schema:
    UUID primary keys, `updated_at`, soft-delete columns.
  - `Secrets` — API keys in Windows Credential Manager via `keyring`
    (never in SQLite or plaintext).
- Carrier ships inside the app; first run copies it to
  `%APPDATA%\Decko\engine\PPT_AI_Editor.pptm` (writable). `update_macros`
  is **not** run at runtime — the carrier is prebuilt at package time.

## 5. SP1 Components

- **App window** (single pywebview window):
  - Left side panel: deck target selector `[Attach to open ▾ / Open file…]`,
    provider + API-key settings, history list.
  - Main region: chat thread (user requests, app replies with action
    summary + verify chips + Fix affordances).
- **DeckController** — see §4. Resilient to transient COM/RPC errors by
  reusing the existing harness retry pattern (DispatchEx retry, kill orphan
  POWERPNT, retry com_error/AttributeError only, never an assertion).
- **LLMClient** — provider-specific request shaping + response parsing for
  the 3 provider modes; routes raw output through the existing sanitizer.
- **ChatOrchestrator** — one turn: user text → `get_snapshot()` →
  `LLMClient.call()` → actions → `ExecuteFromString` → `run_verify()` →
  render result → persist turn to `Store`.
- **Settings** — provider ∈ {anthropic, openai, generic}, `base_url`
  (generic only), model name, API key (keyring). Validated on save.
- **Installer** — Inno Setup wrapping the PyInstaller output →
  `Decko-Setup.exe`. Double-click install, Start-menu shortcut. Detects
  PowerPoint presence at launch.

## 6. SP1 Data flow (one turn)

```
User types in chat
  → JS bridge → ChatOrchestrator.run(text)
  → DeckController.get_snapshot()            [COM]
  → LLMClient.call(snapshot, text, key)      [user's provider]
  → actions JSON  (sanitized)
  → DeckController.run_actions(json)         [carrier ExecuteFromString]
  → DeckController.run_verify()              [carrier verify loop]
  → result {applied, skipped, failures[], warnings[]}
  → Store.insert_turn(...)                   [SQLite]
  → chat bubble + verify chips + [Fix errors] / [Fix this]
       (Fix buttons build the existing repair prompt → next turn)
```

## 7. SP1 Error handling

- **No PowerPoint installed** — blocking friendly dialog at startup with an
  install link; exit.
- **No open deck (attach mode)** — inline chat error; offer "Open file"
  instead.
- **COM transient / RPC failure** — reuse harness retry; surface as
  "PowerPoint busy, retried".
- **LLM errors** — bad key → settings prompt; timeout / rate-limit →
  inline retry; malformed JSON → existing sanitizer → if still invalid,
  route through the existing Fix-Errors failure-contract path.
- **Engine partial failure** — surface the existing `FAILURES (N)`
  per-action contract (exact index, type, reason) in the chat reply.
- **No silent save in attach mode** — the user owns Ctrl+S; Apply has no
  undo (same as today). This is stated explicitly in the UI, not hidden.

## 8. SP1 Testing

Reuse the project's deterministic COM-harness discipline. LLM is **stubbed**
in gated tests (no network), exactly like existing `run_smoke_*`.

- `core_loop` — snapshot → stub-LLM (fixed actions) → `ExecuteFromString`
  → verify; asserts applied/skipped/failure counts.
- `llmclient_unit` — mocked HTTP; asserts request shaping + response
  parsing for all 3 provider modes. No network.
- `store_unit` — SQLite schema + sync-ready columns (UUID PK,
  `updated_at`, soft-delete) round-trip.
- `packaging_smoke` — PyInstaller build succeeds; the bundled carrier
  loads via COM from the packaged location.
- **UI layer is NOT deterministically gated** — verified by manual
  screenshot review. Honest scope (per the themed-forms lesson: a VBA/UI
  visual cannot be deterministically asserted).
- All prior Decko harnesses remain green (engine untouched).

## 9. Success metric (for the eventual autoresearch goal)

SP1 is done when:
1. The SP1 deterministic harness suite (`core_loop`, `llmclient_unit`,
   `store_unit`, `packaging_smoke`) = 100%, exit 0.
2. The full existing Decko engine regression gate stays green.
3. PyInstaller produces a launchable `Decko-Setup.exe` whose installed app
   runs the stubbed core loop end-to-end on a machine with PowerPoint and
   **no Python installed**.

Visual/UX polish and real-LLM behavior are explicitly out of the
deterministic gate (manual verification), consistent with project history.

## 10. Out of scope for SP1

python-pptx port (D1), Mac/iOS (D2), owner-side billing/usage (D3), cloud
accounts/sync (D4, → SP4), licensing (→ SP3), templates visual UI (→ SP2),
branding/auto-update (→ SP5).

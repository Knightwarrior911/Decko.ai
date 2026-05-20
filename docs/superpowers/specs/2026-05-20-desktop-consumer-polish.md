# Decko Desktop — Consumer Polish (SP5)

**Status:** APPROVED — 2026-05-20
**Owner:** Claude (autonomous execution under /goal lock)
**Branch:** `feat/desktop-consumer-polish` off `main` (HEAD `fba2dea`)

## 1. Problem

`python -m app.main` opens a window that looks like an engineer test fixture:

- Sidebar packed with raw provider/model/base-URL inputs and the API key field.
- Internal jargon labels: "Deck DNA", "Decks-as-code", "Attach to open deck",
  "Build deck from spec", "Extract spec from deck", "base URL (generic only)",
  "Generate variants (append)".
- No onboarding, no empty state, no example prompts.
- One dense dark theme, 14px Segoe UI, near-mono palette (#0d0d0f / #161618 /
  #1d1d20 / #222 / faint accent #3a6ea5).
- Error text is raw Python: `f"{type(e).__name__}: {e}"`.
- Templates panel hidden behind a small `Templates ▸` ghost button.

Non-technical PowerPoint users will not get past first launch.

## 2. Goal

Polish UI/UX only — engine and backend `Api` surface stay intact — so a
first-time non-technical PowerPoint user can:

1. Open `Decko.exe`,
2. Walk a 3-step setup wizard,
3. Connect to an open deck or open a `.pptx`,
4. Send a chat message ("rewrite slide 3 for X") that mutates the deck,
5. Save.

…inside 5 minutes, with zero docs and zero engineer jargon visible.

## 3. Non-goals

- No engine changes (`src/*.bas` frozen except shipped `Do_extract_slides`).
- No new actions.
- No new `Api` methods (renames may force tweaks; mirror precisely).
- No licensing / telemetry / cloud sync (SP3/SP4 skipped).
- No installer signing or auto-update (separate SP5 sub-task).
- No marketing site.
- No replacing pywebview, Python backend, or COM bridge.

## 4. Scope — exact UI changes

### 4.1 Jargon rename (every user-facing string)

| Old (current `index.html` / `app.js`)          | New                                               |
| ---------------------------------------------- | ------------------------------------------------- |
| `Attach to open deck`                          | `Connect to open deck`                            |
| `Open a .pptx file`                            | `Open a PowerPoint file…`                         |
| `Start session`                                | `Start`                                           |
| `Save PowerPoint`                              | `Save deck`                                       |
| `Templates ▸`                                  | `Templates`                                       |
| `Deck DNA (captured)`                          | `My saved layouts`                                |
| `Decks-as-code`                                | (collapse under) `Power user tools`               |
| `Build deck from spec`                         | `Build from JSON`                                 |
| `Extract spec from deck`                       | `Export deck as JSON`                             |
| `Generate variants (append)`                   | `Generate variants`                               |
| `base URL (generic only)`                      | (hide behind Advanced expander) `Custom base URL` |
| `model (e.g. claude-opus-4-7 / gpt-4o / ...)`  | (replace with curated dropdown — see 4.4)         |
| `Save settings`                                | `Save`                                            |
| `Apply` / `Del` (captured list buttons)        | `Apply` / `Delete`                                |
| `Replace whole deck`                           | `Replace the whole deck`                          |
| `Optional: one-line brief for Fill-with-AI`    | `One-line brief (optional)`                       |
| Composer placeholder `Describe the change…`    | `What would you like to change? Try: "rewrite slide 3 for Acme"` |

Jargon **blocklist** (must not appear anywhere in `index.html` or `app.js`
after polish — verified by `tests/run_smoke_ui_polish.py`):

- `Deck DNA`
- `Decks-as-code`
- `Build deck from spec`
- `Extract spec from deck`
- `base URL (generic only)`
- `Attach to open deck`

### 4.2 First-run wizard

Triggered when `boot()` returns `has_key === false` AND `localStorage
.getItem("decko_wizard_done") !== "1"`. Three screens (modal over chat area):

1. **Welcome** — `"Welcome to Decko"` headline, one-line "AI edits for your
   PowerPoint decks." subtitle, `Get started` button.
2. **Choose your AI** — provider radio (Anthropic / OpenAI / Other), model
   dropdown (curated per provider, see 4.4), API key input (password), inline
   note "We store your key locally only — never sent to Decko servers." `Save
   and continue` button.
3. **Pick a deck** — two large buttons: `Connect to open deck` (primary) and
   `Open a PowerPoint file…` (secondary). On click, runs `open_session`. On
   success, hides wizard and sets `decko_wizard_done = "1"`.

Wizard CSS: full-area overlay (`position:fixed; inset:0; background:rgba(13,
13,15,.94); z-index:100;`), card centered, max-width 520px.

### 4.3 Settings dialog (gear icon)

Replace the always-visible LLM section in the sidebar with a gear icon button
in the sidebar header. Click → modal dialog with the same fields as the
wizard's "Choose your AI" step plus:

- API key display: `••••••••• Connected ✓` when `has_key === true`, with an
  "Update key" link that reveals the input.
- Light/dark theme toggle (4.10).
- `Save` (primary) and `Cancel` (ghost) at the bottom.

The settings dialog calls the same `api.save_settings(...)` backend method
unchanged.

### 4.4 Curated model dropdown per provider

```
Anthropic
  claude-opus-4-7        (recommended)
  claude-sonnet-4-6
  claude-haiku-4-5

OpenAI
  gpt-4o                 (recommended)
  gpt-4o-mini

Other (advanced)
  free-text model field + Custom base URL field
```

Implemented as `<select id="modelPick">` populated from JS constant
`MODELS_BY_PROVIDER`. When provider === `generic`, hide the picker and reveal
the legacy free-text `#model` input + `#baseUrl` input under an `<details>`
expander labeled `Advanced`.

### 4.5 Empty/idle hero state

When no session is open (`currentSession === null` and threadTitle empty),
the main chat pane shows:

```
What deck do you want to edit today?

[ Connect to open deck ]        ← primary, large
[ Open a PowerPoint file… ]     ← secondary, ghost
```

Replace the small dropdown+`Start session` button currently in the sidebar's
Deck section. The sidebar Deck section reduces to a "Current deck: <name>"
status line + a `Switch deck` ghost link.

### 4.6 Composer example chips

Above the textarea, 4 chips (clickable). On click, prefill the textarea and
focus it.

1. `Rewrite slide 3 for a different company`
2. `Make this chart a stacked bar`
3. `Apply our brand colors`
4. `Make slide 1 more visual`

Hidden while a `working…` bubble is showing.

### 4.7 Templates panel rename + discoverability

- Section headers:
  - `Built-in templates` → unchanged.
  - `Deck DNA (captured)` → `My saved layouts`.
  - `Variants` → unchanged.
  - `Decks-as-code` → `Power user tools` (collapsed by default).
- `Build from JSON` and `Export deck as JSON` live inside `Power user tools`.
- Templates button in sidebar: rename `Templates ▸` → `Templates`. Add a
  one-time coach-mark tooltip "Browse layouts here →" on first session open
  (driven by `localStorage.decko_tpl_seen`).

### 4.8 Tooltips

Every form field in the Settings dialog, wizard, and Templates panel gets a
`title=""` attribute with a plain-English one-liner. Examples:

| Field                  | Tooltip                                                     |
| ---------------------- | ----------------------------------------------------------- |
| Provider               | Which AI service Decko sends your edit requests to.         |
| Model                  | Which model to use. Recommended defaults are pre-selected.  |
| API key                | Your own key for the chosen AI service. Stored on this PC.  |
| Custom base URL        | Only needed if your AI is hosted at a non-standard URL.     |
| Replace the whole deck | Wipes existing slides before building from JSON.            |
| Capture active slide   | Save the current slide as a reusable layout.                |

Visual tooltip: standard browser `title=""` (sufficient; no custom popover).

### 4.9 Visual polish (CSS)

- Base font: 15px (was 14px implicit), system stack unchanged.
- App header bar: dark bar across the top of the sidebar containing
  `Decko` wordmark (700 weight, accent color #3a6ea5) and gear icon.
- App icon: `app/web/decko_icon.svg` (a stylized "D" inside a slide-shaped
  rectangle, accent fill). Referenced from `index.html` as `<link rel="icon">`
  and shown 24px in the header.
- Spacing: bump padding on bubbles to `.7rem 1rem`, gap between sidebar
  sections to `1.4rem`.
- Buttons: primary uses accent `#3a6ea5`/hover `#4a82be`; ghost stays dark.
- Loading state: replace italic `"working… (PowerPoint + LLM)"` with named
  bubble `"Editing your deck…"` and a CSS pulse animation.
- Toasts: a `#toast` element fixed bottom-right; `showToast(text, kind)`
  appends a 3-second toast (`success` = accent green `#2e8b57`, `info` = blue,
  `error` = `#a14040`).

### 4.10 Title bar + dirty indicator

- Window title (`webview.create_window("Decko", …)` argument): updated via
  `webview.windows[0].title` on session open to `Decko — <deck name>` and to
  `Decko — <deck name> ●` when dirty.
- Dirty tracking: any successful `send()` that returns a non-error summary
  marks the deck dirty (`window.deckDirty = true`); `save_powerpoint()`
  success clears it.
- `#savePptBtn` (now "Save deck") gets `.dirty` class when dirty → accent
  background.

Title-update mechanism: a new tiny `Api.set_window_title(title)` method that
sets `webview.windows[0].title`. (This is a narrow exception to the "no new
Api methods" rule — purely cosmetic, doesn't touch the engine.)

### 4.11 Plain-English error translation

Wrap every `bubble("fail", r.error)` callsite in a `friendlyError(raw)` JS
helper:

| Raw substring contains                | Friendly                                                              |
| -------------------------------------- | --------------------------------------------------------------------- |
| `NoOpenDeckError` / `No deck open`     | No deck is open in PowerPoint. Open a deck, or click *Open file*.     |
| `NoPowerPointError` / `not found`      | Microsoft PowerPoint isn't installed. Install it and try again.       |
| `EmptyDeckError` / `Empty deck`        | The deck has no slides. Add at least one and try again.               |
| `LLM API error 401`                    | Your AI key was rejected. Open Settings and check it.                 |
| `LLM API error 429`                    | The AI is rate-limited. Wait a minute and try again.                  |
| `Spec is not valid JSON`               | That isn't valid JSON. Check the brackets and quotes.                 |
| (default)                              | Something went wrong: `<raw, truncated to 200 chars>`.                |

### 4.12 Light theme

Toggle in Settings dialog. Stored in `localStorage.decko_theme`. CSS variables
on `<html data-theme="light|dark">`:

| Var               | Dark default   | Light          |
| ----------------- | -------------- | -------------- |
| `--bg`            | `#0d0d0f`      | `#fafafa`      |
| `--side-bg`       | `#161618`      | `#f0f0f2`      |
| `--surface`       | `#1d1d20`      | `#ffffff`      |
| `--text`          | `#eee`         | `#1a1a1c`      |
| `--muted`         | `#888`         | `#666`         |
| `--border`        | `#333`         | `#d0d0d4`      |
| `--accent`        | `#3a6ea5`      | `#2c5d8f`      |
| `--user-bubble`   | `#22344a`      | `#dde6f2`      |
| `--app-bubble`    | `#1d1d20`      | `#f3f3f5`      |

## 5. Files touched

- `app/web/index.html` — rewrite (full).
- `app/web/app.css` — rewrite (full, ~150 lines with vars + light theme).
- `app/web/app.js` — rewrite (full, ~450 lines with wizard, settings,
  hero, chips, tooltips, toasts, dirty tracking, friendly errors).
- `app/web/decko_icon.svg` — new.
- `app/main.py` — add `Api.set_window_title(title)` only.
- `tests/run_smoke_ui_polish.py` — new.
- `tests/run_smoke_app.py` — register new gate.
- `docs/screenshots/consumer-polish/` — new directory with 5 before/after PNG pairs.

## 6. Metric

Binary gates (deterministic):

- `python tests/run_smoke_ui_polish.py` → exit 0 (jargon blocklist absent;
  selectors for wizard, settings dialog, hero, chips, tooltips exist).
- `python tests/run_smoke_app.py` → all gates PASS (no regression).
- `python tests/run_smoke.py` → engine guard PASS.
- `python tests/test_guidance_coverage.py` → PASS.
- `python tests/test_guidance_doc_sync.py` → PASS.

Manual (committed artifact):

- 5 before/after PNG pairs in `docs/screenshots/consumer-polish/`:
  1. First launch (current dense sidebar → wizard step 1).
  2. After API key save (sidebar with raw inputs → wizard step 2).
  3. Empty/idle (small dropdown → hero with 2 big buttons).
  4. Composer (placeholder only → chips above).
  5. Templates panel (Deck DNA / Decks-as-code labels → My saved layouts / Power user tools).

## 7. Risks

- pywebview title mutation through `Api` is unusual; verify the
  `window.title = …` path works on the installed pywebview version before
  baking it in. Fallback: skip title-bar dirty indicator, keep only the
  `Save deck` button highlight.
- `webview.windows[0]` may not exist before the window is loaded; gate the
  call with `if webview.windows:`.
- Light theme: ensure chart bubbles + warnings remain legible (test both
  themes manually).
- Wizard local-storage flag: must be cleared if user clicks `Update key` and
  the saved key is rejected — handled by re-running wizard if `boot()` reports
  no key.

## 8. Out of scope (firm)

- Engine, new actions, licensing, telemetry, cloud sync, Mac, installer
  signing, marketing site, replacing pywebview.

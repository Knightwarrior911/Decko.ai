# Consumer-Polish before/after walkthrough

5 paired screenshots covering the documented gaps in
`docs/superpowers/specs/2026-05-20-desktop-consumer-polish.md`:

| # | Scene                | Before (origin/main)                                  | After (this branch)                                        |
| - | -------------------- | ----------------------------------------------------- | ---------------------------------------------------------- |
| 1 | First launch         | Dense sidebar, naked API-key field, no onboarding     | Branded header + first-run wizard ("Welcome to Decko")     |
| 2 | Idle / empty state   | Tiny dropdown + "Start session" button in sidebar     | Centered hero ("What deck do you want to edit today?") + 2 big CTAs |
| 3 | Composer             | Single textarea with placeholder "Describe the change…" | 4 example chips above the textarea + Ctrl+Enter hint + plain "Editing your deck…" loading state |
| 4 | Templates panel      | "Deck DNA (captured)" / "Decks-as-code" jargon        | "My saved layouts" / "Power user tools" (collapsed)         |
| 5 | Settings / LLM area  | Provider/model/baseURL/key inline in sidebar          | Settings dialog (gear icon) with curated model dropdown, masked-key status, theme toggle |

Each before/after pair shares the same window size and rough composition so
the diff is visual not contextual.

## Regenerating

```
python tools/capture_app_screenshots.py before   # against origin/main HTML
python tools/capture_app_screenshots.py after    # against polished HTML
```

`capture_app_screenshots.py` boots pywebview with a STUB `Api` (no
PowerPoint, no LLM) and drives the page through each state via
`evaluate_js`. To capture `before` with the legacy HTML, temporarily
restore `app/web/` from `origin/main` (or any prior commit), run the
capture, then restore your working tree.

# Follow-up: ResolveSlideScope crashes on a 0-slide deck

Date: 2026-05-17
Status: DEFERRED (engine fix not done in SP1; app-layer guard shipped instead)

## Bug
`src/modExportSnapshot.bas` `ResolveSlideScope(pres, scope)` executes
`ReDim allArr(1 To total)` where `total = pres.Slides.Count`. When the
deck has 0 slides, `ReDim allArr(1 To 0)` raises VBA Run-time error 9
(Subscript out of range), producing a modal dialog that wedges
PowerPoint COM. Any caller (snapshot on an empty deck) hits this.

## SP1 mitigation (shipped)
`app/deck_controller.py` raises `EmptyDeckError` before calling the
engine when `deck.Slides.Count == 0`, so the desktop app never reaches
the crashing path. Engine `src/` left frozen per the SP1 locked premise.

## Proposed engine fix (deferred — own spec/plan/gate later)
In `ResolveSlideScope`, before the `ReDim`, handle `total = 0`:
return an empty 0-length array (no slides in scope) instead of
`ReDim allArr(1 To 0)`. Add a deterministic harness asserting
`BuildSnapshotJson` on a 0-slide presentation returns valid JSON with
an empty `slides` array and does not raise. Then run the full engine
regression gate.

## Why deferred
SP1's locked premise is "engine reused unchanged". Touching the engine
requires update_macros + the full engine regression gate, which is a
separate unit of work. User decision (2026-05-17): guard now, fix
engine later.

# PPT AI Editor — Phase 2 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand action surface from 15 → 70 by adding granular text, layout, cross-cutting, speaker notes, images, slide structure, table, group, connector, and native chart actions, plus a standalone `frmImportSlides` UserForm and a snapshot v2 with per-paragraph and chart fidelity.

**Architecture:** New per-bucket modules under `src/` (`modActionsText`, `modActionsLayout`, `modActionsImage`, `modActionsSlide`, `modActionsTable`, `modActionsGroup`, `modActionsConnector`, `modActionsChart`). Snapshot exporter extended in place. `modExecuteInstructions` extended with validation + dispatch cases for the 55 new types. UserForm grows; PromptTemplate function shows all 70 schemas.

**Tech Stack:** PowerPoint VBA, VBA-JSON, Python 3 with pywin32 + python-pptx, Office 2016+ (32-bit or 64-bit).

**Spec:** `docs/specs/2026-05-08-ppt-ai-editor-phase2-design.md`.

---

## File structure (Phase 2 final state)

```
Documents/PPT_AI_Editor/
├── PPT_AI_Editor.pptm
├── src/
│   ├── modJSON.bas                         (unchanged)
│   ├── modBackup.bas                       (unchanged)
│   ├── modExportSnapshot.bas               (extended — paragraphs, notes, gaps, chart, group, table_extra)
│   ├── modActions.bas                      (V1 unchanged; +set_speaker_notes, +append_speaker_notes appended)
│   ├── modActionsText.bas        NEW
│   ├── modActionsLayout.bas      NEW       (Layout + cross-cutting batch)
│   ├── modActionsTable.bas       NEW
│   ├── modActionsChart.bas       NEW
│   ├── modActionsImage.bas       NEW
│   ├── modActionsSlide.bas       NEW
│   ├── modActionsGroup.bas       NEW
│   ├── modActionsConnector.bas   NEW
│   ├── modExecuteInstructions.bas          (extended — validation + dispatch for ~70 actions)
│   ├── modUI.bas                           (extended — ImportSlides entry point)
│   ├── frmExport.frm/.frx                  (PromptTemplate fn extended to all 70 schemas)
│   ├── frmExecute.frm/.frx                 (unchanged)
│   └── frmImportSlides.frm/.frx  NEW
├── tools/
│   ├── build_carrier.py                    (unchanged)
│   ├── build_forms.py                      (extended — adds frmImportSlides; updates PromptTemplate)
│   └── precheck_carrier.py       NEW
├── tests/
│   ├── make_test_decks.py                  (extended — adds phase2.pptx fixture)
│   └── run_smoke.py                        (extended — ~30 new test functions)
├── test_decks/
│   ├── smoke_3slide.pptx
│   ├── full_visual.pptx
│   └── phase2.pptx               NEW
└── docs/specs/2026-05-08-ppt-ai-editor-phase2-design.md
```

---

## Section 0: Standard task workflow

Every task in Phase 2 follows this micro-pattern. Each task body shows only the deltas (the failing test, the VBA code, the validation/dispatch cases, the expected output). The shared workflow steps below are referenced by number from each task — do not re-document them per task.

### 0.1 Pre-flight

Before touching code:

```bash
# Kill any user-held PowerPoint instance to avoid hidden popups
powershell -Command "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"

# Make sure existing smoke is green
python tests/run_smoke.py
```

### 0.2 Add the failing test

Edit `tests/run_smoke.py`. Append the new test function and add a call from `main()`. Use qualified macro names (`PPT_AI_Editor!<MacroName>`).

### 0.3 Run smoke; expect failure

```bash
python tests/run_smoke.py
```

Expected: existing tests pass, new test fails for the *right* reason (the macro doesn't exist yet, or a property is missing in the snapshot).

### 0.4 Edit VBA source

Edit the appropriate `.bas` file in `src/`. Always check for VBA reserved-word collisions in identifiers: `Dir`, `Date`, `Time`, `Name`, `Error`, `String`, `Type`. Never use `Application.PathSeparator` (PowerPoint VBA does not have it — hardcode `"\"`).

For long string literals (more than ~20 line continuations), build the string with `s = s & ...` lines inside a `Private Function` instead of a `Const ... & _` chain.

### 0.5 Sync + carrier compile precheck + smoke

```bash
powershell -Command "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"
python update_macros.py
python tools/precheck_carrier.py
python tests/run_smoke.py
```

`precheck_carrier.py` (introduced in Task 0) opens the carrier headlessly, calls `BuildSnapshotJson` on a known-good deck, and exits non-zero if the carrier doesn't compile. If precheck fails, read the .bas you just edited, fix the compile error, re-sync. **Do not run smoke until precheck passes.**

If precheck reports "Sub or function not defined: BuildSnapshotJson" right after a sync, that is the **compile-error symptom** — your new module didn't compile and broke the project. Treat as immediate fix, not as a spurious COM glitch.

### 0.6 Commit

```bash
git add <changed files>
git commit -m "<feat|test|fix>: <one-line summary>"
```

Carrier `PPT_AI_Editor.pptm` is already tracked from V1 — committing it as part of every task is fine since it changes any time you sync.

---

## Section 1: precheck and test deck (foundation tasks)

### Task 0: `tools/precheck_carrier.py` — guardrail script

**Files:**
- Create: `C:\Users\vinit\Documents\PPT_AI_Editor\tools\precheck_carrier.py`

- [ ] **Step 1: Write the script**

```python
"""Verify the carrier's VBProject compiles after a sync.

Run: python tools/precheck_carrier.py

Exits 0 if `BuildSnapshotJson` is callable on a fresh test deck; exits
non-zero with a diagnostic message if the carrier has a compile error
or the macro is missing.

Usage in agent workflow: always run this AFTER `python update_macros.py`
and BEFORE `python tests/run_smoke.py`. If it fails, read the .bas file
you most recently edited and fix the compile error. Do not proceed to
smoke.
"""
import os
import shutil
import sys
import tempfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
CARRIER = REPO_ROOT / "PPT_AI_Editor.pptm"
SAMPLE_DECK = REPO_ROOT / "test_decks" / "smoke_3slide.pptx"


def main() -> int:
    if not CARRIER.exists():
        print(f"ERROR: carrier not found at {CARRIER}")
        return 2
    if not SAMPLE_DECK.exists():
        print(f"ERROR: sample deck not found at {SAMPLE_DECK}")
        return 2

    try:
        import win32com.client
    except ImportError:
        print("ERROR: pywin32 not installed")
        return 2

    tmpdir = Path(tempfile.mkdtemp(prefix="pptai_precheck_"))
    deck_copy = tmpdir / SAMPLE_DECK.name
    shutil.copy2(SAMPLE_DECK, deck_copy)

    app = win32com.client.DispatchEx("PowerPoint.Application")
    deck = None
    carrier = None
    try:
        carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
        deck = app.Presentations.Open(str(deck_copy), WithWindow=True)
        deck.Windows(1).Activate()

        try:
            json_text = app.Run("PPT_AI_Editor!BuildSnapshotJson")
        except Exception as e:
            print(f"FAIL: BuildSnapshotJson call raised: {e}")
            print("  This usually means the carrier has a compile error.")
            print("  Read the .bas you most recently edited and fix the error.")
            return 1

        if not isinstance(json_text, str) or not json_text.strip().startswith("{"):
            print(f"FAIL: BuildSnapshotJson returned non-JSON: {json_text!r}")
            return 1

        print("OK: carrier compiles and BuildSnapshotJson returns valid JSON")
        return 0
    finally:
        for p in (deck, carrier):
            if p is not None:
                try:
                    p.Saved = True
                except Exception:
                    pass
                try:
                    p.Close()
                except Exception:
                    pass
        try:
            app.Quit()
        except Exception:
            pass
        shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run it on the current carrier; expect OK**

```bash
python tools/precheck_carrier.py
```

Expected stdout: `OK: carrier compiles and BuildSnapshotJson returns valid JSON`. Exit code 0.

- [ ] **Step 3: Commit**

```bash
git add tools/precheck_carrier.py
git commit -m "feat: precheck_carrier.py — verify VBProject compiles after sync"
```

---

### Task 1: `phase2.pptx` test fixture

Generate a deck with a native chart, grouped shapes, merged table cells, multi-paragraph bullet body, speaker notes, and a plain rectangle. The smoke tests for snapshot v2 and most new actions will use this fixture.

**Files:**
- Modify: `tests/make_test_decks.py`
- Generated: `test_decks/phase2.pptx`

- [ ] **Step 1: Add `make_phase2(path)` to `tests/make_test_decks.py`**

Insert after `make_full_visual` and before `def main()`:

```python
def make_phase2(path: Path) -> None:
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN

    pres = Presentation()
    layout_blank = pres.slide_layouts[6]

    # --- Slide 1: multi-paragraph bullet body + speaker notes
    s1 = pres.slides.add_slide(layout_blank)

    title = s1.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(0.8))
    title.text_frame.text = "Bullet Body Slide"
    title.text_frame.paragraphs[0].runs[0].font.size = Pt(28)
    title.text_frame.paragraphs[0].runs[0].font.bold = True

    body = s1.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(9), Inches(4))
    tf = body.text_frame
    tf.text = "First point about revenue"
    tf.paragraphs[0].runs[0].font.size = Pt(20)

    p2 = tf.add_paragraph()
    p2.text = "Second point about margins"
    p2.runs[0].font.size = Pt(20)

    p3 = tf.add_paragraph()
    p3.text = "Third point about headcount"
    p3.runs[0].font.size = Pt(20)

    s1.notes_slide.notes_text_frame.text = "Highlight Q3 outperformance and Y/Y growth"

    # --- Slide 2: native chart
    s2 = pres.slides.add_slide(layout_blank)
    chart_data = CategoryChartData()
    chart_data.categories = ["Q1", "Q2", "Q3", "Q4"]
    chart_data.add_series("FY24", (100, 110, 120, 130))
    s2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1), Inches(1.5), Inches(8), Inches(4.5),
        chart_data,
    )

    # --- Slide 3: grouped shapes (a 3-shape mini diagram)
    s3 = pres.slides.add_slide(layout_blank)
    box1 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1.5), Inches(2), Inches(1))
    box1.text_frame.text = "A"
    box2 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4), Inches(1.5), Inches(2), Inches(1))
    box2.text_frame.text = "B"
    box3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7), Inches(1.5), Inches(2), Inches(1))
    box3.text_frame.text = "C"
    # Grouping needs the shape collection; python-pptx supports
    s3.shapes._spTree  # ensure tree is materialized
    # We do not group here — we want phase2 actions Do_group_shapes
    # to do that as a real test.

    # --- Slide 4: table with future-merged cells + plain rect
    s4 = pres.slides.add_slide(layout_blank)
    rows, cols = 4, 3
    tbl_shape = s4.shapes.add_table(rows, cols, Inches(0.5), Inches(0.5), Inches(9), Inches(3))
    tbl = tbl_shape.table
    headers = ["Metric", "FY24", "FY25"]
    for c, h in enumerate(headers):
        tbl.cell(0, c).text = h
    body_data = [
        ["Revenue", "100", "112"],
        ["Margin", "26%", "28%"],
        ["Headcount", "1,200", "1,250"],
    ]
    for r, row in enumerate(body_data, start=1):
        for c, v in enumerate(row):
            tbl.cell(r, c).text = v

    rect = s4.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4.5), Inches(2), Inches(0.8)
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    rect.text_frame.text = "Plain"

    pres.save(str(path))
```

Update `main()`:

```python
def main() -> int:
    DECKS_DIR.mkdir(parents=True, exist_ok=True)
    make_smoke_3slide(DECKS_DIR / "smoke_3slide.pptx")
    make_full_visual(DECKS_DIR / "full_visual.pptx")
    make_phase2(DECKS_DIR / "phase2.pptx")
    print(f"[done] wrote 3 decks to {DECKS_DIR}")
    return 0
```

- [ ] **Step 2: Run it**

```bash
python tests/make_test_decks.py
```

Expected: `[done] wrote 3 decks to .../test_decks` and `phase2.pptx` exists.

- [ ] **Step 3: COM-verify the fixture**

```bash
python -c "
import win32com.client
app = win32com.client.DispatchEx('PowerPoint.Application')
try:
    pres = app.Presentations.Open(r'C:\\Users\\vinit\\Documents\\PPT_AI_Editor\\test_decks\\phase2.pptx', WithWindow=False)
    try:
        assert pres.Slides.Count == 4, f'slides: {pres.Slides.Count}'
        # Slide 1 should have body with 3 paragraphs
        body = pres.Slides(1).Shapes(2)
        paras = body.TextFrame.TextRange.Paragraphs().Count
        assert paras == 3, f'paragraphs: {paras}'
        # Slide 1 speaker notes
        notes = pres.Slides(1).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        assert 'Q3' in notes, f'notes: {notes!r}'
        # Slide 2 has a native chart
        assert pres.Slides(2).Shapes(1).HasChart, 'no chart on slide 2'
        # Slide 4 has a table and a rectangle
        s4 = pres.Slides(4)
        assert any(sh.HasTable for sh in s4.Shapes), 'no table on slide 4'
        print('OK: phase2.pptx fixture verified')
    finally:
        pres.Close()
finally:
    app.Quit()
"
```

Expected: `OK: phase2.pptx fixture verified`.

- [ ] **Step 4: Commit**

```bash
git add tests/make_test_decks.py test_decks/phase2.pptx
git commit -m "test: phase2.pptx fixture (chart, group, table, paragraphs, notes)"
```

---

## Section 2: Snapshot v2

Six small TDD increments. After each, run precheck + smoke and confirm green.

### Task 2: snapshot adds `occupied_rects`

**Files:**
- Modify: `tests/run_smoke.py`
- Modify: `src/modExportSnapshot.bas`

- [ ] **Step 1: Append failing test**

Add after `test_executor_end_to_end` and before `def main()`:

```python
def test_snapshot_occupied_rects():
    print("test_snapshot_occupied_rects")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        for sl in snap["slides"]:
            assert "occupied_rects" in sl, f"slide {sl['slide_number']} missing occupied_rects"
            for r in sl["occupied_rects"]:
                for k in ("shape_id", "left", "top", "right", "bottom"):
                    assert k in r, f"occupied_rect missing {k}"
                assert r["right"] >= r["left"], "right < left"
                assert r["bottom"] >= r["top"], "bottom < top"
        print("  ok  [occupied_rects on every slide]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add the call in `main()` after `test_executor_end_to_end()`.

- [ ] **Step 2: Run; expect failure (`occupied_rects` missing)**

```bash
python tests/run_smoke.py
```

- [ ] **Step 3: Modify `BuildSlideDict` in `src/modExportSnapshot.bas`**

Add this line right before `d.Add "shapes", BuildShapesCollection(sl)`:

```vb
    d.Add "occupied_rects", BuildOccupiedRects(sl)
```

Add helper at the end of the module:

```vb
Private Function BuildOccupiedRects(sl As Slide) As Collection
    Dim col As New Collection
    Dim sh As Shape
    For Each sh In sl.Shapes
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "shape_id", sh.Id
        d.Add "left", CDbl(sh.Left)
        d.Add "top", CDbl(sh.Top)
        d.Add "right", CDbl(sh.Left + sh.Width)
        d.Add "bottom", CDbl(sh.Top + sh.Height)
        col.Add d
    Next sh
    Set BuildOccupiedRects = col
End Function
```

- [ ] **Step 4: Sync + precheck + smoke**

```bash
powershell -Command "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"
python update_macros.py
python tools/precheck_carrier.py
python tests/run_smoke.py
```

All tests pass.

- [ ] **Step 5: Commit**

```bash
git add tests/run_smoke.py src/modExportSnapshot.bas PPT_AI_Editor.pptm
git commit -m "feat: snapshot adds occupied_rects per slide"
```

---

### Task 3: snapshot adds `speaker_notes`

**Files:**
- Modify: `tests/run_smoke.py`
- Modify: `src/modExportSnapshot.bas`

- [ ] **Step 1: Append failing test**

```python
def test_snapshot_speaker_notes():
    print("test_snapshot_speaker_notes")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        s1 = snap["slides"][0]
        assert "speaker_notes" in s1, "slide 1 missing speaker_notes"
        assert "Q3" in s1["speaker_notes"], f"unexpected notes: {s1['speaker_notes']!r}"
        # Slides without notes still have the key with empty string
        s2 = snap["slides"][1]
        assert s2["speaker_notes"] == "" or "speaker_notes" in s2
        print("  ok  [speaker_notes present]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Modify `BuildSlideDict` to include notes**

Add after the `occupied_rects` line:

```vb
    d.Add "speaker_notes", BuildSpeakerNotes(sl)
```

Add helper:

```vb
Private Function BuildSpeakerNotes(sl As Slide) As String
    On Error Resume Next
    Dim notesText As String
    notesText = ""
    Dim notesPg As Object
    Set notesPg = sl.NotesPage
    Dim ph As Object
    Dim i As Long
    For i = 1 To notesPg.Shapes.Placeholders.Count
        Set ph = notesPg.Shapes.Placeholders(i)
        If ph.HasTextFrame Then
            If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
                notesText = ph.TextFrame.TextRange.Text
                Exit For
            End If
        End If
    Next i
    BuildSpeakerNotes = notesText
End Function
```

- [ ] **Step 4: Sync + precheck + smoke**

- [ ] **Step 5: Commit**

```bash
git commit -am "feat: snapshot adds speaker_notes per slide"
```

---

### Task 4: snapshot adds `paragraphs[]` (text + bullet + indent + runs)

**Files:**
- Modify: `tests/run_smoke.py`
- Modify: `src/modExportSnapshot.bas`

- [ ] **Step 1: Append failing test**

```python
def test_snapshot_paragraphs():
    print("test_snapshot_paragraphs")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body = next(s for s in snap["slides"][0]["shapes"]
                    if s.get("text", "").startswith("First point"))
        assert "paragraphs" in body, "body missing paragraphs"
        ps = body["paragraphs"]
        assert_eq(len(ps), 3, "paragraph count")
        assert_eq(ps[0]["text"].strip(), "First point about revenue", "p0 text")
        assert_eq(ps[2]["text"].strip(), "Third point about headcount", "p2 text")
        for p in ps:
            for k in ("index", "text", "bullet_style", "indent_level", "runs"):
                assert k in p, f"paragraph missing {k}"
            assert isinstance(p["index"], int)
            assert isinstance(p["indent_level"], int)
            assert len(p["runs"]) >= 1, "no runs"
            for run in p["runs"]:
                assert "text" in run and "font" in run
        print("  ok  [paragraphs with text/bullet/indent/runs]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Extend `BuildShapeDict` and add helpers**

In `BuildShapeDict`, replace the existing text block:

```vb
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            d.Add "text", sh.TextFrame.TextRange.Text
            d.Add "font", BuildFontDict(sh.TextFrame.TextRange.Font)
        End If
    End If
```

with:

```vb
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            d.Add "text", sh.TextFrame.TextRange.Text
            d.Add "font", BuildFontDict(sh.TextFrame.TextRange.Font)
            d.Add "paragraphs", BuildParagraphsCollection(sh.TextFrame.TextRange)
        End If
    End If
```

Add helpers:

```vb
Private Function BuildParagraphsCollection(tr As TextRange) As Collection
    Dim col As New Collection
    Dim n As Long
    n = tr.Paragraphs().Count
    Dim i As Long
    For i = 1 To n
        col.Add BuildParagraphDict(tr.Paragraphs(i), i - 1)
    Next i
    Set BuildParagraphsCollection = col
End Function

Private Function BuildParagraphDict(para As TextRange, zeroIdx As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "index", zeroIdx
    d.Add "text", para.Text
    d.Add "bullet_style", BulletStyleName(para.ParagraphFormat.Bullet.Type, para.ParagraphFormat.Bullet.Style)
    d.Add "indent_level", CLng(para.IndentLevel) - 1  ' VBA gives 1-5; we expose 0-4
    d.Add "runs", BuildRunsCollection(para)
    Set BuildParagraphDict = d
End Function

Private Function BulletStyleName(btype As Long, bstyle As Long) As String
    ' ppBulletNone = 0, ppBulletUnnumbered = 1, ppBulletNumbered = 2, ppBulletPicture = 3
    Select Case btype
        Case 0
            BulletStyleName = "none"
        Case 2
            BulletStyleName = "number"
        Case 3
            BulletStyleName = "image"
        Case Else
            ' Use bullet character to refine ppBulletUnnumbered
            ' ppBulletStyleNumbered, ppBulletStyleAlphaUcParenBoth, etc. are not always set
            BulletStyleName = "disc"
    End Select
End Function

Private Function BuildRunsCollection(para As TextRange) As Collection
    Dim col As New Collection
    Dim n As Long
    n = para.Runs().Count
    Dim i As Long
    For i = 1 To n
        Dim run As TextRange
        Set run = para.Runs(i)
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "text", run.Text
        d.Add "font", BuildFontDict(run.Font)
        col.Add d
    Next i
    Set BuildRunsCollection = col
End Function
```

- [ ] **Step 4: Sync + precheck + smoke**

- [ ] **Step 5: Commit**

```bash
git commit -am "feat: snapshot adds paragraphs[] with bullet/indent/runs"
```

---

### Task 5: snapshot adds `group_children`

**Files:**
- Modify: `tests/run_smoke.py`
- Modify: `src/modExportSnapshot.bas`

- [ ] **Step 1: Append failing test**

The phase2 fixture has 3 ungrouped boxes on slide 3. To exercise the `group_children` path, programmatically group them in the test (using PowerPoint COM directly), then verify snapshot reports them as children.

```python
def test_snapshot_group_children():
    print("test_snapshot_group_children")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        # Group the 3 boxes on slide 3 by name — note: VBA-side IDs differ
        # We'll group via COM Range API
        s3 = deck.Slides(3)
        ids = [sh.Id for sh in s3.Shapes]
        names = [sh.Name for sh in s3.Shapes]
        # Use Range with all shape names
        rng = s3.Shapes.Range(names)
        grp = rng.Group()
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        slide3 = snap["slides"][2]
        group_shape = next((s for s in slide3["shapes"] if "group_children" in s), None)
        assert group_shape is not None, "no shape with group_children"
        kids = group_shape["group_children"]
        assert len(kids) == 3, f"expected 3 children, got {len(kids)}"
        texts = sorted([k.get("text", "").strip() for k in kids])
        assert_eq(texts, ["A", "B", "C"], "group child texts")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Extend `BuildShapeDict`**

Add after the `If sh.Type = msoPicture Then` block:

```vb
    If sh.Type = msoGroup Then
        d.Add "group_children", BuildGroupChildren(sh)
    End If
```

Add helper:

```vb
Private Function BuildGroupChildren(sh As Shape) As Collection
    Dim col As New Collection
    Dim child As Shape
    For Each child In sh.GroupItems
        col.Add BuildShapeDict(child)
    Next child
    Set BuildGroupChildren = col
End Function
```

- [ ] **Step 4: Sync + precheck + smoke**

- [ ] **Step 5: Commit**

```bash
git commit -am "feat: snapshot adds group_children for grouped shapes"
```

---

### Task 6: snapshot adds `chart {}` (native only)

**Files:**
- Modify: `tests/run_smoke.py`
- Modify: `src/modExportSnapshot.bas`

- [ ] **Step 1: Append failing test**

```python
def test_snapshot_chart():
    print("test_snapshot_chart")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        s2 = snap["slides"][1]
        chart_shape = next((s for s in s2["shapes"] if s["type"] == "chart"), None)
        assert chart_shape is not None, "no chart shape on slide 2"
        assert "chart" in chart_shape, "missing chart{} key"
        ch = chart_shape["chart"]
        assert ch["is_native"] is True, "is_native must be True"
        for k in ("type", "title", "axis_titles", "legend_position", "series"):
            assert k in ch, f"chart missing {k}"
        assert isinstance(ch["series"], list) and len(ch["series"]) >= 1
        s0 = ch["series"][0]
        assert s0["name"] == "FY24", f"series name {s0['name']!r}"
        assert s0["categories"] == ["Q1", "Q2", "Q3", "Q4"], "categories"
        assert s0["values"] == [100, 110, 120, 130] or s0["values"] is None
        print("  ok  [chart{} with type/title/axis/legend/series]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Extend `BuildShapeDict`**

Add after the `If sh.HasTable Then` block, replacing the existing chart-related portion:

```vb
    If sh.HasChart Then
        d.Add "chart", BuildChartDict(sh.Chart)
    End If
```

Add helper:

```vb
Private Function BuildChartDict(ch As Chart) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "is_native", True
    d.Add "type", ChartTypeName(ch.ChartType)

    On Error Resume Next
    If ch.HasTitle Then
        d.Add "title", ch.ChartTitle.Text
    Else
        d.Add "title", Null
    End If
    Err.Clear

    Dim ax As Object
    Set ax = CreateObject("Scripting.Dictionary")
    Dim xt As String, yt As String
    xt = "": yt = ""
    On Error Resume Next
    If ch.HasAxis(1) Then  ' xlCategory = 1
        If ch.Axes(1).HasTitle Then xt = ch.Axes(1).AxisTitle.Text
    End If
    If ch.HasAxis(2) Then  ' xlValue = 2
        If ch.Axes(2).HasTitle Then yt = ch.Axes(2).AxisTitle.Text
    End If
    Err.Clear
    On Error GoTo 0
    ax.Add "x", xt
    ax.Add "y", yt
    d.Add "axis_titles", ax

    On Error Resume Next
    Dim leg As String: leg = "none"
    If ch.HasLegend Then leg = LegendPositionName(ch.Legend.Position)
    Err.Clear
    On Error GoTo 0
    d.Add "legend_position", leg

    d.Add "series", BuildSeriesCollection(ch)

    Set BuildChartDict = d
End Function

Private Function BuildSeriesCollection(ch As Chart) As Collection
    Dim col As New Collection
    On Error Resume Next
    Dim n As Long: n = ch.SeriesCollection.Count
    Dim i As Long
    For i = 1 To n
        Dim s As Object
        Set s = CreateObject("Scripting.Dictionary")
        s.Add "name", ch.SeriesCollection(i).Name
        Dim cats As Variant
        cats = ch.SeriesCollection(i).XValues
        s.Add "categories", VariantArrayToCollection(cats)
        Dim vals As Variant
        vals = ch.SeriesCollection(i).Values
        s.Add "values", VariantArrayToCollection(vals)
        col.Add s
    Next i
    Err.Clear
    On Error GoTo 0
    Set BuildSeriesCollection = col
End Function

Private Function VariantArrayToCollection(arr As Variant) As Variant
    On Error Resume Next
    Dim col As New Collection
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        col.Add arr(i)
    Next i
    If Err.Number <> 0 Then
        Err.Clear
        VariantArrayToCollection = Null
        Exit Function
    End If
    Set VariantArrayToCollection = col
End Function

Private Function ChartTypeName(t As Long) As String
    Select Case t
        Case 51: ChartTypeName = "columnClustered"     ' xlColumnClustered
        Case 52: ChartTypeName = "columnStacked"        ' xlColumnStacked
        Case 4:  ChartTypeName = "line"                 ' xlLine
        Case 5:  ChartTypeName = "pie"                  ' xlPie
        Case 57: ChartTypeName = "barClustered"         ' xlBarClustered
        Case 1:  ChartTypeName = "area"                 ' xlArea
        Case -4169: ChartTypeName = "scatter"           ' xlXYScatter
        Case Else: ChartTypeName = "type_" & t
    End Select
End Function

Private Function LegendPositionName(p As Long) As String
    Select Case p
        Case -4131: LegendPositionName = "left"          ' xlLegendPositionLeft
        Case -4152: LegendPositionName = "right"          ' xlLegendPositionRight
        Case -4160: LegendPositionName = "top"            ' xlLegendPositionTop
        Case -4107: LegendPositionName = "bottom"         ' xlLegendPositionBottom
        Case 2:     LegendPositionName = "corner"
        Case Else:  LegendPositionName = "right"
    End Select
End Function
```

Also fix `ClassifyShapeType` so `chart` is recognized BEFORE `picture` to avoid charts being classified as pictures: it already does (`ElseIf sh.HasChart Then ClassifyShapeType = "chart"` after picture check) — verify and correct order if needed: `HasChart` check must come before `Type = msoPicture` if both can be true on the same shape. Edit `ClassifyShapeType`:

```vb
Private Function ClassifyShapeType(sh As Shape) As String
    If sh.Type = msoPlaceholder Then
        Select Case sh.PlaceholderFormat.Type
            Case ppPlaceholderTitle, ppPlaceholderCenterTitle
                ClassifyShapeType = "title"
            Case ppPlaceholderBody, ppPlaceholderObject, ppPlaceholderSubtitle
                ClassifyShapeType = "body"
            Case Else
                ClassifyShapeType = "other"
        End Select
    ElseIf sh.HasChart Then
        ClassifyShapeType = "chart"
    ElseIf sh.HasTable Then
        ClassifyShapeType = "table"
    ElseIf sh.Type = msoPicture Then
        ClassifyShapeType = "picture"
    ElseIf sh.HasTextFrame Then
        ClassifyShapeType = "textbox"
    Else
        ClassifyShapeType = "other"
    End If
End Function
```

- [ ] **Step 4: Sync + precheck + smoke**

If `BuildSeriesCollection` triggers Excel workbook open in your environment (visible in Task Manager), set `s.Add "values", Null` and `s.Add "categories", Null` instead of attempting to read them, and rerun. Document the limitation in commit message.

- [ ] **Step 5: Commit**

```bash
git commit -am "feat: snapshot adds chart{} for native HasChart shapes"
```

---

### Task 7: snapshot adds `table_extra` (merged cells)

**Files:**
- Modify: `tests/run_smoke.py`
- Modify: `src/modExportSnapshot.bas`

- [ ] **Step 1: Append failing test**

The phase2 fixture has no merged cells yet. The test merges row 0 cells (col 1+2) at runtime via COM, then verifies snapshot reports the merge.

```python
def test_snapshot_table_extra():
    print("test_snapshot_table_extra")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        # Merge cells (1,2) and (1,3) on the table on slide 4
        s4 = deck.Slides(4)
        tbl = next(sh for sh in s4.Shapes if sh.HasTable).Table
        tbl.Cell(1, 2).Merge tbl.Cell(1, 3)

        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        slide4 = snap["slides"][3]
        tab = next(s for s in slide4["shapes"] if s["type"] == "table")
        assert "table_extra" in tab, "missing table_extra"
        merges = tab["table_extra"]["merged_cells"]
        assert len(merges) >= 1, f"no merges reported"
        m = merges[0]
        assert_eq(m["row"], 1, "merged row")
        assert_eq(m["col"], 2, "merged col")
        assert m["col_span"] == 2, f"col_span {m['col_span']}"
        print("  ok  [table_extra reports merged cells]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Extend `BuildShapeDict`**

Replace the existing table block:

```vb
    If sh.HasTable Then
        d.Add "table", BuildTableDict(sh.Table)
    End If
```

with:

```vb
    If sh.HasTable Then
        d.Add "table", BuildTableDict(sh.Table)
        d.Add "table_extra", BuildTableExtra(sh.Table)
    End If
```

Add helper:

```vb
Private Function BuildTableExtra(tbl As Table) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim merges As New Collection

    Dim r As Long, c As Long
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    For r = 1 To tbl.Rows.Count
        For c = 1 To tbl.Columns.Count
            Dim cellObj As Object
            Set cellObj = tbl.Cell(r, c)
            On Error Resume Next
            Dim isMerged As Boolean: isMerged = False
            ' Cell.Merged returns msoTrue/msoFalse
            isMerged = (cellObj.Merged = msoTrue)
            If isMerged Then
                Dim rs As Long, cs As Long
                rs = 1: cs = 1
                rs = cellObj.RowSpan
                cs = cellObj.ColSpan
                If rs > 1 Or cs > 1 Then
                    Dim key As String
                    key = r & "_" & c
                    If Not seen.Exists(key) Then
                        seen.Add key, 1
                        Dim m As Object
                        Set m = CreateObject("Scripting.Dictionary")
                        m.Add "row", r
                        m.Add "col", c
                        m.Add "row_span", rs
                        m.Add "col_span", cs
                        merges.Add m
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo 0
        Next c
    Next r

    d.Add "merged_cells", merges
    Set BuildTableExtra = d
End Function
```

- [ ] **Step 4: Sync + precheck + smoke**

- [ ] **Step 5: Commit**

```bash
git commit -am "feat: snapshot adds table_extra with merged_cells"
```

---

## Section 3: Granular text actions (5 tasks for 8 actions)

Each task in this section adds a couple of actions to a new module `src/modActionsText.bas`, plus validation + dispatch cases in `src/modExecuteInstructions.bas`, plus a smoke test.

The shared helper for finding a paragraph by zero-based index:

```vb
' Helper used by every text-action Sub. Add to modActionsText.bas at top.
Public Function FindParagraph(slideNum As Long, shapeId As Long, zeroIdx As Long) As TextRange
    Set FindParagraph = Nothing
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Exit Function
    If Not sh.HasTextFrame Then Exit Function
    Dim n As Long: n = sh.TextFrame.TextRange.Paragraphs().Count
    If zeroIdx < 0 Or zeroIdx >= n Then Exit Function
    Set FindParagraph = sh.TextFrame.TextRange.Paragraphs(zeroIdx + 1)
End Function
```

### Task 8: `modActionsText` skeleton + `Do_set_paragraph_text`

**Files:**
- Create: `src/modActionsText.bas`
- Modify: `src/modExecuteInstructions.bas` (validate + dispatch)
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_set_paragraph_text():
    print("test_action_set_paragraph_text")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body = next(s for s in snap["slides"][0]["shapes"]
                    if s.get("text", "").startswith("First point"))
        sid = body["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_paragraph_text", 1, sid, 1, "REVISED")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body2 = next(s for s in snap2["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(body2["paragraphs"][1]["text"].strip(), "REVISED", "paragraph 1 text")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsText.bas`**

```vb
Attribute VB_Name = "modActionsText"
Option Explicit

Public Function FindParagraph(slideNum As Long, shapeId As Long, zeroIdx As Long) As TextRange
    Set FindParagraph = Nothing
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Exit Function
    If Not sh.HasTextFrame Then Exit Function
    Dim n As Long: n = sh.TextFrame.TextRange.Paragraphs().Count
    If zeroIdx < 0 Or zeroIdx >= n Then Exit Function
    Set FindParagraph = sh.TextFrame.TextRange.Paragraphs(zeroIdx + 1)
End Function

Public Sub Do_set_paragraph_text(slideNum As Long, shapeId As Long, _
                                 paragraphIndex As Long, value As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_paragraph_text", "paragraph not found"
    p.Text = value
End Sub
```

- [ ] **Step 4: Add validation + dispatch in `src/modExecuteInstructions.bas`**

In `ValidateAction`, add a case (alphabetically near `set_text`):

```vb
        Case "set_paragraph_text"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

In `DispatchAction`:

```vb
        Case "set_paragraph_text"
            modActionsText.Do_set_paragraph_text CLng(act("slide")), CLng(act("shape_id")), _
                                                 CLng(act("paragraph_index")), CStr(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsText + Do_set_paragraph_text"
```

---

### Task 9: `Do_add_paragraph` + `Do_delete_paragraph`

**Files:**
- Modify: `src/modActionsText.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_paragraph_add_delete():
    print("test_action_paragraph_add_delete")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body = next(s for s in snap["slides"][0]["shapes"]
                    if s.get("text", "").startswith("First point"))
        sid = body["shape_id"]

        app.Run("PPT_AI_Editor!Do_add_paragraph", 1, sid, 0, "INSERTED")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body2 = next(s for s in snap2["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(len(body2["paragraphs"]), 4, "paragraph count after add")
        assert_eq(body2["paragraphs"][1]["text"].strip(), "INSERTED", "inserted text")

        app.Run("PPT_AI_Editor!Do_delete_paragraph", 1, sid, 1)
        snap3 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body3 = next(s for s in snap3["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(len(body3["paragraphs"]), 3, "paragraph count after delete")
        # The originally-first paragraph should remain
        assert "First point" in body3["paragraphs"][0]["text"]
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsText.bas`**

```vb
Public Sub Do_add_paragraph(slideNum As Long, shapeId As Long, _
                            afterParagraphIndex As Long, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 3001, "Do_add_paragraph", "shape not found"
    If Not sh.HasTextFrame Then Err.Raise vbObjectError + 3002, "Do_add_paragraph", "no text frame"

    Dim tr As TextRange: Set tr = sh.TextFrame.TextRange
    Dim newText As String: newText = vbCrLf & value
    If afterParagraphIndex < 0 Then
        ' Insert before first paragraph: prepend
        tr.Text = value & vbCrLf & tr.Text
        Exit Sub
    End If

    Dim n As Long: n = tr.Paragraphs().Count
    If afterParagraphIndex >= n Then
        ' Append at end
        tr.Text = tr.Text & vbCrLf & value
        Exit Sub
    End If

    Dim p As TextRange: Set p = tr.Paragraphs(afterParagraphIndex + 1)
    p.InsertAfter newText
End Sub

Public Sub Do_delete_paragraph(slideNum As Long, shapeId As Long, paragraphIndex As Long)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_delete_paragraph", "paragraph not found"
    p.Delete
End Sub
```

- [ ] **Step 4: Add validation + dispatch**

`ValidateAction`:

```vb
        Case "add_paragraph"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "after_paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_paragraph"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "add_paragraph"
            modActionsText.Do_add_paragraph CLng(act("slide")), CLng(act("shape_id")), _
                                            CLng(act("after_paragraph_index")), CStr(act("value"))
        Case "delete_paragraph"
            modActionsText.Do_delete_paragraph CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_add_paragraph + Do_delete_paragraph"
```

---

### Task 10: `Do_set_bullet_style` + `Do_set_indent_level`

**Files:**
- Modify: `src/modActionsText.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_bullet_indent():
    print("test_action_bullet_indent")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body = next(s for s in snap["slides"][0]["shapes"]
                    if s.get("text", "").startswith("First point"))
        sid = body["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_bullet_style", 1, sid, 0, "number")
        app.Run("PPT_AI_Editor!Do_set_indent_level", 1, sid, 1, 1)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body2 = next(s for s in snap2["slides"][0]["shapes"] if s["shape_id"] == sid)
        assert_eq(body2["paragraphs"][0]["bullet_style"], "number", "p0 bullet style")
        assert_eq(body2["paragraphs"][1]["indent_level"], 1, "p1 indent level")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsText.bas`**

```vb
Public Sub Do_set_bullet_style(slideNum As Long, shapeId As Long, _
                               paragraphIndex As Long, value As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_bullet_style", "paragraph not found"
    Dim b As Object: Set b = p.ParagraphFormat.Bullet
    Select Case LCase(value)
        Case "none"
            b.Type = 0  ' ppBulletNone
        Case "number"
            b.Type = 2  ' ppBulletNumbered
        Case "letter"
            b.Type = 2
            b.Style = 16  ' ppBulletStyleAlphaUcParenBoth (best approximation)
        Case "disc", "bullet"
            b.Type = 1  ' ppBulletUnnumbered
            b.Character = 8226  ' • bullet char
        Case "square"
            b.Type = 1
            b.Character = 9632  ' ■
        Case "dash"
            b.Type = 1
            b.Character = 8211  ' –
        Case Else
            Err.Raise vbObjectError + 3003, "Do_set_bullet_style", "unknown bullet style: " & value
    End Select
End Sub

Public Sub Do_set_indent_level(slideNum As Long, shapeId As Long, _
                               paragraphIndex As Long, value As Long)
    If value < 0 Or value > 4 Then Err.Raise vbObjectError + 3004, "Do_set_indent_level", "indent_level must be 0..4"
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_indent_level", "paragraph not found"
    p.IndentLevel = value + 1   ' VBA expects 1..5
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_bullet_style"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_indent_level"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "set_bullet_style"
            modActionsText.Do_set_bullet_style CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CStr(act("value"))
        Case "set_indent_level"
            modActionsText.Do_set_indent_level CLng(act("slide")), CLng(act("shape_id")), _
                                               CLng(act("paragraph_index")), CLng(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_set_bullet_style + Do_set_indent_level"
```

---

### Task 11: paragraph-level font ops

**Files:**
- Modify: `src/modActionsText.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_paragraph_font():
    print("test_action_paragraph_font")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body = next(s for s in snap["slides"][0]["shapes"]
                    if s.get("text", "").startswith("First point"))
        sid = body["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_paragraph_font_size", 1, sid, 0, 30)
        app.Run("PPT_AI_Editor!Do_set_paragraph_font_color", 1, sid, 1, "#FF0000")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body2 = next(s for s in snap2["slides"][0]["shapes"] if s["shape_id"] == sid)
        # Per-paragraph font is reflected in run-level font
        run0 = body2["paragraphs"][0]["runs"][0]
        run1 = body2["paragraphs"][1]["runs"][0]
        assert_eq(int(run0["font"]["size"]), 30, "p0 size after set")
        assert_eq(run1["font"]["color"].upper(), "#FF0000", "p1 color after set")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsText.bas`**

```vb
Public Sub Do_set_paragraph_font_size(slideNum As Long, shapeId As Long, _
                                      paragraphIndex As Long, value As Long)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_paragraph_font_size", "paragraph not found"
    p.Font.Size = value
End Sub

Public Sub Do_set_paragraph_font_color(slideNum As Long, shapeId As Long, _
                                       paragraphIndex As Long, hexValue As String)
    Dim p As TextRange: Set p = FindParagraph(slideNum, shapeId, paragraphIndex)
    If p Is Nothing Then Err.Raise vbObjectError + 3001, "Do_set_paragraph_font_color", "paragraph not found"
    p.Font.Color.RGB = modActions.HexToRgb(hexValue)
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_paragraph_font_size"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_paragraph_font_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "paragraph_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "set_paragraph_font_size"
            modActionsText.Do_set_paragraph_font_size CLng(act("slide")), CLng(act("shape_id")), _
                                                       CLng(act("paragraph_index")), CLng(act("value"))
        Case "set_paragraph_font_color"
            modActionsText.Do_set_paragraph_font_color CLng(act("slide")), CLng(act("shape_id")), _
                                                        CLng(act("paragraph_index")), CStr(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: paragraph-level font size + color"
```

---

### Task 12: `Do_find_replace_text`

**Files:**
- Modify: `src/modActionsText.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_find_replace_text():
    print("test_action_find_replace_text")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        # Phase 2 fixture has "First point about revenue" on slide 1
        # Replace "revenue" → "REVENUE" deck-wide
        app.Run("PPT_AI_Editor!Do_find_replace_text", "deck", "revenue", "REVENUE")
        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        body = next(s for s in after["slides"][0]["shapes"]
                    if "REVENUE" in s.get("text", ""))
        assert body, "find_replace did not change any text"
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsText.bas`**

```vb
Public Sub Do_find_replace_text(scope As String, findText As String, replaceText As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim slideNumFilter As Long: slideNumFilter = 0  ' 0 = all
    If LCase(Left(scope, 6)) = "slide:" Then
        slideNumFilter = CLng(Mid(scope, 7))
        If slideNumFilter < 1 Or slideNumFilter > pres.Slides.Count Then
            Err.Raise vbObjectError + 3005, "Do_find_replace_text", "slide_out_of_range"
        End If
    ElseIf LCase(scope) <> "deck" Then
        Err.Raise vbObjectError + 3006, "Do_find_replace_text", "scope must be 'deck' or 'slide:N'"
    End If

    Dim i As Long
    For i = 1 To pres.Slides.Count
        If slideNumFilter = 0 Or slideNumFilter = i Then
            Dim sh As Shape
            For Each sh In pres.Slides(i).Shapes
                ReplaceInShape sh, findText, replaceText
            Next sh
        End If
    Next i
End Sub

Private Sub ReplaceInShape(sh As Shape, findText As String, replaceText As String)
    On Error Resume Next
    If sh.HasTextFrame Then
        If sh.TextFrame.HasText Then
            Dim t As String: t = sh.TextFrame.TextRange.Text
            If InStr(t, findText) > 0 Then
                sh.TextFrame.TextRange.Text = Replace(t, findText, replaceText)
            End If
        End If
    End If
    If sh.HasTable Then
        Dim r As Long, c As Long
        For r = 1 To sh.Table.Rows.Count
            For c = 1 To sh.Table.Columns.Count
                Dim cellShape As Shape: Set cellShape = sh.Table.Cell(r, c).Shape
                If cellShape.HasTextFrame Then
                    If cellShape.TextFrame.HasText Then
                        Dim ct As String: ct = cellShape.TextFrame.TextRange.Text
                        If InStr(ct, findText) > 0 Then
                            cellShape.TextFrame.TextRange.Text = Replace(ct, findText, replaceText)
                        End If
                    End If
                End If
            Next c
        Next r
    End If
    If sh.Type = msoGroup Then
        Dim child As Shape
        For Each child In sh.GroupItems
            ReplaceInShape child, findText, replaceText
        Next child
    End If
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "find_replace_text"
            ValidateAction = RequireFields(act, Array("scope", "find", "replace"))
```

`DispatchAction`:

```vb
        Case "find_replace_text"
            modActionsText.Do_find_replace_text CStr(act("scope")), CStr(act("find")), CStr(act("replace"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_find_replace_text (scope=deck or slide:N)"
```

---

## Section 4: Layout actions (5 tasks for 10 actions)

### Task 13: `modActionsLayout` + `Do_align_shapes`

**Files:**
- Create: `src/modActionsLayout.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_align_shapes():
    print("test_action_align_shapes")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes = sorted(
            [s for s in snap["slides"][2]["shapes"] if s.get("text") in ("A", "B", "C")],
            key=lambda s: s["pos"]["left"],
        )
        ids = [b["shape_id"] for b in boxes]
        # Make second box's top different from others first
        app.Run("PPT_AI_Editor!Do_move_shape", 3, ids[1], boxes[1]["pos"]["left"], boxes[1]["pos"]["top"] + 30.0)
        app.Run("PPT_AI_Editor!Do_align_shapes", 3, ids, "top")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes2 = [s for s in snap2["slides"][2]["shapes"] if s["shape_id"] in ids]
        tops = [b["pos"]["top"] for b in boxes2]
        assert max(tops) - min(tops) < 0.5, f"tops not aligned: {tops}"
        print("  ok  [align top]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsLayout.bas`**

```vb
Attribute VB_Name = "modActionsLayout"
Option Explicit

Public Sub Do_align_shapes(slideNum As Long, shapeIds As Variant, anchor As String)
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 2 Then Err.Raise vbObjectError + 4001, "Do_align_shapes", "need >=2 shapes"
    Dim ref As Shape: Set ref = shapes(0)
    Dim i As Long
    For i = 1 To n - 1
        Select Case LCase(anchor)
            Case "left":    shapes(i).Left = ref.Left
            Case "right":   shapes(i).Left = ref.Left + ref.Width - shapes(i).Width
            Case "top":     shapes(i).Top = ref.Top
            Case "bottom":  shapes(i).Top = ref.Top + ref.Height - shapes(i).Height
            Case "hcenter": shapes(i).Left = ref.Left + (ref.Width - shapes(i).Width) / 2
            Case "vcenter": shapes(i).Top = ref.Top + (ref.Height - shapes(i).Height) / 2
            Case Else: Err.Raise vbObjectError + 4002, "Do_align_shapes", "unknown anchor: " & anchor
        End Select
    Next i
End Sub

Public Function ShapesByIds(slideNum As Long, shapeIds As Variant, ByRef out() As Shape) As Long
    Dim ids() As Long
    Dim cnt As Long: cnt = NormalizeIdsArray(shapeIds, ids)
    ReDim out(0 To cnt - 1)
    Dim i As Long, found As Long: found = 0
    For i = 0 To cnt - 1
        Dim sh As Shape: Set sh = modActions.FindShape(slideNum, ids(i))
        If Not sh Is Nothing Then
            Set out(found) = sh
            found = found + 1
        End If
    Next i
    If found < cnt Then ReDim Preserve out(0 To found - 1)
    ShapesByIds = found
End Function

Public Function NormalizeIdsArray(v As Variant, ByRef out() As Long) As Long
    Dim col As Object
    If TypeName(v) = "Collection" Then
        Set col = v
        ReDim out(0 To col.Count - 1)
        Dim i As Long
        For i = 1 To col.Count
            out(i - 1) = CLng(col(i))
        Next i
        NormalizeIdsArray = col.Count
    ElseIf IsArray(v) Then
        Dim lo As Long, hi As Long
        lo = LBound(v): hi = UBound(v)
        ReDim out(0 To hi - lo)
        For i = lo To hi
            out(i - lo) = CLng(v(i))
        Next i
        NormalizeIdsArray = hi - lo + 1
    Else
        ReDim out(0 To 0)
        out(0) = CLng(v)
        NormalizeIdsArray = 1
    End If
End Function
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "align_shapes"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "anchor"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
```

`DispatchAction`:

```vb
        Case "align_shapes"
            modActionsLayout.Do_align_shapes CLng(act("slide")), act("shape_ids"), CStr(act("anchor"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsLayout + Do_align_shapes"
```

---

### Task 14: `Do_distribute_horizontal` + `Do_distribute_vertical`

**Files:**
- Modify: `src/modActionsLayout.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_distribute():
    print("test_action_distribute")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes = sorted(
            [s for s in snap["slides"][2]["shapes"] if s.get("text") in ("A", "B", "C")],
            key=lambda s: s["pos"]["left"],
        )
        ids = [b["shape_id"] for b in boxes]
        # Move B way to the left
        app.Run("PPT_AI_Editor!Do_move_shape", 3, ids[1], 100.0, boxes[1]["pos"]["top"])
        app.Run("PPT_AI_Editor!Do_distribute_horizontal", 3, ids)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes2 = sorted(
            [s for s in snap2["slides"][2]["shapes"] if s["shape_id"] in ids],
            key=lambda s: s["pos"]["left"],
        )
        gap1 = boxes2[1]["pos"]["left"] - boxes2[0]["pos"]["left"]
        gap2 = boxes2[2]["pos"]["left"] - boxes2[1]["pos"]["left"]
        assert abs(gap1 - gap2) < 1.0, f"gaps unequal: {gap1} vs {gap2}"
        print("  ok  [distribute horizontal]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsLayout.bas`**

```vb
Public Sub Do_distribute_horizontal(slideNum As Long, shapeIds As Variant)
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 3 Then Err.Raise vbObjectError + 4003, "Do_distribute_horizontal", "need >=3 shapes"
    SortShapesByLeft shapes
    Dim minLeft As Single: minLeft = shapes(0).Left
    Dim maxLeft As Single: maxLeft = shapes(n - 1).Left
    Dim step As Single: step = (maxLeft - minLeft) / (n - 1)
    Dim i As Long
    For i = 1 To n - 2
        shapes(i).Left = minLeft + step * i
    Next i
End Sub

Public Sub Do_distribute_vertical(slideNum As Long, shapeIds As Variant)
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 3 Then Err.Raise vbObjectError + 4003, "Do_distribute_vertical", "need >=3 shapes"
    SortShapesByTop shapes
    Dim minTop As Single: minTop = shapes(0).Top
    Dim maxTop As Single: maxTop = shapes(n - 1).Top
    Dim step As Single: step = (maxTop - minTop) / (n - 1)
    Dim i As Long
    For i = 1 To n - 2
        shapes(i).Top = minTop + step * i
    Next i
End Sub

Private Sub SortShapesByLeft(ByRef arr() As Shape)
    Dim i As Long, j As Long, tmp As Shape
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j).Left < arr(i).Left Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Sub SortShapesByTop(ByRef arr() As Shape)
    Dim i As Long, j As Long, tmp As Shape
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(j).Top < arr(i).Top Then
                Set tmp = arr(i): Set arr(i) = arr(j): Set arr(j) = tmp
            End If
        Next j
    Next i
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "distribute_horizontal", "distribute_vertical"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
```

`DispatchAction`:

```vb
        Case "distribute_horizontal"
            modActionsLayout.Do_distribute_horizontal CLng(act("slide")), act("shape_ids")
        Case "distribute_vertical"
            modActionsLayout.Do_distribute_vertical CLng(act("slide")), act("shape_ids")
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_distribute_horizontal + Do_distribute_vertical"
```

---

### Task 15: `Do_tile_grid` + `Do_fit_to_slide_margins`

**Files:**
- Modify: `src/modActionsLayout.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_tile_and_fit():
    print("test_action_tile_and_fit")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes = [s for s in snap["slides"][2]["shapes"] if s.get("text") in ("A", "B", "C")]
        ids = [b["shape_id"] for b in boxes]
        app.Run("PPT_AI_Editor!Do_tile_grid", 3, ids, 3, 10.0)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        # All three should now have same top
        boxes2 = [s for s in snap2["slides"][2]["shapes"] if s["shape_id"] in ids]
        tops = [b["pos"]["top"] for b in boxes2]
        assert max(tops) - min(tops) < 0.5, f"tile_grid: tops not equal: {tops}"

        # fit_to_slide_margins on the rectangle on slide 4
        rect_id = next(s["shape_id"] for s in snap2["slides"][3]["shapes"] if s.get("text") == "Plain")
        app.Run("PPT_AI_Editor!Do_fit_to_slide_margins", 4, rect_id, 36.0)
        snap3 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        rect3 = next(s for s in snap3["slides"][3]["shapes"] if s["shape_id"] == rect_id)
        slide_w = snap3["deck"]["slide_width_pt"]
        slide_h = snap3["deck"]["slide_height_pt"]
        assert abs(rect3["pos"]["left"] - 36.0) < 1.0, f"left {rect3['pos']['left']}"
        assert abs(rect3["pos"]["width"] - (slide_w - 72.0)) < 1.0, f"width {rect3['pos']['width']}"
        print("  ok  [tile + fit]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsLayout.bas`**

```vb
Public Sub Do_tile_grid(slideNum As Long, shapeIds As Variant, cols As Long, gapPt As Single)
    If cols < 1 Then Err.Raise vbObjectError + 4004, "Do_tile_grid", "cols must be >=1"
    Dim shapes() As Shape
    Dim n As Long: n = ShapesByIds(slideNum, shapeIds, shapes)
    If n < 1 Then Err.Raise vbObjectError + 4001, "Do_tile_grid", "no shapes"
    Dim originX As Single: originX = shapes(0).Left
    Dim originY As Single: originY = shapes(0).Top
    Dim cellW As Single: cellW = shapes(0).Width + gapPt
    Dim cellH As Single: cellH = shapes(0).Height + gapPt
    Dim i As Long
    For i = 0 To n - 1
        Dim row As Long: row = i \ cols
        Dim col As Long: col = i Mod cols
        shapes(i).Left = originX + col * cellW
        shapes(i).Top = originY + row * cellH
    Next i
End Sub

Public Sub Do_fit_to_slide_margins(slideNum As Long, shapeId As Long, marginPt As Single)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 4001, "Do_fit_to_slide_margins", "shape not found"
    Dim sw As Single: sw = ActivePresentation.PageSetup.SlideWidth
    Dim shgt As Single: shgt = ActivePresentation.PageSetup.SlideHeight
    sh.Left = marginPt
    sh.Top = marginPt
    sh.LockAspectRatio = msoFalse
    sh.Width = sw - 2 * marginPt
    sh.Height = shgt - 2 * marginPt
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "tile_grid"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids", "cols", "gap_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "fit_to_slide_margins"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "tile_grid"
            modActionsLayout.Do_tile_grid CLng(act("slide")), act("shape_ids"), _
                                          CLng(act("cols")), CSng(act("gap_pt"))
        Case "fit_to_slide_margins"
            Dim m As Single: m = 36.0
            If act.Exists("margin_pt") Then m = CSng(act("margin_pt"))
            modActionsLayout.Do_fit_to_slide_margins CLng(act("slide")), CLng(act("shape_id")), m
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_tile_grid + Do_fit_to_slide_margins"
```

---

### Task 16: `Do_add_line` + `Do_add_shape`

**Files:**
- Modify: `src/modActionsLayout.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_add_line_and_shape():
    print("test_action_add_line_and_shape")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        before_count_s4 = len(before["slides"][3]["shapes"])

        app.Run("PPT_AI_Editor!Do_add_line", 4, 50.0, 50.0, 500.0, 50.0, "#000000", 1.5)
        app.Run("PPT_AI_Editor!Do_add_shape", 4, "capsule",
                {"left": 100.0, "top": 100.0, "width": 200.0, "height": 60.0},
                "#1F4E79", "#FFFFFF", 2.0)

        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert len(after["slides"][3]["shapes"]) == before_count_s4 + 2, "expected 2 new shapes"
        print("  ok  [add_line + add_shape]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

NOTE: `Application.Run` cannot pass nested dicts directly. We pass the position fields as separate args; update the test to use 4 separate `left/top/width/height` instead of a dict, and adjust the VBA Sub signature accordingly.

Replace the `app.Run("PPT_AI_Editor!Do_add_shape", ...)` call with:

```python
        app.Run("PPT_AI_Editor!Do_add_shape", 4, "capsule",
                100.0, 100.0, 200.0, 60.0,
                "#1F4E79", "#FFFFFF", 2.0)
```

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsLayout.bas`**

```vb
Public Sub Do_add_line(slideNum As Long, x1 As Single, y1 As Single, _
                       x2 As Single, y2 As Single, _
                       hexColor As String, weightPt As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 4005, "Do_add_line", "slide_out_of_range"
    End If
    Dim ln As Shape
    Set ln = pres.Slides(slideNum).Shapes.AddLine(x1, y1, x2, y2)
    ln.Line.ForeColor.RGB = modActions.HexToRgb(hexColor)
    ln.Line.Weight = weightPt
End Sub

Public Sub Do_add_shape(slideNum As Long, kind As String, _
                        leftPt As Single, topPt As Single, _
                        widthPt As Single, heightPt As Single, _
                        fillHex As String, strokeHex As String, _
                        strokeWeight As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 4005, "Do_add_shape", "slide_out_of_range"
    End If
    Dim msoKind As Long: msoKind = ResolveAutoShapeKind(kind)
    Dim sh As Shape
    Set sh = pres.Slides(slideNum).Shapes.AddShape(msoKind, leftPt, topPt, widthPt, heightPt)
    If Len(fillHex) > 0 Then
        sh.Fill.Visible = msoTrue
        sh.Fill.Solid
        sh.Fill.ForeColor.RGB = modActions.HexToRgb(fillHex)
    Else
        sh.Fill.Visible = msoFalse
    End If
    If Len(strokeHex) > 0 Then
        sh.Line.Visible = msoTrue
        sh.Line.ForeColor.RGB = modActions.HexToRgb(strokeHex)
        sh.Line.Weight = strokeWeight
    Else
        sh.Line.Visible = msoFalse
    End If
End Sub

Private Function ResolveAutoShapeKind(kind As String) As Long
    Select Case LCase(kind)
        Case "rect", "rectangle":  ResolveAutoShapeKind = 1   ' msoShapeRectangle
        Case "rrect", "round_rect": ResolveAutoShapeKind = 5  ' msoShapeRoundedRectangle
        Case "oval", "ellipse":     ResolveAutoShapeKind = 9
        Case "circle":              ResolveAutoShapeKind = 9
        Case "capsule":             ResolveAutoShapeKind = 73 ' msoShapeFlowchartTerminator (capsule-ish)
        Case "arrow", "right_arrow": ResolveAutoShapeKind = 13 ' msoShapeRightArrow
        Case "diamond":             ResolveAutoShapeKind = 4
        Case "triangle":            ResolveAutoShapeKind = 7  ' msoShapeIsoscelesTriangle
        Case Else:                  Err.Raise vbObjectError + 4006, "ResolveAutoShapeKind", "unknown kind: " & kind
    End Select
End Function
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "add_line"
            ValidateAction = RequireFields(act, Array("slide", "x1", "y1", "x2", "y2", "color", "weight_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "add_shape"
            ValidateAction = RequireFields(act, Array("slide", "kind", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
```

`DispatchAction`:

```vb
        Case "add_line"
            modActionsLayout.Do_add_line CLng(act("slide")), CSng(act("x1")), CSng(act("y1")), _
                                         CSng(act("x2")), CSng(act("y2")), _
                                         CStr(act("color")), CSng(act("weight_pt"))
        Case "add_shape"
            Dim pos As Object: Set pos = act("pos")
            Dim fh As String: fh = ""
            Dim sh As String: sh = ""
            Dim sw As Single: sw = 1.0
            If act.Exists("fill") Then If Not IsNull(act("fill")) Then fh = CStr(act("fill"))
            If act.Exists("stroke") Then If Not IsNull(act("stroke")) Then sh = CStr(act("stroke"))
            If act.Exists("stroke_weight_pt") Then sw = CSng(act("stroke_weight_pt"))
            modActionsLayout.Do_add_shape CLng(act("slide")), CStr(act("kind")), _
                                          CSng(pos("left")), CSng(pos("top")), _
                                          CSng(pos("width")), CSng(pos("height")), _
                                          fh, sh, sw
```

NOTE: `add_shape` instructions JSON DOES use a nested `pos` dict (LLM-friendly). Only the smoke test invokes the inner Sub directly with positional args. The dispatch unpacks the dict.

Update the smoke test to call via the executor instead of `Do_add_shape` directly:

```python
        instructions = {"actions": [
            {"type": "add_line", "slide": 4, "x1": 50.0, "y1": 50.0, "x2": 500.0, "y2": 50.0,
             "color": "#000000", "weight_pt": 1.5},
            {"type": "add_shape", "slide": 4, "kind": "capsule",
             "pos": {"left": 100.0, "top": 100.0, "width": 200.0, "height": 60.0},
             "fill": "#1F4E79", "stroke": "#FFFFFF", "stroke_weight_pt": 2.0}
        ]}
        summary = app.Run("PPT_AI_Editor!ExecuteFromString", json.dumps(instructions))
        assert "2 applied" in summary, f"summary: {summary}"
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_add_line + Do_add_shape with auto-shape kind map"
```

---

### Task 17: `Do_set_shape_kind`, `Do_clear_slide`, `Do_move_shape_relative`

**Files:**
- Modify: `src/modActionsLayout.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_kind_clear_relative():
    print("test_action_kind_clear_relative")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        # Plain rectangle on slide 4 → set_shape_kind to capsule
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        rect_id = next(s["shape_id"] for s in snap["slides"][3]["shapes"] if s.get("text") == "Plain")
        app.Run("PPT_AI_Editor!Do_set_shape_kind", 4, rect_id, "capsule")
        # No assertion on shape kind from snapshot (we don't expose AutoShapeType yet);
        # just verify the call succeeds without raising.

        # move_shape_relative
        before_top = next(s["pos"]["top"] for s in snap["slides"][3]["shapes"] if s["shape_id"] == rect_id)
        app.Run("PPT_AI_Editor!Do_move_shape_relative", 4, rect_id, 0.0, 50.0)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        new_top = next(s["pos"]["top"] for s in snap2["slides"][3]["shapes"] if s["shape_id"] == rect_id)
        assert abs(new_top - (before_top + 50.0)) < 1.0, f"move_shape_relative: {before_top} -> {new_top}"

        # clear_slide on slide 3 keeping nothing
        app.Run("PPT_AI_Editor!Do_clear_slide", 3, [])
        snap3 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert len(snap3["slides"][2]["shapes"]) == 0, "clear_slide should empty slide 3"
        print("  ok  [set_kind + move_relative + clear_slide]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsLayout.bas`**

```vb
Public Sub Do_set_shape_kind(slideNum As Long, shapeId As Long, kind As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 4001, "Do_set_shape_kind", "shape not found"
    On Error Resume Next
    sh.AutoShapeType = ResolveAutoShapeKind(kind)
    If Err.Number <> 0 Then
        Dim msg As String: msg = "not_an_autoshape: " & Err.Description
        Err.Clear
        Err.Raise vbObjectError + 4007, "Do_set_shape_kind", msg
    End If
    On Error GoTo 0
End Sub

Public Sub Do_clear_slide(slideNum As Long, keepShapeIds As Variant)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 4005, "Do_clear_slide", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)

    Dim keepIds() As Long
    Dim keepCount As Long: keepCount = NormalizeIdsArray(keepShapeIds, keepIds)

    Dim toDelete As New Collection
    Dim sh As Shape, i As Long, keep As Boolean
    For Each sh In sl.Shapes
        keep = False
        For i = 0 To keepCount - 1
            If sh.Id = keepIds(i) Then keep = True: Exit For
        Next i
        If Not keep Then toDelete.Add sh.Id
    Next sh

    For i = 1 To toDelete.Count
        Dim victim As Shape: Set victim = modActions.FindShape(slideNum, CLng(toDelete(i)))
        If Not victim Is Nothing Then victim.Delete
    Next i
End Sub

Public Sub Do_move_shape_relative(slideNum As Long, shapeId As Long, dxPt As Single, dyPt As Single)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 4001, "Do_move_shape_relative", "shape not found"
    sh.Left = sh.Left + dxPt
    sh.Top = sh.Top + dyPt
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_shape_kind"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "kind"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "clear_slide"
            ValidateAction = RequireFields(act, Array("slide"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "move_shape_relative"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "dx_pt", "dy_pt"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "set_shape_kind"
            modActionsLayout.Do_set_shape_kind CLng(act("slide")), CLng(act("shape_id")), CStr(act("kind"))
        Case "clear_slide"
            Dim keep As Variant
            If act.Exists("keep_shape_ids") Then
                keep = act("keep_shape_ids")
            Else
                keep = Array()
            End If
            modActionsLayout.Do_clear_slide CLng(act("slide")), keep
        Case "move_shape_relative"
            modActionsLayout.Do_move_shape_relative CLng(act("slide")), CLng(act("shape_id")), _
                                                    CSng(act("dx_pt")), CSng(act("dy_pt"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_set_shape_kind + Do_clear_slide + Do_move_shape_relative"
```

---

## Section 5: Cross-cutting batch (1 task for 3 actions)

### Task 18: `Do_recolor_fill_match`, `Do_recolor_font_match`, `Do_delete_shapes_match`

**Files:**
- Modify: `src/modActionsLayout.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_cross_cutting():
    print("test_action_cross_cutting")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        # Slide 4 has a blue rectangle (#2E75B6) — recolor deck-wide to red
        app.Run("PPT_AI_Editor!Do_recolor_fill_match", "deck", "#2E75B6", "#FF0000")
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        rect = next(s for s in snap["slides"][3]["shapes"] if s.get("text") == "Plain")
        assert_eq(rect["fill"].upper(), "#FF0000", "rect fill after recolor_match")

        # delete_shapes_match by text_contains
        app.Run("PPT_AI_Editor!Do_delete_shapes_match", "slide:4", "", "", "Plain")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert all(s.get("text") != "Plain" for s in snap2["slides"][3]["shapes"]), "delete_shapes_match failed"
        print("  ok  [cross-cutting recolor + delete_shapes_match]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsLayout.bas`**

```vb
Public Sub Do_recolor_fill_match(scope As String, fromHex As String, toHex As String)
    ApplyByScope scope, "fill", fromHex, toHex
End Sub

Public Sub Do_recolor_font_match(scope As String, fromHex As String, toHex As String)
    ApplyByScope scope, "font", fromHex, toHex
End Sub

Public Sub Do_delete_shapes_match(scope As String, kindFilter As String, _
                                  fillFilter As String, textContains As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim slideFilter As Long: slideFilter = ParseScope(scope)
    Dim i As Long
    For i = 1 To pres.Slides.Count
        If slideFilter = 0 Or slideFilter = i Then
            Dim toDelete As New Collection
            Dim sh As Shape
            For Each sh In pres.Slides(i).Shapes
                If MatchesShape(sh, kindFilter, fillFilter, textContains) Then
                    toDelete.Add sh.Id
                End If
            Next sh
            Dim j As Long
            For j = 1 To toDelete.Count
                Dim victim As Shape: Set victim = modActions.FindShape(i, CLng(toDelete(j)))
                If Not victim Is Nothing Then victim.Delete
            Next j
        End If
    Next i
End Sub

Private Function ParseScope(scope As String) As Long
    If LCase(Left(scope, 6)) = "slide:" Then
        ParseScope = CLng(Mid(scope, 7))
    ElseIf LCase(scope) = "deck" Then
        ParseScope = 0
    Else
        Err.Raise vbObjectError + 4008, "ParseScope", "scope must be 'deck' or 'slide:N'"
    End If
End Function

Private Function MatchesShape(sh As Shape, kindFilter As String, _
                              fillFilter As String, textContains As String) As Boolean
    If Len(kindFilter) > 0 Then
        ' Compare against ResolveAutoShapeKind result if applicable
        On Error Resume Next
        Dim k As Long: k = ResolveAutoShapeKind(kindFilter)
        If sh.AutoShapeType <> k Then
            MatchesShape = False
            Err.Clear
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If Len(fillFilter) > 0 Then
        On Error Resume Next
        If sh.Fill.Type <> msoFillSolid Then
            MatchesShape = False
            Err.Clear
            Exit Function
        End If
        Dim hex As String: hex = modExportSnapshot_RgbToHex(sh.Fill.ForeColor.RGB)
        If LCase(hex) <> LCase(fillFilter) Then
            MatchesShape = False
            Err.Clear
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If
    If Len(textContains) > 0 Then
        Dim hasText As Boolean: hasText = False
        On Error Resume Next
        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then
                If InStr(sh.TextFrame.TextRange.Text, textContains) > 0 Then hasText = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
        If Not hasText Then
            MatchesShape = False
            Exit Function
        End If
    End If
    MatchesShape = True
End Function

Private Sub ApplyByScope(scope As String, propertyKind As String, _
                         fromHex As String, toHex As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim slideFilter As Long: slideFilter = ParseScope(scope)
    Dim i As Long
    For i = 1 To pres.Slides.Count
        If slideFilter = 0 Or slideFilter = i Then
            Dim sh As Shape
            For Each sh In pres.Slides(i).Shapes
                ApplyToShape sh, propertyKind, fromHex, toHex
            Next sh
        End If
    Next i
End Sub

Private Sub ApplyToShape(sh As Shape, propertyKind As String, _
                         fromHex As String, toHex As String)
    On Error Resume Next
    If propertyKind = "fill" Then
        If sh.Fill.Type = msoFillSolid Then
            Dim curHex As String: curHex = modExportSnapshot_RgbToHex(sh.Fill.ForeColor.RGB)
            If LCase(curHex) = LCase(fromHex) Then
                sh.Fill.ForeColor.RGB = modActions.HexToRgb(toHex)
            End If
        End If
    ElseIf propertyKind = "font" Then
        If sh.HasTextFrame Then
            If sh.TextFrame.HasText Then
                Dim fc As String: fc = modExportSnapshot_RgbToHex(sh.TextFrame.TextRange.Font.Color.RGB)
                If LCase(fc) = LCase(fromHex) Then
                    sh.TextFrame.TextRange.Font.Color.RGB = modActions.HexToRgb(toHex)
                End If
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' Wrapper around the Private RgbToHex in modExportSnapshot — expose it Public there if not already.
Private Function modExportSnapshot_RgbToHex(v As Long) As String
    modExportSnapshot_RgbToHex = modExportSnapshot.RgbToHex(v)
End Function
```

NOTE: `modExportSnapshot.RgbToHex` was made `Public` in V1 Task 8. If it is still `Private`, change it to `Public` as part of this task.

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "recolor_fill_match", "recolor_font_match"
            ValidateAction = RequireFields(act, Array("scope", "from", "to"))
        Case "delete_shapes_match"
            ValidateAction = RequireFields(act, Array("scope"))
```

`DispatchAction`:

```vb
        Case "recolor_fill_match"
            modActionsLayout.Do_recolor_fill_match CStr(act("scope")), CStr(act("from")), CStr(act("to"))
        Case "recolor_font_match"
            modActionsLayout.Do_recolor_font_match CStr(act("scope")), CStr(act("from")), CStr(act("to"))
        Case "delete_shapes_match"
            Dim kf As String, ff As String, tc As String
            kf = "" : ff = "" : tc = ""
            If act.Exists("kind") Then kf = CStr(act("kind"))
            If act.Exists("fill") Then ff = CStr(act("fill"))
            If act.Exists("text_contains") Then tc = CStr(act("text_contains"))
            modActionsLayout.Do_delete_shapes_match CStr(act("scope")), kf, ff, tc
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: cross-cutting recolor_*_match + delete_shapes_match"
```

---

## Section 6: Speaker notes (1 task for 2 actions)

### Task 19: `Do_set_speaker_notes` + `Do_append_speaker_notes`

**Files:**
- Modify: `src/modActions.bas` (append)
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_speaker_notes():
    print("test_action_speaker_notes")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        app.Run("PPT_AI_Editor!Do_set_speaker_notes", 2, "Slide 2 talking points")
        app.Run("PPT_AI_Editor!Do_append_speaker_notes", 2, "Additional context")
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        notes = snap["slides"][1]["speaker_notes"]
        assert "Slide 2 talking points" in notes, f"set: {notes!r}"
        assert "Additional context" in notes, f"append: {notes!r}"
        print("  ok  [speaker notes]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActions.bas`**

```vb
Public Sub Do_set_speaker_notes(slideNum As Long, value As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 5001, "Do_set_speaker_notes", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim ph As Object
    Dim i As Long
    For i = 1 To sl.NotesPage.Shapes.Placeholders.Count
        Set ph = sl.NotesPage.Shapes.Placeholders(i)
        If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
            ph.TextFrame.TextRange.Text = value
            Exit Sub
        End If
    Next i
End Sub

Public Sub Do_append_speaker_notes(slideNum As Long, value As String)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 5001, "Do_append_speaker_notes", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim ph As Object
    Dim i As Long
    For i = 1 To sl.NotesPage.Shapes.Placeholders.Count
        Set ph = sl.NotesPage.Shapes.Placeholders(i)
        If ph.PlaceholderFormat.Type = ppPlaceholderBody Then
            Dim cur As String: cur = ph.TextFrame.TextRange.Text
            If Len(cur) = 0 Then
                ph.TextFrame.TextRange.Text = value
            Else
                ph.TextFrame.TextRange.Text = cur & vbCrLf & value
            End If
            Exit Sub
        End If
    Next i
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_speaker_notes", "append_speaker_notes"
            ValidateAction = RequireFields(act, Array("slide", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
```

`DispatchAction`:

```vb
        Case "set_speaker_notes"
            modActions.Do_set_speaker_notes CLng(act("slide")), CStr(act("value"))
        Case "append_speaker_notes"
            modActions.Do_append_speaker_notes CLng(act("slide")), CStr(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_set_speaker_notes + Do_append_speaker_notes"
```

---

## Section 7: Images (1 task for 2 actions)

### Task 20: `modActionsImage` + `Do_insert_picture` + `Do_replace_picture`

**Files:**
- Create: `src/modActionsImage.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

A test image is needed. Use the existing `test_decks` directory and create a tiny PNG via the smoke harness.

- [ ] **Step 1: Append failing test**

```python
def test_action_images():
    print("test_action_images")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        # Drop a small image into tmpdir
        png_path = tmpdir / "tiny.png"
        png_path.write_bytes(bytes.fromhex(
            "89504E470D0A1A0A0000000D49484452000000010000000108060000001F15C489"
            "0000000A49444154789C636000000002000156A2C8B40000000049454E44AE426082"
        ))
        app.Run("PPT_AI_Editor!Do_insert_picture", 4, str(png_path), 50.0, 50.0, 100.0, 100.0)
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        pics = [s for s in snap["slides"][3]["shapes"] if s["type"] == "picture"]
        assert pics, "no picture inserted"
        pic_id = pics[0]["shape_id"]

        # Replace it
        app.Run("PPT_AI_Editor!Do_replace_picture", 4, pic_id, str(png_path))
        # Re-snapshot to confirm shape still exists
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        pics2 = [s for s in snap2["slides"][3]["shapes"] if s["type"] == "picture"]
        assert pics2, "picture missing after replace"
        print("  ok  [insert + replace picture]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsImage.bas`**

```vb
Attribute VB_Name = "modActionsImage"
Option Explicit

Public Sub Do_insert_picture(slideNum As Long, path As String, _
                             leftPt As Single, topPt As Single, _
                             widthPt As Single, heightPt As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 6001, "Do_insert_picture", "slide_out_of_range"
    End If
    If Not FileExists(path) Then
        Err.Raise vbObjectError + 6002, "Do_insert_picture", "file_not_found: " & path
    End If
    pres.Slides(slideNum).Shapes.AddPicture _
        FileName:=path, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=leftPt, Top:=topPt, _
        Width:=widthPt, Height:=heightPt
End Sub

Public Sub Do_replace_picture(slideNum As Long, shapeId As Long, path As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 6001, "Do_replace_picture", "shape not found"
    If sh.Type <> msoPicture Then Err.Raise vbObjectError + 6003, "Do_replace_picture", "shape is not a picture"
    If Not FileExists(path) Then Err.Raise vbObjectError + 6002, "Do_replace_picture", "file_not_found: " & path

    Dim L As Single, T As Single, W As Single, H As Single
    L = sh.Left: T = sh.Top: W = sh.Width: H = sh.Height
    sh.Delete
    ActivePresentation.Slides(slideNum).Shapes.AddPicture _
        FileName:=path, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
        Left:=L, Top:=T, Width:=W, Height:=H
End Sub

Private Function FileExists(p As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(p)
End Function
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "insert_picture"
            ValidateAction = RequireFields(act, Array("slide", "path", "pos"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "replace_picture"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "path"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "insert_picture"
            Dim ipos As Object: Set ipos = act("pos")
            modActionsImage.Do_insert_picture CLng(act("slide")), CStr(act("path")), _
                                              CSng(ipos("left")), CSng(ipos("top")), _
                                              CSng(ipos("width")), CSng(ipos("height"))
        Case "replace_picture"
            modActionsImage.Do_replace_picture CLng(act("slide")), CLng(act("shape_id")), CStr(act("path"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_insert_picture + Do_replace_picture (file path only)"
```

---

## Section 8: Slide structure (3 tasks for 3 actions)

### Task 21: `modActionsSlide` + `Do_move_slide`

**Files:**
- Create: `src/modActionsSlide.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_move_slide():
    print("test_action_move_slide")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        # Move slide 3 to position 1
        app.Run("PPT_AI_Editor!Do_move_slide", 3, 1)
        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert_eq(len(after["slides"]), len(before["slides"]), "slide count preserved")
        # Old slide 3 had three lettered boxes; should now be slide 1
        s1 = after["slides"][0]
        # If groupchildren — check direct text shapes
        any_box = any(sh.get("text") in ("A", "B", "C") for sh in s1["shapes"])
        assert any_box, f"slide 1 doesn't contain expected boxes: {s1['shapes']}"
        print("  ok  [move_slide]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsSlide.bas`**

```vb
Attribute VB_Name = "modActionsSlide"
Option Explicit

Public Sub Do_move_slide(fromIdx As Long, toIdx As Long)
    Dim pres As Presentation: Set pres = ActivePresentation
    If fromIdx < 1 Or fromIdx > pres.Slides.Count Then
        Err.Raise vbObjectError + 7001, "Do_move_slide", "from out of range"
    End If
    If toIdx < 1 Or toIdx > pres.Slides.Count Then
        Err.Raise vbObjectError + 7002, "Do_move_slide", "to out of range"
    End If
    pres.Slides(fromIdx).MoveTo toIdx
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "move_slide"
            ValidateAction = RequireFields(act, Array("from", "to"))
```

`DispatchAction`:

```vb
        Case "move_slide"
            modActionsSlide.Do_move_slide CLng(act("from")), CLng(act("to"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsSlide + Do_move_slide"
```

---

### Task 22: `Do_extract_slides`

**Files:**
- Modify: `src/modActionsSlide.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_extract_slides():
    print("test_action_extract_slides")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        out_path = str(tmpdir / "extracted.pptx")
        app.Run("PPT_AI_Editor!Do_extract_slides", [1, 3], out_path)
        assert Path(out_path).exists(), f"extracted file missing: {out_path}"
        # Open the extracted deck and verify slide count
        ext = app.Presentations.Open(out_path, WithWindow=False)
        try:
            assert ext.Slides.Count == 2, f"extracted slide count: {ext.Slides.Count}"
        finally:
            ext.Saved = True
            ext.Close()
        print("  ok  [extract_slides]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsSlide.bas`**

```vb
Public Sub Do_extract_slides(slideIndices As Variant, outputPath As String)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If cnt < 1 Then Err.Raise vbObjectError + 7003, "Do_extract_slides", "no slides specified"

    Dim src As Presentation: Set src = ActivePresentation
    Dim app As Object: Set app = Application
    Dim outPres As Presentation
    Set outPres = app.Presentations.Add(WithWindow:=msoFalse)
    Dim i As Long
    For i = 0 To cnt - 1
        Dim n As Long: n = ids(i)
        If n < 1 Or n > src.Slides.Count Then
            Err.Raise vbObjectError + 7004, "Do_extract_slides", "slide index out of range: " & n
        End If
        ' Copy + paste into out deck
        src.Slides(n).Copy
        outPres.Slides.Paste
    Next i

    ' Remove the default blank slide that Presentations.Add inserted
    On Error Resume Next
    Do While outPres.Slides.Count > cnt
        outPres.Slides(1).Delete
    Loop
    Err.Clear
    On Error GoTo 0

    outPres.SaveAs outputPath
    outPres.Close
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "extract_slides"
            ValidateAction = RequireFields(act, Array("slide_indices", "output_path"))
```

`DispatchAction`:

```vb
        Case "extract_slides"
            modActionsSlide.Do_extract_slides act("slide_indices"), CStr(act("output_path"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_extract_slides via copy/paste into new deck"
```

---

### Task 23: `Do_import_slides_from_deck`

**Files:**
- Modify: `src/modActionsSlide.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_import_slides():
    print("test_action_import_slides")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        # Use smoke_3slide.pptx as the source
        from shutil import copy2
        src_copy = tmpdir / "src.pptx"
        copy2(REPO_ROOT / "test_decks" / "smoke_3slide.pptx", src_copy)

        before = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        before_count = len(before["slides"])

        app.Run("PPT_AI_Editor!Do_import_slides_from_deck", str(src_copy), [1, 2], 2)

        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert_eq(len(after["slides"]), before_count + 2, "slide count after import")
        print("  ok  [import_slides_from_deck]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsSlide.bas`**

```vb
Public Sub Do_import_slides_from_deck(sourcePath As String, slideIndices As Variant, _
                                      targetPosition As Long)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(slideIndices, ids)
    If cnt < 1 Then Err.Raise vbObjectError + 7005, "Do_import_slides_from_deck", "no slide indices"
    If Not FileExists(sourcePath) Then
        Err.Raise vbObjectError + 7006, "Do_import_slides_from_deck", "source_not_found: " & sourcePath
    End If

    Dim pres As Presentation: Set pres = ActivePresentation
    If targetPosition < 1 Then targetPosition = 1
    If targetPosition > pres.Slides.Count + 1 Then targetPosition = pres.Slides.Count + 1

    ' Slides.InsertFromFile expects start/end as a contiguous range. If the
    ' caller passes non-contiguous ids, fall back to multiple calls.
    Dim i As Long
    Dim insertedSoFar As Long: insertedSoFar = 0
    For i = 0 To cnt - 1
        Dim startIdx As Long: startIdx = ids(i)
        Dim endIdx As Long: endIdx = startIdx
        ' Look ahead for contiguous run
        Do While i + 1 <= cnt - 1
            If ids(i + 1) = ids(i) + 1 Then
                i = i + 1
                endIdx = ids(i)
            Else
                Exit Do
            End If
        Loop
        pres.Slides.InsertFromFile sourcePath, _
                                   targetPosition - 1 + insertedSoFar, _
                                   startIdx, endIdx
        insertedSoFar = insertedSoFar + (endIdx - startIdx + 1)
    Next i
End Sub

Private Function FileExists(p As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(p)
End Function
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "import_slides_from_deck"
            ValidateAction = RequireFields(act, Array("source_path", "slide_indices", "target_position"))
```

`DispatchAction`:

```vb
        Case "import_slides_from_deck"
            modActionsSlide.Do_import_slides_from_deck CStr(act("source_path")), _
                                                       act("slide_indices"), _
                                                       CLng(act("target_position"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_import_slides_from_deck via Slides.InsertFromFile"
```

---

## Section 9: Tables (3 tasks for 5 actions)

### Task 24: `modActionsTable` + `Do_add_table_row` + `Do_delete_table_row`

**Files:**
- Create: `src/modActionsTable.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_table_row_ops():
    print("test_action_table_row_ops")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab = next(s for s in snap["slides"][3]["shapes"] if s["type"] == "table")
        sid = tab["shape_id"]
        before_rows = tab["table"]["rows"]

        app.Run("PPT_AI_Editor!Do_add_table_row", 4, sid, before_rows)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab2 = next(s for s in snap2["slides"][3]["shapes"] if s["shape_id"] == sid)
        assert_eq(tab2["table"]["rows"], before_rows + 1, "rows after add")

        app.Run("PPT_AI_Editor!Do_delete_table_row", 4, sid, 1)
        snap3 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab3 = next(s for s in snap3["slides"][3]["shapes"] if s["shape_id"] == sid)
        assert_eq(tab3["table"]["rows"], before_rows, "rows after delete")
        print("  ok  [table row ops]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsTable.bas`**

```vb
Attribute VB_Name = "modActionsTable"
Option Explicit

Public Sub Do_add_table_row(slideNum As Long, shapeId As Long, afterRow As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_add_table_row", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_add_table_row", "shape is not a table"
    Dim tbl As Table: Set tbl = sh.Table
    If afterRow < 0 Or afterRow > tbl.Rows.Count Then
        Err.Raise vbObjectError + 8003, "Do_add_table_row", "after_row out of range"
    End If
    If afterRow = 0 Then
        tbl.Rows(1).Select  ' position cursor; Add inserts before
        tbl.Rows.Add 1
    Else
        tbl.Rows.Add afterRow + 1
    End If
End Sub

Public Sub Do_delete_table_row(slideNum As Long, shapeId As Long, rowNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_delete_table_row", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_delete_table_row", "shape is not a table"
    Dim tbl As Table: Set tbl = sh.Table
    If rowNum < 1 Or rowNum > tbl.Rows.Count Then
        Err.Raise vbObjectError + 8004, "Do_delete_table_row", "row out of range"
    End If
    tbl.Rows(rowNum).Delete
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "add_table_row"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "after_row"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_table_row"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "add_table_row"
            modActionsTable.Do_add_table_row CLng(act("slide")), CLng(act("shape_id")), CLng(act("after_row"))
        Case "delete_table_row"
            modActionsTable.Do_delete_table_row CLng(act("slide")), CLng(act("shape_id")), CLng(act("row"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsTable + add/delete row"
```

---

### Task 25: `Do_add_table_col` + `Do_delete_table_col`

**Files:**
- Modify: `src/modActionsTable.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_table_col_ops():
    print("test_action_table_col_ops")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab = next(s for s in snap["slides"][3]["shapes"] if s["type"] == "table")
        sid = tab["shape_id"]
        before_cols = tab["table"]["cols"]

        app.Run("PPT_AI_Editor!Do_add_table_col", 4, sid, before_cols)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab2 = next(s for s in snap2["slides"][3]["shapes"] if s["shape_id"] == sid)
        assert_eq(tab2["table"]["cols"], before_cols + 1, "cols after add")

        app.Run("PPT_AI_Editor!Do_delete_table_col", 4, sid, 2)
        snap3 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab3 = next(s for s in snap3["slides"][3]["shapes"] if s["shape_id"] == sid)
        assert_eq(tab3["table"]["cols"], before_cols, "cols after delete")
        print("  ok  [table col ops]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsTable.bas`**

```vb
Public Sub Do_add_table_col(slideNum As Long, shapeId As Long, afterCol As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_add_table_col", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_add_table_col", "shape is not a table"
    Dim tbl As Table: Set tbl = sh.Table
    If afterCol < 0 Or afterCol > tbl.Columns.Count Then
        Err.Raise vbObjectError + 8005, "Do_add_table_col", "after_col out of range"
    End If
    If afterCol = 0 Then
        tbl.Columns.Add 1
    Else
        tbl.Columns.Add afterCol + 1
    End If
End Sub

Public Sub Do_delete_table_col(slideNum As Long, shapeId As Long, colNum As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_delete_table_col", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_delete_table_col", "shape is not a table"
    Dim tbl As Table: Set tbl = sh.Table
    If colNum < 1 Or colNum > tbl.Columns.Count Then
        Err.Raise vbObjectError + 8006, "Do_delete_table_col", "col out of range"
    End If
    tbl.Columns(colNum).Delete
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "add_table_col"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "after_col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "delete_table_col"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "col"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "add_table_col"
            modActionsTable.Do_add_table_col CLng(act("slide")), CLng(act("shape_id")), CLng(act("after_col"))
        Case "delete_table_col"
            modActionsTable.Do_delete_table_col CLng(act("slide")), CLng(act("shape_id")), CLng(act("col"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: add/delete table columns"
```

---

### Task 26: `Do_merge_cells`

**Files:**
- Modify: `src/modActionsTable.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_merge_cells():
    print("test_action_merge_cells")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab = next(s for s in snap["slides"][3]["shapes"] if s["type"] == "table")
        sid = tab["shape_id"]
        app.Run("PPT_AI_Editor!Do_merge_cells", 4, sid, 1, 1, 1, 3)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        tab2 = next(s for s in snap2["slides"][3]["shapes"] if s["shape_id"] == sid)
        merges = tab2["table_extra"]["merged_cells"]
        assert any(m["row"] == 1 and m["col"] == 1 and m["col_span"] == 3 for m in merges), \
            f"merge not reflected: {merges}"
        print("  ok  [merge_cells]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsTable.bas`**

```vb
Public Sub Do_merge_cells(slideNum As Long, shapeId As Long, _
                          rowA As Long, colA As Long, _
                          rowB As Long, colB As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 8001, "Do_merge_cells", "shape not found"
    If Not sh.HasTable Then Err.Raise vbObjectError + 8002, "Do_merge_cells", "shape is not a table"
    sh.Table.Cell(rowA, colA).Merge sh.Table.Cell(rowB, colB)
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "merge_cells"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "row_a", "col_a", "row_b", "col_b"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "merge_cells"
            modActionsTable.Do_merge_cells CLng(act("slide")), CLng(act("shape_id")), _
                                           CLng(act("row_a")), CLng(act("col_a")), _
                                           CLng(act("row_b")), CLng(act("col_b"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: Do_merge_cells"
```

---

## Section 10: Groups (1 task for 2 actions)

### Task 27: `modActionsGroup` + `Do_group_shapes` + `Do_ungroup`

**Files:**
- Create: `src/modActionsGroup.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_group_ungroup():
    print("test_action_group_ungroup")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes = [s for s in snap["slides"][2]["shapes"] if s.get("text") in ("A", "B", "C")]
        ids = [b["shape_id"] for b in boxes]
        app.Run("PPT_AI_Editor!Do_group_shapes", 3, ids)
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        groups = [s for s in snap2["slides"][2]["shapes"] if "group_children" in s]
        assert len(groups) >= 1, "no group after group_shapes"
        gid = groups[0]["shape_id"]

        app.Run("PPT_AI_Editor!Do_ungroup", 3, gid)
        snap3 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        groups3 = [s for s in snap3["slides"][2]["shapes"] if "group_children" in s]
        assert len(groups3) == 0, "group still present after ungroup"
        print("  ok  [group + ungroup]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsGroup.bas`**

```vb
Attribute VB_Name = "modActionsGroup"
Option Explicit

Public Sub Do_group_shapes(slideNum As Long, shapeIds As Variant)
    Dim ids() As Long
    Dim cnt As Long: cnt = modActionsLayout.NormalizeIdsArray(shapeIds, ids)
    If cnt < 2 Then Err.Raise vbObjectError + 9001, "Do_group_shapes", "need >=2 shapes"
    Dim pres As Presentation: Set pres = ActivePresentation
    Dim sl As Slide: Set sl = pres.Slides(slideNum)

    Dim names() As String
    ReDim names(0 To cnt - 1)
    Dim i As Long
    For i = 0 To cnt - 1
        Dim sh As Shape: Set sh = modActions.FindShape(slideNum, ids(i))
        If sh Is Nothing Then Err.Raise vbObjectError + 9002, "Do_group_shapes", "shape not found: " & ids(i)
        names(i) = sh.Name
    Next i
    sl.Shapes.Range(names).Group
End Sub

Public Sub Do_ungroup(slideNum As Long, shapeId As Long)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 9003, "Do_ungroup", "shape not found"
    If sh.Type <> msoGroup Then Err.Raise vbObjectError + 9004, "Do_ungroup", "shape is not a group"
    sh.Ungroup
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "group_shapes"
            ValidateAction = RequireFields(act, Array("slide", "shape_ids"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
        Case "ungroup"
            ValidateAction = RequireFields(act, Array("slide", "shape_id"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "group_shapes"
            modActionsGroup.Do_group_shapes CLng(act("slide")), act("shape_ids")
        Case "ungroup"
            modActionsGroup.Do_ungroup CLng(act("slide")), CLng(act("shape_id"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsGroup + group_shapes + ungroup"
```

---

## Section 11: Connectors (1 task for 1 action)

### Task 28: `modActionsConnector` + `Do_add_connector`

**Files:**
- Create: `src/modActionsConnector.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_add_connector():
    print("test_action_add_connector")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        boxes = [s for s in snap["slides"][2]["shapes"] if s.get("text") in ("A", "B", "C")]
        # Connect first to second
        a = boxes[0]["shape_id"]
        b = boxes[1]["shape_id"]
        before_count = len(snap["slides"][2]["shapes"])
        app.Run("PPT_AI_Editor!Do_add_connector", 3, a, b, "straight", "filled", "#000000", 1.5)
        after = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        assert len(after["slides"][2]["shapes"]) == before_count + 1, "expected new connector shape"
        print("  ok  [add_connector]")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsConnector.bas`**

```vb
Attribute VB_Name = "modActionsConnector"
Option Explicit

Public Sub Do_add_connector(slideNum As Long, fromId As Long, toId As Long, _
                            kind As String, arrowEnd As String, _
                            hexColor As String, weightPt As Single)
    Dim pres As Presentation: Set pres = ActivePresentation
    If slideNum < 1 Or slideNum > pres.Slides.Count Then
        Err.Raise vbObjectError + 10001, "Do_add_connector", "slide_out_of_range"
    End If
    Dim sl As Slide: Set sl = pres.Slides(slideNum)
    Dim shFrom As Shape, shTo As Shape
    Set shFrom = modActions.FindShape(slideNum, fromId)
    Set shTo = modActions.FindShape(slideNum, toId)
    If shFrom Is Nothing Or shTo Is Nothing Then
        Err.Raise vbObjectError + 10002, "Do_add_connector", "endpoint shape not found"
    End If

    Dim ctype As Long
    Select Case LCase(kind)
        Case "straight": ctype = 1   ' msoConnectorStraight
        Case "elbow":    ctype = 2   ' msoConnectorElbow
        Case "curved":   ctype = 3   ' msoConnectorCurve
        Case Else:       ctype = 1
    End Select

    Dim conn As Shape
    Set conn = sl.Shapes.AddConnector(ctype, 0, 0, 100, 100)
    conn.ConnectorFormat.BeginConnect shFrom, 1
    conn.ConnectorFormat.EndConnect shTo, 1
    conn.RerouteConnections

    conn.Line.ForeColor.RGB = modActions.HexToRgb(hexColor)
    conn.Line.Weight = weightPt

    Select Case LCase(arrowEnd)
        Case "filled":  conn.Line.EndArrowheadStyle = 5  ' msoArrowheadTriangle
        Case "open":    conn.Line.EndArrowheadStyle = 2  ' msoArrowheadOpen
        Case "none":    conn.Line.EndArrowheadStyle = 1  ' msoArrowheadNone
        Case Else:      conn.Line.EndArrowheadStyle = 5
    End Select
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "add_connector"
            ValidateAction = RequireFields(act, Array("slide", "from_shape_id", "to_shape_id", "kind"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateSlide(act)
```

`DispatchAction`:

```vb
        Case "add_connector"
            Dim ae As String, cc As String, cw As Single
            ae = "filled" : cc = "#000000" : cw = 1.0
            If act.Exists("arrow_end") Then ae = CStr(act("arrow_end"))
            If act.Exists("color") Then cc = CStr(act("color"))
            If act.Exists("weight_pt") Then cw = CSng(act("weight_pt"))
            modActionsConnector.Do_add_connector CLng(act("slide")), _
                                                 CLng(act("from_shape_id")), _
                                                 CLng(act("to_shape_id")), _
                                                 CStr(act("kind")), ae, cc, cw
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsConnector + Do_add_connector"
```

---

## Section 12: Native charts (3 tasks for 5 actions)

### Task 29: `modActionsChart` + `Do_set_chart_type`

**Files:**
- Create: `src/modActionsChart.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_set_chart_type():
    print("test_action_set_chart_type")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        ch = next(s for s in snap["slides"][1]["shapes"] if s["type"] == "chart")
        sid = ch["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_chart_type", 2, sid, "barClustered")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        ch2 = next(s for s in snap2["slides"][1]["shapes"] if s["shape_id"] == sid)
        assert_eq(ch2["chart"]["type"], "barClustered", "chart type after set")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Create `src/modActionsChart.bas`**

```vb
Attribute VB_Name = "modActionsChart"
Option Explicit

Public Sub Do_set_chart_type(slideNum As Long, shapeId As Long, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_type", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_type", "not_a_native_chart"
    sh.Chart.ChartType = ChartTypeFromName(value)
End Sub

Public Function ChartTypeFromName(name As String) As Long
    Select Case LCase(name)
        Case "columnclustered", "xlcolumnclustered": ChartTypeFromName = 51
        Case "columnstacked":                         ChartTypeFromName = 52
        Case "line", "xlline":                        ChartTypeFromName = 4
        Case "pie", "xlpie":                          ChartTypeFromName = 5
        Case "barclustered", "xlbarclustered":        ChartTypeFromName = 57
        Case "barstacked":                            ChartTypeFromName = 58
        Case "area", "xlarea":                        ChartTypeFromName = 1
        Case "scatter", "xlxyscatter":                ChartTypeFromName = -4169
        Case "doughnut":                              ChartTypeFromName = -4120
        Case Else:
            Err.Raise vbObjectError + 11003, "ChartTypeFromName", "unknown chart type: " & name
    End Select
End Function
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_chart_type"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "set_chart_type"
            modActionsChart.Do_set_chart_type CLng(act("slide")), CLng(act("shape_id")), CStr(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: modActionsChart + Do_set_chart_type"
```

---

### Task 30: `Do_set_chart_title` + `Do_set_chart_axis_title`

**Files:**
- Modify: `src/modActionsChart.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_chart_titles():
    print("test_action_chart_titles")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        ch = next(s for s in snap["slides"][1]["shapes"] if s["type"] == "chart")
        sid = ch["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_chart_title", 2, sid, "Quarterly Revenue", True)
        app.Run("PPT_AI_Editor!Do_set_chart_axis_title", 2, sid, "x", "Quarter")
        app.Run("PPT_AI_Editor!Do_set_chart_axis_title", 2, sid, "y", "Revenue ($M)")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        ch2 = next(s for s in snap2["slides"][1]["shapes"] if s["shape_id"] == sid)
        assert_eq(ch2["chart"]["title"], "Quarterly Revenue", "chart title")
        assert_eq(ch2["chart"]["axis_titles"]["x"], "Quarter", "x axis title")
        assert_eq(ch2["chart"]["axis_titles"]["y"], "Revenue ($M)", "y axis title")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsChart.bas`**

```vb
Public Sub Do_set_chart_title(slideNum As Long, shapeId As Long, _
                              value As String, enabled As Boolean)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_title", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_title", "not_a_native_chart"
    Dim ch As Chart: Set ch = sh.Chart
    If enabled Then
        ch.HasTitle = True
        ch.ChartTitle.Text = value
    Else
        ch.HasTitle = False
    End If
End Sub

Public Sub Do_set_chart_axis_title(slideNum As Long, shapeId As Long, _
                                   axis As String, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_axis_title", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_axis_title", "not_a_native_chart"
    Dim ch As Chart: Set ch = sh.Chart
    Dim axNum As Long
    Select Case LCase(axis)
        Case "x": axNum = 1   ' xlCategory
        Case "y": axNum = 2   ' xlValue
        Case Else: Err.Raise vbObjectError + 11004, "Do_set_chart_axis_title", "axis must be 'x' or 'y'"
    End Select
    ch.Axes(axNum).HasTitle = True
    ch.Axes(axNum).AxisTitle.Text = value
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_chart_title"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_chart_axis_title"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "axis", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "set_chart_title"
            Dim cte As Boolean: cte = True
            If act.Exists("enabled") Then cte = CBool(act("enabled"))
            modActionsChart.Do_set_chart_title CLng(act("slide")), CLng(act("shape_id")), _
                                               CStr(act("value")), cte
        Case "set_chart_axis_title"
            modActionsChart.Do_set_chart_axis_title CLng(act("slide")), CLng(act("shape_id")), _
                                                    CStr(act("axis")), CStr(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: chart title + axis title"
```

---

### Task 31: `Do_set_chart_legend_position` + `Do_set_series_color`

**Files:**
- Modify: `src/modActionsChart.bas`
- Modify: `src/modExecuteInstructions.bas`
- Modify: `tests/run_smoke.py`

- [ ] **Step 1: Append failing test**

```python
def test_action_chart_legend_and_series():
    print("test_action_chart_legend_and_series")
    app = open_app()
    deck, carrier, tmpdir = open_pair(app, "phase2.pptx")
    try:
        snap = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        ch = next(s for s in snap["slides"][1]["shapes"] if s["type"] == "chart")
        sid = ch["shape_id"]
        app.Run("PPT_AI_Editor!Do_set_chart_legend_position", 2, sid, "bottom")
        app.Run("PPT_AI_Editor!Do_set_series_color", 2, sid, 1, "#1F4E79")
        snap2 = json.loads(app.Run("PPT_AI_Editor!BuildSnapshotJson"))
        ch2 = next(s for s in snap2["slides"][1]["shapes"] if s["shape_id"] == sid)
        assert_eq(ch2["chart"]["legend_position"], "bottom", "legend position")
    finally:
        teardown(app, deck, carrier, tmpdir=tmpdir)
```

Add to `main()`.

- [ ] **Step 2: Run; expect failure**

- [ ] **Step 3: Append to `src/modActionsChart.bas`**

```vb
Public Sub Do_set_chart_legend_position(slideNum As Long, shapeId As Long, value As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_chart_legend_position", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_chart_legend_position", "not_a_native_chart"
    Dim ch As Chart: Set ch = sh.Chart
    If LCase(value) = "none" Then
        ch.HasLegend = False
        Exit Sub
    End If
    ch.HasLegend = True
    Select Case LCase(value)
        Case "left":   ch.Legend.Position = -4131
        Case "right":  ch.Legend.Position = -4152
        Case "top":    ch.Legend.Position = -4160
        Case "bottom": ch.Legend.Position = -4107
        Case "corner": ch.Legend.Position = 2
        Case Else: Err.Raise vbObjectError + 11005, "Do_set_chart_legend_position", "unknown position: " & value
    End Select
End Sub

Public Sub Do_set_series_color(slideNum As Long, shapeId As Long, _
                               seriesIndex As Long, hexValue As String)
    Dim sh As Shape: Set sh = modActions.FindShape(slideNum, shapeId)
    If sh Is Nothing Then Err.Raise vbObjectError + 11001, "Do_set_series_color", "shape not found"
    If Not sh.HasChart Then Err.Raise vbObjectError + 11002, "Do_set_series_color", "not_a_native_chart"
    Dim ch As Chart: Set ch = sh.Chart
    If seriesIndex < 1 Or seriesIndex > ch.SeriesCollection.Count Then
        Err.Raise vbObjectError + 11006, "Do_set_series_color", "series_index out of range"
    End If
    Dim ser As Object: Set ser = ch.SeriesCollection(seriesIndex)
    ser.Format.Fill.ForeColor.RGB = modActions.HexToRgb(hexValue)
End Sub
```

- [ ] **Step 4: Validation + dispatch**

`ValidateAction`:

```vb
        Case "set_chart_legend_position"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
        Case "set_series_color"
            ValidateAction = RequireFields(act, Array("slide", "shape_id", "series_index", "value"))
            If Len(ValidateAction) = 0 Then ValidateAction = ValidateShape(act)
```

`DispatchAction`:

```vb
        Case "set_chart_legend_position"
            modActionsChart.Do_set_chart_legend_position CLng(act("slide")), CLng(act("shape_id")), _
                                                          CStr(act("value"))
        Case "set_series_color"
            modActionsChart.Do_set_series_color CLng(act("slide")), CLng(act("shape_id")), _
                                                CLng(act("series_index")), CStr(act("value"))
```

- [ ] **Step 5: Sync + precheck + smoke**

If `Do_set_series_color` causes Excel to open visibly (check Task Manager during the smoke), demote it to V3 and remove from the spec/dispatch/snapshot. Document the demotion in the commit message.

- [ ] **Step 6: Commit**

```bash
git commit -am "feat: chart legend position + series color"
```

---

## Section 13: frmImportSlides UserForm + ImportSlides macro

### Task 32: `frmImportSlides` UserForm built via COM + `ImportSlides` macro

**Files:**
- Modify: `tools/build_forms.py` (extend to also build frmImportSlides)
- Modify: `src/modUI.bas` (add `Public Sub ImportSlides()`)
- Generated: `src/frmImportSlides.frm`, `src/frmImportSlides.frx`

- [ ] **Step 1: Extend `tools/build_forms.py`**

Below `build_frm_execute(...)` add a `build_frm_import_slides(components)` function. Mirror the existing form-building pattern. Form size: width 480, height 320. Controls:

| Type | Name | Properties |
|---|---|---|
| Label | `lblPath` | Caption "Source deck", Top 12, Left 12, Width 100, Height 20 |
| TextBox | `txtPath` | Top 12, Left 120, Width 280, Height 22, Locked True |
| CommandButton | `btnBrowse` | Caption "Browse...", Top 12, Left 408, Width 60, Height 22 |
| Label | `lblRange` | Caption "Slide range (e.g. 1-3,5,7-9)", Top 50, Left 12, Width 240, Height 20 |
| TextBox | `txtRange` | Top 50, Left 256, Width 212, Height 22 |
| Label | `lblPosition` | Caption "Insert at position", Top 90, Left 12, Width 240, Height 20 |
| TextBox | `txtPosition` | Top 90, Left 256, Width 60, Height 22 |
| CommandButton | `btnImport` | Caption "Import", Top 250, Left 12, Width 80, Height 28, Enabled False |
| CommandButton | `btnCancel` | Caption "Cancel", Top 250, Left 100, Width 80, Height 28 |
| Label | `lblStatus` | Caption "", Top 130, Left 12, Width 456, Height 100 |

Form code:

```vb
Option Explicit

Private Sub UserForm_Initialize()
    txtPath.Text = ""
    txtRange.Text = ""
    txtPosition.Text = "1"
    btnImport.Enabled = False
    lblStatus.Caption = ""
End Sub

Private Sub btnBrowse_Click()
    Dim path As String
    On Error Resume Next
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker = 3
    If Not fd Is Nothing Then
        fd.Filters.Clear
        fd.Filters.Add "PowerPoint Files", "*.pptx; *.pptm"
        If fd.Show = -1 Then
            path = fd.SelectedItems(1)
        End If
    End If
    On Error GoTo 0

    If Len(path) = 0 Then
        path = InputBox("Path to source deck:", "Source deck")
    End If

    If Len(path) > 0 Then
        txtPath.Text = path
        UpdateImportButton
    End If
End Sub

Private Sub txtRange_Change()
    UpdateImportButton
End Sub

Private Sub txtPosition_Change()
    UpdateImportButton
End Sub

Private Sub UpdateImportButton()
    btnImport.Enabled = (Len(txtPath.Text) > 0 And Len(txtRange.Text) > 0 And Len(txtPosition.Text) > 0)
End Sub

Private Sub btnImport_Click()
    On Error GoTo Failure
    Dim ids As Variant
    ids = ParseRange(txtRange.Text)
    Dim pos As Long: pos = CLng(txtPosition.Text)

    Dim before As Long: before = ActivePresentation.Slides.Count
    modActionsSlide.Do_import_slides_from_deck txtPath.Text, ids, pos
    Dim after As Long: after = ActivePresentation.Slides.Count
    lblStatus.Caption = "Imported " & (after - before) & " slide(s) at position " & pos
    Exit Sub
Failure:
    lblStatus.Caption = "ERROR: " & Err.Description
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

' Parse "1-3,5,7-9" → array of integers
Private Function ParseRange(s As String) As Variant
    Dim parts() As String
    parts = Split(s, ",")
    Dim col As New Collection
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim p As String: p = Trim(parts(i))
        If InStr(p, "-") > 0 Then
            Dim ab() As String: ab = Split(p, "-")
            Dim a As Long: a = CLng(Trim(ab(0)))
            Dim b As Long: b = CLng(Trim(ab(1)))
            Dim k As Long
            For k = a To b
                col.Add k
            Next k
        Else
            col.Add CLng(p)
        End If
    Next i
    Dim arr() As Long
    ReDim arr(0 To col.Count - 1)
    Dim j As Long
    For j = 1 To col.Count
        arr(j - 1) = col(j)
    Next j
    ParseRange = arr
End Function
```

In `tools/build_forms.py`'s `main()`, after the `build_frm_execute(...)` call, add:

```python
build_frm_import_slides(components)
```

And export the .frm:

```python
import_slides_comp = components("frmImportSlides")
import_slides_comp.Export(str(SRC_DIR / "frmImportSlides.frm"))
```

- [ ] **Step 2: Run `tools/build_forms.py`**

```bash
powershell -Command "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"
python tools/build_forms.py
```

Expected: `[add] frmImportSlides`, no errors. Verify `src/frmImportSlides.frm` and `src/frmImportSlides.frx` exist.

- [ ] **Step 3: Add `ImportSlides` macro to `src/modUI.bas`**

Append:

```vb
Public Sub ImportSlides()
    frmImportSlides.Show vbModeless
End Sub
```

- [ ] **Step 4: Sync + precheck**

```bash
python update_macros.py
python tools/precheck_carrier.py
```

- [ ] **Step 5: Smoke**

```bash
python tests/run_smoke.py
```

All existing tests pass. (No automated test for the form — that is verified manually by Alt+F8 → ImportSlides.)

- [ ] **Step 6: Commit**

```bash
git add tools/build_forms.py src/modUI.bas src/frmImportSlides.frm src/frmImportSlides.frx PPT_AI_Editor.pptm
git commit -m "feat: frmImportSlides UserForm + ImportSlides macro"
```

---

## Section 14: Final integration

### Task 33: Update prompt template in `src/frmExport.frm` (and `tools/build_forms.py`)

**Files:**
- Modify: `src/frmExport.frm`
- Modify: `tools/build_forms.py`

The current `PromptTemplate()` function in `frmExport.frm` lists 15 schemas. Replace it with the full Phase 2 catalogue (~70 schemas). Use the section structure from spec §6 and use `s = s & ...` per line — NEVER `& _` line continuation.

Build the new template once in `tools/build_forms.py` so the form-rebuild script and the on-disk `.frm` stay in sync.

- [ ] **Step 1: Replace `PromptTemplate()` in BOTH `src/frmExport.frm` and `tools/build_forms.py`**

Open both files and replace the existing `Private Function PromptTemplate() As String ... End Function` with this version. Same body in both places:

```vb
Private Function PromptTemplate() As String
    Dim s As String

    s = "You are editing a PowerPoint presentation. Below is the current state as JSON:" & vbCrLf & vbCrLf
    s = s & "```json" & vbCrLf & "{snapshot}" & vbCrLf & "```" & vbCrLf & vbCrLf
    s = s & "I want the following changes:" & vbCrLf & vbCrLf
    s = s & "[REPLACE THIS LINE WITH YOUR REQUEST]" & vbCrLf & vbCrLf

    s = s & "Return ONLY a valid instructions JSON. No prose, no explanation, no markdown" & vbCrLf
    s = s & "code fences. Top-level shape:" & vbCrLf & vbCrLf
    s = s & "{""actions"": [ <action>, <action>, ... ]}" & vbCrLf & vbCrLf

    s = s & "Each action is one of EXACTLY these schemas. Field names are STRICT - do not" & vbCrLf
    s = s & "rename ""value"" to ""text""/""color""/""size""/""fill"". Use names verbatim." & vbCrLf & vbCrLf

    s = s & "ATOMIC OPS (V1):" & vbCrLf
    s = s & "  {""type"":""set_text"",""slide"":1,""shape_id"":3,""value"":""Hello""}" & vbCrLf
    s = s & "  {""type"":""set_font_size"",""slide"":1,""shape_id"":3,""value"":28}" & vbCrLf
    s = s & "  {""type"":""set_font_bold"",""slide"":1,""shape_id"":3,""value"":true}" & vbCrLf
    s = s & "  {""type"":""set_font_italic"",""slide"":1,""shape_id"":3,""value"":false}" & vbCrLf
    s = s & "  {""type"":""set_font_color"",""slide"":1,""shape_id"":3,""value"":""#FF0000""}" & vbCrLf
    s = s & "  {""type"":""set_fill_color"",""slide"":1,""shape_id"":4,""value"":""#2E75B6""}" & vbCrLf
    s = s & "  {""type"":""move_shape"",""slide"":1,""shape_id"":4,""left"":100,""top"":200}" & vbCrLf
    s = s & "  {""type"":""resize_shape"",""slide"":1,""shape_id"":4,""width"":250,""height"":80}" & vbCrLf
    s = s & "  {""type"":""delete_shape"",""slide"":1,""shape_id"":7}" & vbCrLf
    s = s & "  {""type"":""add_slide"",""position"":3,""layout_index"":1}" & vbCrLf
    s = s & "  {""type"":""delete_slide"",""slide"":4}" & vbCrLf
    s = s & "  {""type"":""duplicate_slide"",""slide"":2}" & vbCrLf
    s = s & "  {""type"":""set_cell_text"",""slide"":1,""shape_id"":5,""row"":2,""col"":1,""value"":""Revenue""}" & vbCrLf
    s = s & "  {""type"":""swap_table_columns"",""slide"":1,""shape_id"":5,""col_a"":1,""col_b"":2}" & vbCrLf
    s = s & "  {""type"":""swap_table_rows"",""slide"":1,""shape_id"":5,""row_a"":1,""row_b"":2}" & vbCrLf & vbCrLf

    s = s & "GRANULAR TEXT:" & vbCrLf
    s = s & "  {""type"":""set_paragraph_text"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""...""}" & vbCrLf
    s = s & "  {""type"":""add_paragraph"",""slide"":1,""shape_id"":3,""after_paragraph_index"":-1,""value"":""...""}" & vbCrLf
    s = s & "  {""type"":""delete_paragraph"",""slide"":1,""shape_id"":3,""paragraph_index"":2}" & vbCrLf
    s = s & "  {""type"":""set_bullet_style"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""disc""}" & vbCrLf
    s = s & "  {""type"":""set_indent_level"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":1}" & vbCrLf
    s = s & "  {""type"":""set_paragraph_font_size"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":24}" & vbCrLf
    s = s & "  {""type"":""set_paragraph_font_color"",""slide"":1,""shape_id"":3,""paragraph_index"":0,""value"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""find_replace_text"",""scope"":""deck"",""find"":""ACME"",""replace"":""NewCo""}" & vbCrLf & vbCrLf

    s = s & "LAYOUT / COMPOSITION:" & vbCrLf
    s = s & "  {""type"":""align_shapes"",""slide"":1,""shape_ids"":[3,4,5],""anchor"":""top""}" & vbCrLf
    s = s & "  {""type"":""distribute_horizontal"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""distribute_vertical"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""tile_grid"",""slide"":1,""shape_ids"":[3,4,5,6],""cols"":2,""gap_pt"":10}" & vbCrLf
    s = s & "  {""type"":""fit_to_slide_margins"",""slide"":1,""shape_id"":3,""margin_pt"":36}" & vbCrLf
    s = s & "  {""type"":""add_line"",""slide"":1,""x1"":50,""y1"":50,""x2"":700,""y2"":50,""color"":""#000000"",""weight_pt"":1.5}" & vbCrLf
    s = s & "  {""type"":""add_shape"",""slide"":1,""kind"":""capsule"",""pos"":{""left"":100,""top"":100,""width"":200,""height"":60},""fill"":""#1F4E79"",""stroke"":null,""stroke_weight_pt"":0}" & vbCrLf
    s = s & "  {""type"":""set_shape_kind"",""slide"":1,""shape_id"":4,""kind"":""capsule""}" & vbCrLf
    s = s & "  {""type"":""clear_slide"",""slide"":1,""keep_shape_ids"":[3]}" & vbCrLf
    s = s & "  {""type"":""move_shape_relative"",""slide"":1,""shape_id"":4,""dx_pt"":0,""dy_pt"":50}" & vbCrLf & vbCrLf

    s = s & "CROSS-CUTTING BATCH:" & vbCrLf
    s = s & "  {""type"":""recolor_fill_match"",""scope"":""deck"",""from"":""#0000FF"",""to"":""#FF0000""}" & vbCrLf
    s = s & "  {""type"":""recolor_font_match"",""scope"":""slide:2"",""from"":""#000000"",""to"":""#1F4E79""}" & vbCrLf
    s = s & "  {""type"":""delete_shapes_match"",""scope"":""deck"",""text_contains"":""Confidential""}" & vbCrLf & vbCrLf

    s = s & "SPEAKER NOTES:" & vbCrLf
    s = s & "  {""type"":""set_speaker_notes"",""slide"":1,""value"":""Talking points...""}" & vbCrLf
    s = s & "  {""type"":""append_speaker_notes"",""slide"":1,""value"":""Add this too""}" & vbCrLf & vbCrLf

    s = s & "IMAGES (LOCAL FILE PATHS ONLY — no URLs):" & vbCrLf
    s = s & "  {""type"":""insert_picture"",""slide"":1,""path"":""C:\\\\path\\\\to\\\\img.png"",""pos"":{""left"":50,""top"":50,""width"":200,""height"":150}}" & vbCrLf
    s = s & "  {""type"":""replace_picture"",""slide"":1,""shape_id"":7,""path"":""C:\\\\path\\\\to\\\\new.png""}" & vbCrLf & vbCrLf

    s = s & "SLIDE STRUCTURE:" & vbCrLf
    s = s & "  {""type"":""move_slide"",""from"":3,""to"":1}" & vbCrLf
    s = s & "  {""type"":""extract_slides"",""slide_indices"":[1,3,5],""output_path"":""C:\\\\path\\\\out.pptx""}" & vbCrLf
    s = s & "  {""type"":""import_slides_from_deck"",""source_path"":""C:\\\\path\\\\other.pptx"",""slide_indices"":[1,2],""target_position"":3}" & vbCrLf & vbCrLf

    s = s & "TABLES:" & vbCrLf
    s = s & "  {""type"":""add_table_row"",""slide"":1,""shape_id"":5,""after_row"":2}" & vbCrLf
    s = s & "  {""type"":""delete_table_row"",""slide"":1,""shape_id"":5,""row"":3}" & vbCrLf
    s = s & "  {""type"":""add_table_col"",""slide"":1,""shape_id"":5,""after_col"":2}" & vbCrLf
    s = s & "  {""type"":""delete_table_col"",""slide"":1,""shape_id"":5,""col"":3}" & vbCrLf
    s = s & "  {""type"":""merge_cells"",""slide"":1,""shape_id"":5,""row_a"":1,""col_a"":1,""row_b"":1,""col_b"":3}" & vbCrLf & vbCrLf

    s = s & "GROUPS:" & vbCrLf
    s = s & "  {""type"":""group_shapes"",""slide"":1,""shape_ids"":[3,4,5]}" & vbCrLf
    s = s & "  {""type"":""ungroup"",""slide"":1,""shape_id"":12}" & vbCrLf & vbCrLf

    s = s & "CONNECTORS:" & vbCrLf
    s = s & "  {""type"":""add_connector"",""slide"":1,""from_shape_id"":3,""to_shape_id"":7,""kind"":""elbow"",""arrow_end"":""filled"",""color"":""#000000"",""weight_pt"":1.5}" & vbCrLf & vbCrLf

    s = s & "NATIVE CHARTS (Shape.HasChart=True only; pasted images skipped):" & vbCrLf
    s = s & "  {""type"":""set_chart_type"",""slide"":1,""shape_id"":4,""value"":""barClustered""}" & vbCrLf
    s = s & "  {""type"":""set_chart_title"",""slide"":1,""shape_id"":4,""value"":""Q3 Revenue"",""enabled"":true}" & vbCrLf
    s = s & "  {""type"":""set_chart_axis_title"",""slide"":1,""shape_id"":4,""axis"":""x"",""value"":""Quarter""}" & vbCrLf
    s = s & "  {""type"":""set_chart_legend_position"",""slide"":1,""shape_id"":4,""value"":""bottom""}" & vbCrLf
    s = s & "  {""type"":""set_series_color"",""slide"":1,""shape_id"":4,""series_index"":1,""value"":""#1F4E79""}" & vbCrLf & vbCrLf

    s = s & "RULES:" & vbCrLf
    s = s & "- Use only shape_ids that exist in the snapshot. Do not invent ids." & vbCrLf
    s = s & "- Slide / row / col / series / position numbers are 1-based." & vbCrLf
    s = s & "- paragraph_index is 0-based to match the snapshot's paragraphs[].index." & vbCrLf
    s = s & "- after_paragraph_index = -1 means insert at top." & vbCrLf
    s = s & "- Colors are #RRGGBB hex strings." & vbCrLf
    s = s & "- Lengths in points." & vbCrLf
    s = s & "- Booleans are JSON true / false (lowercase, no quotes)." & vbCrLf
    s = s & "- For 'every X with property P -> Y' requests: enumerate matching shape_ids" & vbCrLf
    s = s & "  in the snapshot and emit one action per match (or use a *_match helper)." & vbCrLf
    s = s & "- For 'rebuild this slide': use clear_slide first, then a sequence of" & vbCrLf
    s = s & "  add_shape / set_text / move_shape actions to populate." & vbCrLf
    s = s & "- File paths must be absolute and use double backslashes in JSON strings." & vbCrLf
    s = s & "- One field name per action - never substitute aliases."

    PromptTemplate = s
End Function
```

- [ ] **Step 2: Sync + precheck + smoke**

```bash
powershell -Command "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"
python update_macros.py
python tools/precheck_carrier.py
python tests/run_smoke.py
```

- [ ] **Step 3: Manual sanity-check the template**

In an interactive Python shell:

```python
import win32com.client
app = win32com.client.DispatchEx("PowerPoint.Application")
carrier = app.Presentations.Open(r"C:\Users\vinit\Documents\PPT_AI_Editor\PPT_AI_Editor.pptm", WithWindow=True)
deck = app.Presentations.Open(r"C:\Users\vinit\Documents\PPT_AI_Editor\test_decks\full_visual.pptx", WithWindow=True)
deck.Windows(1).Activate()
# Force-open frmExport, fetch text from txtSnapshot via the form (not directly possible — manual check)
```

Or just do a manual smoke: open the carrier, run `ExportSnapshot`, click "Copy snapshot + prompt template", paste into a text editor, verify all 11 buckets are visible in the template.

- [ ] **Step 4: Commit**

```bash
git add src/frmExport.frm tools/build_forms.py PPT_AI_Editor.pptm
git commit -m "docs(prompt): full Phase 2 prompt template (70 schemas, 11 buckets)"
```

---

### Task 34: Final full-loop manual smoke + carrier commit + tag v2.0

**Files:**
- (no source changes; integration verification only)

- [ ] **Step 1: Run full automated suite**

```bash
powershell -Command "Stop-Process -Name POWERPNT -Force -ErrorAction SilentlyContinue"
python tests/make_test_decks.py
python update_macros.py
python tools/precheck_carrier.py
python tests/run_smoke.py
```

Expected: all ~40 tests pass, "all tests passed" line at end.

- [ ] **Step 2: Verify carrier components**

```python
python -c "
import win32com.client
app = win32com.client.DispatchEx('PowerPoint.Application')
try:
    pres = app.Presentations.Open(r'C:\\Users\\vinit\\Documents\\PPT_AI_Editor\\PPT_AI_Editor.pptm', WithWindow=True)
    expected = {
        'modJSON': 1,
        'modBackup': 1,
        'modExportSnapshot': 1,
        'modActions': 1,
        'modActionsText': 1,
        'modActionsLayout': 1,
        'modActionsTable': 1,
        'modActionsChart': 1,
        'modActionsImage': 1,
        'modActionsSlide': 1,
        'modActionsGroup': 1,
        'modActionsConnector': 1,
        'modExecuteInstructions': 1,
        'modUI': 1,
        'frmExport': 3,
        'frmExecute': 3,
        'frmImportSlides': 3,
    }
    found = {c.Name: c.Type for c in pres.VBProject.VBComponents}
    missing = [(n, t) for n, t in expected.items() if found.get(n) != t]
    extras = [n for n in found if n not in expected]
    if missing:
        print('MISSING:', missing)
        raise SystemExit(1)
    if extras:
        print('EXTRAS:', extras)
    print('OK: 17 components present with correct types')
finally:
    pres.Saved = True
    pres.Close()
    app.Quit()
"
```

Expected: `OK: 17 components present with correct types`.

- [ ] **Step 3: Commit any pending carrier changes**

```bash
git add PPT_AI_Editor.pptm
git diff --cached --stat
```

If anything changed (it should, given prompt template + frmImportSlides + new modules):

```bash
git commit -m "chore: commit Phase 2 carrier .pptm with 17 components"
```

- [ ] **Step 4: Tag v2.0**

```bash
git tag -a v2.0 -m "PPT AI Editor V2: 70 actions, snapshot v2, frmImportSlides, native chart ops"
```

- [ ] **Step 5: Update memory**

Add a one-line entry to `C:\Users\vinit\.claude\projects\C--Users-vinit\memory\ppt-ai-editor.md` summarizing Phase 2 completion (run-time output as part of the final report; do not commit it from the project repo).

---

## Self-review

**1. Spec coverage:**

| Spec section | Implemented in task |
|---|---|
| §3.1 occupied_rects | Task 2 |
| §3.1 speaker_notes | Task 3 |
| §3.2 paragraphs[] / runs[] | Task 4 |
| §3.2 group_children | Task 5 |
| §3.2 chart{} (native) | Task 6 |
| §3.2 table_extra | Task 7 |
| §4.A 8 granular text actions | Tasks 8-12 |
| §4.B 10 layout actions | Tasks 13-17 |
| §4.C 3 cross-cutting actions | Task 18 |
| §4.D 2 speaker-notes actions | Task 19 |
| §4.E 2 image actions | Task 20 |
| §4.F 3 slide-structure actions | Tasks 21-23 |
| §4.G 5 table actions | Tasks 24-26 |
| §4.H 2 group actions | Task 27 |
| §4.I 1 connector action | Task 28 |
| §4.K 5 chart actions | Tasks 29-31 |
| §5 frmImportSlides | Task 32 |
| §6 prompt template | Task 33 |
| §8 smoke plan | Tasks 2-31 (each adds tests) |
| Compile-error guardrail (per pre-spec request) | Task 0 |

**2. Placeholder scan:** No "TBD" / "TODO" / hand-wavy steps. Every code step has the actual code.

**3. Type / signature consistency:**
- `modActions.HexToRgb`, `modActions.FindShape`, `modExportSnapshot.RgbToHex` — referenced in multiple tasks, signatures stable. (Note: Task 18 requires `modExportSnapshot.RgbToHex` to be `Public`; if it is still `Private`, change to `Public` in Task 18.)
- `modActionsLayout.NormalizeIdsArray` — defined in Task 13, used in Tasks 17, 22, 23, 27.
- `modActionsLayout.ResolveAutoShapeKind` — defined in Task 16, used in Task 17 (`Do_set_shape_kind`).
- `BuildSnapshotJson` always has the same signature; new keys are additive.
- `ExecuteFromString` unchanged — only `ValidateAction` and `DispatchAction` grow.
- `precheck_carrier.py` script signature: zero-arg, exits 0 on success.

All consistent.

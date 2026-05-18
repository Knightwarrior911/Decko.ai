"""Prove P1-P5 against the built citi_final.pptx (structural, not pixel)."""
import os, time, json
import win32com.client as w
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent.parent
DECK = ROOT / "tests" / "citi_repro" / "citi_final.pptx"
CARRIER = ROOT / "PPT_AI_Editor.pptm"

os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1"); time.sleep(1.2)
app = w.DispatchEx("PowerPoint.Application"); app.Visible = True
try: app.AutomationSecurity = 1
except Exception: pass

# P1: whole-shape set_font_size on a TABLE must succeed (no "no text frame")
carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
pres = app.Presentations.Open(str(DECK), WithWindow=True)
s1 = pres.Slides(1)
tbl_sh = None
for sh in s1.Shapes:
    if sh.HasTable: tbl_sh = sh; break
log = Path(str(DECK) + ".verify_log.jsonl")
if log.exists(): log.unlink()
def run(p):
    try: return app.Run("ExecuteFromString", p)
    except Exception: return app.Run("PPT_AI_Editor.pptm!ExecuteFromString", p)
before = tbl_sh.Table.Cell(3, 1).Shape.TextFrame.TextRange.Font.Size
act = {"actions": [{"type": "set_font_size", "slide": 1, "shape_id": int(tbl_sh.Id), "value": 7}]}
pres.Windows(1).Activate()
run(json.dumps(act))
after = tbl_sh.Table.Cell(3, 1).Shape.TextFrame.TextRange.Font.Size
after2 = tbl_sh.Table.Cell(10, 3).Shape.TextFrame.TextRange.Font.Size
print(f"P1 set_font_size on TABLE: cell(3,1) {before}->{after}, cell(10,3)->{after2}  "
      f"=> {'PASS (applied to all cells, no error)' if after == 7 and after2 == 7 else 'FAIL'}")

# slide 2 chart probes
s2 = pres.Slides(2)
ch = None
for sh in s2.Shapes:
    if sh.HasChart: ch = sh.Chart; break
nser = ch.SeriesCollection().Count
names = [ch.SeriesCollection(i).Name for i in range(1, nser + 1)]
print(f"slide2 chart: {nser} series -> {names}")

# P2: legend must NOT contain a 'Total' entry
leg_n = ch.Legend.LegendEntries().Count if ch.HasLegend else 0
# legend entries have no .Text in all builds; infer by count vs series (Total excluded => leg_n < nser)
print(f"P2 legend entries={leg_n}, series={nser}  => "
      f"{'PASS (Total excluded from legend)' if leg_n < nser else 'CHECK (legend has all series)'}")

# P3: Goodwill series (the one with zeros) must have no label on zero points
gi = next((i for i, n in enumerate(names, 1) if 'Goodwill' in n), None)
if gi:
    vals = ch.SeriesCollection(gi).Values
    zero_pts = [p for p in range(1, len(vals) + 1) if float(vals[p - 1]) == 0]
    labeled = []
    for p in zero_pts:
        try:
            if ch.SeriesCollection(gi).Points(p).HasDataLabel: labeled.append(p)
        except Exception: pass
    print(f"P3 Goodwill zero-points={zero_pts}, still-labeled={labeled}  => "
          f"{'PASS (no labels on zeros)' if not labeled else 'FAIL'}")

# P4: combo present -> a line series on secondary axis + an auto 'Total' series
types = {ch.SeriesCollection(i).Name: int(ch.SeriesCollection(i).ChartType) for i in range(1, nser + 1)}
axg = {ch.SeriesCollection(i).Name: ch.SeriesCollection(i).AxisGroup for i in range(1, nser + 1)}
eff_line = types.get("Reported Efficiency Ratio") == 4 and axg.get("Reported Efficiency Ratio") == 2
has_total = "Total" in names
print(f"P4 combo: efficiency line@secondary={eff_line}, auto Total series={has_total}  => "
      f"{'PASS' if eff_line and has_total else 'CHECK'}  types={types} axisgroup={axg}")

# P5: real bullets — slide1 'hl' textbox + a slide2 card must have real
# ParagraphFormat.Bullet (Type<>0) + IndentLevel, and NOT literal bullet glyphs.
def bullet_report(slide, want_name_contains):
    for sh in slide.Shapes:
        if not sh.HasTextFrame: continue
        tr = sh.TextFrame.TextRange
        txt = tr.Text
        if want_name_contains and want_name_contains not in txt: continue
        if "Up 14%" in txt or "Higher volumes" in txt or "Higher severance" in txt:
            paras = tr.Paragraphs()
            out = []
            for i in range(1, paras.Count + 1):
                pf = tr.Paragraphs(i).ParagraphFormat
                bt = pf.Bullet.Type
                il = tr.Paragraphs(i).IndentLevel
                first = tr.Paragraphs(i).Text[:1]
                out.append((i, bt, il, first))
            return txt[:40], out
    return None, None

t1, r1 = bullet_report(s1, None)
print(f"P5 slide1 highlights text starts {t1!r}")
glyphbad = False
if r1:
    for i, bt, il, first in r1:
        if first in ("•", "-", "▪", "–"): glyphbad = True
        print(f"   para{i}: Bullet.Type={bt} IndentLevel={il} firstChar={first!r}")
    real = all(bt != 0 for _, bt, _, _ in r1) and not glyphbad
    print(f"P5 slide1 => {'PASS (real bullets, no glyph chars in text)' if real else 'FAIL'}")
t2, r2 = bullet_report(s2, None)
if r2:
    glyphbad2 = any(f in ("•", "-", "▪", "–") for _, _, _, f in r2)
    real2 = all(bt != 0 for _, bt, _, _ in r2) and not glyphbad2
    print(f"P5 slide2 card => {'PASS (real bullets)' if real2 else 'FAIL'}  {r2}")

pres.Saved = True
carrier.Saved = True
app.Quit()

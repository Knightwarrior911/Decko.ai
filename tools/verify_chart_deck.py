"""Open decko_chart_test.pptx and verify each slide holds a real chart object."""
from pathlib import Path
import win32com.client as w

DECK = Path.home() / "Downloads" / "decko_chart_test.pptx"

# msoChart = 3
MSO_CHART = 3

app = w.DispatchEx("PowerPoint.Application")
app.Visible = True
try:
    app.AutomationSecurity = 1
except Exception:
    pass
pres = app.Presentations.Open(str(DECK), WithWindow=True, ReadOnly=True)
ok = bad = 0
for i in range(1, pres.Slides.Count + 1):
    sl = pres.Slides(i)
    charts = []
    for sh in sl.Shapes:
        try:
            if sh.HasChart:
                charts.append(sh)
        except Exception:
            pass
    if i == 1 and not charts:
        continue  # initial blank slide
    if not charts:
        print(f"slide {i:2d}: NO CHART  (shapes={sl.Shapes.Count})")
        bad += 1
        continue
    ch = charts[0].Chart
    try:
        ctype = ch.ChartType
    except Exception as e:
        ctype = f"<err {e}>"
    try:
        ntitle = ch.ChartTitle.Text if ch.HasTitle else "(no title)"
    except Exception:
        ntitle = "(title err)"
    try:
        nser = ch.SeriesCollection().Count
    except Exception:
        nser = "?"
    print(f"slide {i:2d}: chart type={ctype:>6} series={nser} title={ntitle!r}")
    ok += 1
print(f"\n{ok} slides with charts, {bad} missing")
pres.Close()
app.Quit()

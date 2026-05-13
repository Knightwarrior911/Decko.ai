"""Diagnose + force-apply data label number format on slide 33's waterfall."""
import sys
from pathlib import Path
import win32com.client, pythoncom

DECK = r"C:\Users\vinit\Downloads\decko_loadfile_test.pptx"
FMT = "#,##0;(#,##0)"

pythoncom.CoInitialize()
app = win32com.client.Dispatch("PowerPoint.Application")
app.Visible = True
# Find / open
opened = None
for i in range(1, app.Presentations.Count + 1):
    pres = app.Presentations(i)
    if str(pres.FullName).lower() == DECK.lower():
        opened = pres
        break
if opened is None:
    opened = app.Presentations.Open(DECK)

sl = opened.Slides(33)
chart_shape = None
for sh in sl.Shapes:
    if sh.HasChart:
        chart_shape = sh
        break
if chart_shape is None:
    print("No chart on slide 33")
    sys.exit(1)

ch = chart_shape.Chart
print(f"chart type: {ch.ChartType}")
sc = ch.SeriesCollection()
print(f"series count: {sc.Count}")
print(f"chart.HasTitle: {ch.HasTitle}")

for i in range(1, sc.Count + 1):
    ser = sc(i)
    try:
        nm = ser.Name
    except Exception as e:
        nm = f"<err:{e}>"
    try:
        has_lbl = ser.HasDataLabels
    except Exception as e:
        has_lbl = f"<err:{e}>"
    print(f"  series {i}: name={nm!r} HasDataLabels={has_lbl}")
    # Force-enable + format
    try:
        ser.HasDataLabels = True
    except Exception as e:
        print(f"    SET HasDataLabels failed: {e}")
    try:
        ser.DataLabels().NumberFormat = FMT
        print(f"    set NumberFormat={FMT!r} OK")
    except Exception as e:
        print(f"    set NumberFormat failed: {e}")
    try:
        pts = ser.Points()
        n = pts.Count
        ok = 0
        for p in range(1, n + 1):
            try:
                pt = pts(p)
                pt.DataLabel.NumberFormat = FMT
                ok += 1
            except Exception:
                pass
        print(f"    per-point NumberFormat: {ok}/{n} points set")
    except Exception as e:
        print(f"    per-point failed: {e}")

opened.Save()
print("Saved.")

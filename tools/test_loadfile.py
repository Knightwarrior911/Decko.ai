"""Exercise the load-from-file path: read data/test_all_charts.actions.json
(78 actions, ~14 KB — the batch that corrupted when pasted into the textbox)
and run it through ExecuteFromString, exactly as btnApply does after a file load.
"""
import json
import time
from pathlib import Path
import win32com.client

ROOT = Path(__file__).resolve().parent.parent
CARRIER = ROOT / "PPT_AI_Editor.pptm"
ACTIONS_FILE = ROOT / "data" / "test_all_charts.actions.json"
DECK = Path.home() / "Downloads" / "decko_loadfile_test.pptx"
LOG = Path(str(DECK) + ".action_log.jsonl")

app = win32com.client.DispatchEx("PowerPoint.Application")
app.Visible = True
try:
    app.AutomationSecurity = 1
except Exception:
    pass

if LOG.exists():
    LOG.unlink()

pres_carrier = app.Presentations.Open(str(CARRIER), WithWindow=True)
time.sleep(1.0)
if DECK.exists():
    try:
        DECK.unlink()
    except Exception:
        pass
deck = app.Presentations.Add()
deck.SaveAs(str(DECK))
# give it 2 slides so the file's "insert at position 1..39" assumption holds cleanly
lay = deck.SlideMaster.CustomLayouts(7)
while deck.Slides.Count < 2:
    deck.Slides.AddSlide(deck.Slides.Count + 1, lay)
deck.Windows(1).Activate()

text = ACTIONS_FILE.read_text(encoding="utf-8")
print(f"file: {ACTIONS_FILE}  ({len(text)} chars, {text.count(chr(10))+1} lines)")

# This is exactly btnApply_Click's effective call after a file load:
try:
    result = app.Run("ExecuteFromString", text)
except Exception:
    result = app.Run("PPT_AI_Editor.pptm!ExecuteFromString", text)
print(f"ExecuteFromString -> {result}")

# Parse the action log.
entries = [json.loads(l) for l in LOG.read_text(encoding="utf-8").splitlines() if l.strip()]
by_status = {}
bad = []
for e in entries:
    by_status[e["status"]] = by_status.get(e["status"], 0) + 1
    if e["status"] != "ok":
        bad.append((e["op"], json.dumps(e.get("params", ""))[:140], e.get("reason", "")))
print(f"log: {dict(by_status)}  (total {len(entries)})")
for op, params, reason in bad[:10]:
    print(f"  NOT-OK {op}: {reason}  | {params}")

# Verify charts landed.
nchart = 0
for i in range(1, deck.Slides.Count + 1):
    for sh in deck.Slides(i).Shapes:
        try:
            if sh.HasChart:
                nchart += 1
        except Exception:
            pass
print(f"deck: {deck.Slides.Count} slides, {nchart} chart shapes")
deck.Save()
print(f"saved: {DECK}")
print("(PowerPoint left open)")

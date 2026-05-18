"""Combine slide1 + slide2 action files into one 2-slide deck.

Slide 1 actions stay as-is (add_slide position 1, slide:1).
Slide 2 actions remapped: slide 1 -> 2, add_slide position 1 -> 2.
ref_name namespaces are disjoint (s1: tbl/hl/revchart; s2: ec/dt/cb_*).
"""
import json
from pathlib import Path

W = Path(__file__).resolve().parent
s1 = json.loads((W / "slide1.actions.json").read_text(encoding="utf-8"))["actions"]
s2 = json.loads((W / "slide2.actions.json").read_text(encoding="utf-8"))["actions"]

out = list(s1)
for a in s2:
    a = dict(a)
    if a.get("type") == "add_slide":
        a["position"] = 2
    if a.get("slide") == 1:
        a["slide"] = 2
    out.append(a)

(W / "citi_final.actions.json").write_text(
    json.dumps({"actions": out}, ensure_ascii=False, indent=1), encoding="utf-8")
print(f"wrote citi_final.actions.json  (s1={len(s1)} + s2={len(s2)} = {len(out)} actions)")

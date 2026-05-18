"""templates gate: drive Api template/Deck-DNA/spec ops on a real deck
via COM (NO LLM). Exit 0 only on PASS."""
import os
import sys
import tempfile
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from app.main import Api                                  # noqa: E402
from app.template_slots import (TEMPLATE_NAMES,           # noqa: E402
                                default_content)


def run_once() -> list[str]:
    fails: list[str] = []
    tmp = os.path.join(tempfile.mkdtemp(prefix="sp2_"), "d.pptx")
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    p = app.Presentations.Add()
    p.SaveAs(tmp)
    p.Close()
    app.Quit()
    time.sleep(2.0)

    api = Api()
    api.boot()
    r = api.open_session("file", tmp)
    if not r.get("ok"):
        return [f"open_session failed: {r}"]
    try:
        for name in TEMPLATE_NAMES:
            res = api.apply_template(name, default_content(name),
                                     {"mode": "append"})
            if not res.get("ok") or "applied" not in (
                    res.get("summary", "").lower()):
                fails.append(f"apply {name}: {res}")
        cap = api.capture_template("sp2_cap")
        if not cap.get("ok"):
            fails.append(f"capture: {cap}")
        names = [c["name"] for c in
                 api.list_captured_templates()["templates"]]
        if "sp2_cap" not in names:
            fails.append(f"captured not listed: {names}")
        rn = api.rename_template("sp2_cap", "sp2_ren")
        if not rn.get("ok"):
            fails.append(f"rename: {rn}")
        names2 = [c["name"] for c in
                  api.list_captured_templates()["templates"]]
        if "sp2_ren" not in names2 or "sp2_cap" in names2:
            fails.append(f"rename not reflected: {names2}")
        dl = api.delete_template("sp2_ren")
        if not dl.get("ok"):
            fails.append(f"delete: {dl}")
        names3 = [c["name"] for c in
                  api.list_captured_templates()["templates"]]
        if "sp2_ren" in names3:
            fails.append(f"delete not reflected: {names3}")
        v = api.generate_variants({"template": "title", "n": 2,
                                   "content": default_content("title")})
        if not v.get("ok"):
            fails.append(f"variants: {v}")
        bs = api.build_deck_from_spec(
            [{"template": "quote",
              "content": default_content("quote")}])
        if not bs.get("ok"):
            fails.append(f"build_from_spec: {bs}")
        es = api.extract_spec()
        if not es.get("ok") or "{" not in (es.get("spec") or ""):
            fails.append(f"extract_spec: {es}")
        return fails
    finally:
        api.shutdown()


def main() -> int:
    last = None
    for attempt in range(1, 4):
        try:
            fails = run_once()
            break
        except Exception as e:  # noqa: BLE001
            last = e
            print(f"  retry transient (attempt {attempt}): {e!r}")
            os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
            time.sleep(5.0)
    else:
        print(f"FAIL: templates gate failed after retries: {last!r}")
        return 1
    if fails:
        for f in fails:
            print(f"  FAIL [templates] {f}")
        print("\nRESULT: FAIL")
        return 1
    print("  templates: apply7 / capture / list / rename / delete / "
          "variants / build_from_spec / extract OK")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())

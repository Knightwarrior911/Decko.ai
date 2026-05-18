# SP2 — Templates & Deck-DNA Visual Layer — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add an in-app slide-over panel that drives the engine's template / Deck-DNA / decks-as-code surface directly (no LLM for deterministic ops), with optional AI slot-fill.

**Architecture:** Reuse the SP1 app. New pure `app/template_slots.py` (builtin slot maps + placeholders). New `Api` methods build canonical `{"actions":[…]}` and run via the existing single-COM-thread `DeckController`/`ExecuteFromString`; each logs a session turn. `fill_with_ai` reuses `LLMClient`. A right slide-over panel in `app/web`. Engine `src/` is FROZEN.

**Tech Stack:** Python 3.11, pywin32 (COM), httpx, pytest, pywebview, PyInstaller. Spec: `docs/superpowers/specs/2026-05-18-sp2-templates-ui-design.md` (decisions DC1–DC7 are locked — do not re-decide).

---

## File Structure

```
app/template_slots.py        NEW  pure: BUILTIN_SLOTS, PLACEHOLDERS, default_content(), TEMPLATE_NAMES
app/llm_client.py            MOD  add build_fill_prompt(slots, brief) module fn
app/deck_controller.py       MOD  add get_deck_spec(), captured_registry_path()
app/main.py                  MOD  Api: 10 new methods (templates surface)
app/web/index.html           MOD  Templates button + slide-over panel markup
app/web/app.css              MOD  panel styles
app/web/app.js               MOD  panel logic
tests/app/test_template_slots.py   NEW  pure unit
tests/app/test_llm_client.py       MOD  build_fill_prompt + fill parse (mocked httpx)
tests/app/test_store.py            MOD  list_captured registry-read unit (temp file)
tests/run_smoke_app_templates.py   NEW  COM gate (no LLM): apply7/capture/list/rename/delete/variants/spec/extract
tests/run_smoke_app.py             MOD  add "templates" gate
README.md                          MOD  SP2 note
```

Engine `src/*.bas`/`*.frm` and `PPT_AI_Editor.pptm` are NEVER modified.

Carrier entry points used (all confirmed Public, callable via
`app.Run("PPT_AI_Editor!<name>")`): `ExecuteFromString(json)`,
`ExtractDeckSpecJson()`, `DefaultRegistryPath()`,
`NumberedTemplateList(path)`. Engine actions used inside
`ExecuteFromString`: `apply_template`, `capture_template`,
`delete_template`, `rename_template` (params `from`/`to`),
`generate_variants`, `build_deck_from_spec`.

---

## Task 1: template_slots.py (pure)

**Files:**
- Create: `app/template_slots.py`
- Test: `tests/app/test_template_slots.py`

- [ ] **Step 1: Write the failing test** — `tests/app/test_template_slots.py`
```python
from app.template_slots import (BUILTIN_SLOTS, TEMPLATE_NAMES,
                                 default_content)


def test_seven_builtins_with_authoritative_slots():
    assert set(TEMPLATE_NAMES) == {
        "title", "section", "bullets", "two_col", "comparison",
        "kpi_dashboard", "quote"}
    assert BUILTIN_SLOTS["title"] == ["title", "subtitle"]
    assert BUILTIN_SLOTS["section"] == ["section_number", "section_title"]
    assert BUILTIN_SLOTS["bullets"] == ["heading", "bullets"]
    assert BUILTIN_SLOTS["two_col"] == ["heading", "left_body",
                                        "right_body"]
    assert BUILTIN_SLOTS["comparison"] == ["heading", "left_label",
                                           "left_body", "right_label",
                                           "right_body"]
    assert BUILTIN_SLOTS["kpi_dashboard"] == ["heading", "tiles"]
    assert BUILTIN_SLOTS["quote"] == ["quote_text", "attribution"]


def test_default_content_shapes():
    c = default_content("bullets")
    assert isinstance(c["bullets"], list) and c["bullets"]
    assert isinstance(c["heading"], str) and c["heading"]
    k = default_content("kpi_dashboard")
    assert isinstance(k["tiles"], list)
    assert set(k["tiles"][0].keys()) == {"stat", "label"}
    t = default_content("title")
    assert set(t.keys()) == {"title", "subtitle"}


def test_default_content_unknown_raises():
    import pytest
    with pytest.raises(KeyError):
        default_content("nope")
```

- [ ] **Step 2: Run → fail**
Run: `python -m pytest tests/app/test_template_slots.py -q`
Expected: FAIL `ModuleNotFoundError: No module named 'app.template_slots'`

- [ ] **Step 3: Implement** — `app/template_slots.py`
```python
"""Authoritative builtin-template slot map + placeholder defaults.
Mirrors modActionsTemplate.ValidateTemplateSlots (engine FROZEN).
Pure data — no COM, no I/O."""

BUILTIN_SLOTS = {
    "title": ["title", "subtitle"],
    "section": ["section_number", "section_title"],
    "bullets": ["heading", "bullets"],
    "two_col": ["heading", "left_body", "right_body"],
    "comparison": ["heading", "left_label", "left_body",
                   "right_label", "right_body"],
    "kpi_dashboard": ["heading", "tiles"],
    "quote": ["quote_text", "attribution"],
}

TEMPLATE_NAMES = list(BUILTIN_SLOTS)

_PLACEHOLDER = {
    "title": "Title",
    "subtitle": "Subtitle",
    "section_number": "01",
    "section_title": "Section title",
    "heading": "Heading",
    "left_body": "Left content",
    "right_body": "Right content",
    "left_label": "Option A",
    "right_label": "Option B",
    "quote_text": "Quote goes here.",
    "attribution": "Attribution",
}


def default_content(template: str) -> dict:
    slots = BUILTIN_SLOTS[template]          # KeyError if unknown
    out = {}
    for s in slots:
        if s == "bullets":
            out[s] = ["First point", "Second point", "Third point"]
        elif s == "tiles":
            out[s] = [{"stat": "00", "label": "Metric one"},
                      {"stat": "00", "label": "Metric two"},
                      {"stat": "00", "label": "Metric three"}]
        else:
            out[s] = _PLACEHOLDER.get(s, s.replace("_", " ").title())
    return out
```

- [ ] **Step 4: Run → pass**
Run: `python -m pytest tests/app/test_template_slots.py -q`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**
```bash
git add app/template_slots.py tests/app/test_template_slots.py
git commit -m "$(cat <<'EOF'
feat(app): template_slots — builtin slot map + placeholder defaults

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 2: build_fill_prompt (pure) + parse

**Files:**
- Modify: `app/llm_client.py` (add a module function; do NOT change `LLMClient`/`call`/`sanitize_actions`)
- Test: `tests/app/test_llm_client.py` (append)

- [ ] **Step 1: Append failing test** to `tests/app/test_llm_client.py`
```python
def test_build_fill_prompt_names_slots_and_brief():
    from app.llm_client import build_fill_prompt
    p = build_fill_prompt(["title", "subtitle"], "Q3 board review")
    assert "title" in p and "subtitle" in p
    assert "Q3 board review" in p
    assert "JSON" in p
```

- [ ] **Step 2: Run → fail**
Run: `python -m pytest tests/app/test_llm_client.py::test_build_fill_prompt_names_slots_and_brief -q`
Expected: FAIL `ImportError: cannot import name 'build_fill_prompt'`

- [ ] **Step 3: Implement** — append to `app/llm_client.py` (module scope, after `sanitize_actions`)
```python
def build_fill_prompt(slots: list[str], brief: str) -> str:
    return (
        "Return ONLY a JSON object whose keys are EXACTLY these slots: "
        + ", ".join(slots) + ".\n"
        "For a slot named 'bullets' use a JSON array of short strings. "
        "For 'tiles' use a JSON array of objects {\"stat\":\"\",\"label\":\"\"}. "
        "All other slots are short strings.\n"
        "Draft concise, professional content for this brief:\n"
        + brief + "\n"
        'Output only the JSON object, no prose, no code fences.'
    )
```

- [ ] **Step 4: Run → pass**
Run: `python -m pytest tests/app/test_llm_client.py -q`
Expected: PASS (existing + 1 new)

- [ ] **Step 5: Commit**
```bash
git add app/llm_client.py tests/app/test_llm_client.py
git commit -m "$(cat <<'EOF'
feat(app): build_fill_prompt for Fill-with-AI slot drafting

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 3: DeckController helpers (additive, COM)

**Files:**
- Modify: `app/deck_controller.py` (add two methods; do NOT change existing methods/signatures)

No unit test (COM exercised by the Task 9 harness — project discipline).

- [ ] **Step 1: Add methods** — insert into `class DeckController` right after the existing `get_prompt_template` method:
```python
    def get_deck_spec(self) -> str:
        # Reverse of build_deck_from_spec — returns the deck's spec JSON.
        return self._run("ExtractDeckSpecJson")

    def captured_registry_path(self) -> str:
        return self._run("DefaultRegistryPath")

    def run_action(self, action: dict) -> str:
        # One-off action batch via the frozen engine (verify loop on).
        import json
        return self._run("ExecuteFromString",
                         json.dumps({"actions": [action]}))
```

- [ ] **Step 2: Syntax check**
Run: `python -c "import ast;ast.parse(open('app/deck_controller.py').read());print('ok')"`
Expected: `ok`
Run: `python -m pytest tests/app -q`
Expected: existing suite still green (deck_controller not imported at collection)

- [ ] **Step 3: Commit**
```bash
git add app/deck_controller.py
git commit -m "$(cat <<'EOF'
feat(app): DeckController get_deck_spec / captured_registry_path / run_action

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 4: Api — list_builtin_templates + list_captured_templates

**Files:**
- Modify: `app/main.py` (add two `Api` methods + needed imports)
- Test: `tests/app/test_store.py` (append — list_captured registry parsing, no COM)

The registry JSON shape (written by the engine) is
`{"templates": {"<name>": {<slot>: <value>, ...}, ...}}`. Parsing is
pure and unit-testable; reading the path is COM (Task 9 covers it).

- [ ] **Step 1: Append failing test** to `tests/app/test_store.py`
```python
def test_parse_captured_registry(tmp_path):
    import json
    from app.main import parse_captured_registry
    reg = tmp_path / "templates.json"
    reg.write_text(json.dumps({"templates": {
        "kpi_card": {"value": "X", "label": "Y"},
        "hero": {"title": "T"}}}), encoding="utf-8")
    out = parse_captured_registry(str(reg))
    names = {o["name"] for o in out}
    assert names == {"kpi_card", "hero"}
    kpi = next(o for o in out if o["name"] == "kpi_card")
    assert sorted(kpi["slots"]) == ["label", "value"]


def test_parse_captured_registry_missing_file(tmp_path):
    from app.main import parse_captured_registry
    assert parse_captured_registry(str(tmp_path / "none.json")) == []
```

- [ ] **Step 2: Run → fail**
Run: `python -m pytest tests/app/test_store.py::test_parse_captured_registry -q`
Expected: FAIL `ImportError: cannot import name 'parse_captured_registry'`

- [ ] **Step 3: Implement** — in `app/main.py`:
  add at module scope (after imports):
```python
def parse_captured_registry(path: str) -> list:
    import json
    import os
    if not path or not os.path.exists(path):
        return []
    try:
        data = json.loads(open(path, "r", encoding="utf-8").read())
    except Exception:
        return []
    tpls = (data or {}).get("templates", {})
    return [{"name": n, "slots": sorted(list(v.keys()))}
            for n, v in tpls.items()] if isinstance(tpls, dict) else []
```
  add these methods to `class Api` (after `boot`):
```python
    def list_builtin_templates(self):
        from app.template_slots import BUILTIN_SLOTS
        return {"templates": [{"name": n, "slots": s}
                              for n, s in BUILTIN_SLOTS.items()]}

    def list_captured_templates(self):
        if self.dc is None:
            return {"templates": []}
        try:
            path = self._com(self.dc.captured_registry_path)
        except Exception:
            return {"templates": []}
        return {"templates": parse_captured_registry(path)}
```

- [ ] **Step 4: Run → pass**
Run: `python -m pytest tests/app/test_store.py -q`
Expected: PASS (existing + 2 new)

- [ ] **Step 5: Commit**
```bash
git add app/main.py tests/app/test_store.py
git commit -m "$(cat <<'EOF'
feat(app): Api list_builtin_templates / list_captured_templates

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 5: Api — apply_template + capture_template (COM; logged turns)

**Files:**
- Modify: `app/main.py` (add two `Api` methods)

No unit test (COM; Task 9 harness gates these).

- [ ] **Step 1: Implement** — add to `class Api` (after `send`):
```python
    def _log_turn(self, request: str, summary: str, actions: dict):
        import json
        warnings = 0
        try:
            warnings = self.orch._warn_count(summary) if self.orch else 0
        except Exception:
            warnings = 0
        if self.session_id is not None:
            self.store.add_turn(request=request,
                                actions_json=json.dumps(actions),
                                result_summary=summary or "",
                                warnings=warnings,
                                session_id=self.session_id)

    def _require_session(self):
        if self.orch is None or self.dc is None:
            return {"error": "Start a session first."}
        return None

    def apply_template(self, template, content, target):
        g = self._require_session()
        if g:
            return g
        act = {"type": "apply_template", "template": template,
               "content": content}
        tgt = "append"
        if target and target.get("mode") == "replace":
            act["slide"] = int(target.get("slide"))
            tgt = f"replace slide {act['slide']}"
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Apply template: {template} ({tgt})",
                       summary, {"actions": [act]})
        return {"ok": True, "summary": summary}

    def capture_template(self, name):
        g = self._require_session()
        if g:
            return g
        if not str(name).strip():
            return {"error": "Enter a template name."}
        act = {"type": "capture_template", "name": str(name).strip()}
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Capture template: {name}", summary,
                       {"actions": [act]})
        return {"ok": True, "summary": summary}
```

- [ ] **Step 2: Syntax + unit regression**
Run: `python -c "import ast;ast.parse(open('app/main.py').read());print('ok')"` → `ok`
Run: `python -m pytest tests/app -q` → existing suite green

- [ ] **Step 3: Commit**
```bash
git add app/main.py
git commit -m "$(cat <<'EOF'
feat(app): Api apply_template + capture_template (logged turns)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 6: Api — rename_template + delete_template (COM)

**Files:**
- Modify: `app/main.py`

- [ ] **Step 1: Implement** — add to `class Api`:
```python
    def rename_template(self, from_name, to_name):
        g = self._require_session()
        if g:
            return g
        if not str(from_name).strip() or not str(to_name).strip():
            return {"error": "Both names are required."}
        act = {"type": "rename_template",
               "from": str(from_name).strip(),
               "to": str(to_name).strip()}
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Rename template: {from_name} -> {to_name}",
                       summary, {"actions": [act]})
        return {"ok": True, "summary": summary}

    def delete_template(self, name):
        g = self._require_session()
        if g:
            return g
        act = {"type": "delete_template", "name": str(name).strip()}
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn(f"Delete template: {name}", summary,
                       {"actions": [act]})
        return {"ok": True, "summary": summary}
```

- [ ] **Step 2: Syntax + unit regression**
Run: `python -c "import ast;ast.parse(open('app/main.py').read());print('ok')"` → `ok`
Run: `python -m pytest tests/app -q` → green

- [ ] **Step 3: Commit**
```bash
git add app/main.py
git commit -m "$(cat <<'EOF'
feat(app): Api rename_template + delete_template

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 7: Api — generate_variants + build_deck_from_spec + extract_spec (COM)

**Files:**
- Modify: `app/main.py`

- [ ] **Step 1: Implement** — add to `class Api`:
```python
    def generate_variants(self, payload):
        # payload: {"template": name, "n": int, "content": {...}} OR
        #          {"templates": [names], "content": {...}}
        g = self._require_session()
        if g:
            return g
        act = {"type": "generate_variants"}
        act.update(payload or {})
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn("Generate variants", summary, {"actions": [act]})
        return {"ok": True, "summary": summary}

    def build_deck_from_spec(self, spec, clear_existing=False):
        g = self._require_session()
        if g:
            return g
        import json
        if isinstance(spec, str):
            try:
                spec = json.loads(spec)
            except Exception as e:  # noqa: BLE001
                return {"error": f"Spec is not valid JSON: {e}"}
        act = {"type": "build_deck_from_spec", "spec": spec}
        if clear_existing:
            act["clear_existing"] = True
        try:
            summary = self._com(lambda: self.dc.run_action(act))
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        self._log_turn("Build deck from spec", summary,
                       {"actions": [act]})
        return {"ok": True, "summary": summary}

    def extract_spec(self):
        g = self._require_session()
        if g:
            return g
        try:
            js = self._com(self.dc.get_deck_spec)
        except Exception as e:  # noqa: BLE001
            return {"error": f"{type(e).__name__}: {e}"}
        return {"ok": True, "spec": js}
```

- [ ] **Step 2: Syntax + unit regression**
Run: `python -c "import ast;ast.parse(open('app/main.py').read());print('ok')"` → `ok`
Run: `python -m pytest tests/app -q` → green

- [ ] **Step 3: Commit**
```bash
git add app/main.py
git commit -m "$(cat <<'EOF'
feat(app): Api generate_variants + build_deck_from_spec + extract_spec

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 8: Api — fill_with_ai (LLM; mocked test)

**Files:**
- Modify: `app/main.py`
- Test: `tests/app/test_llm_client.py` (append — end-to-end fill via mocked transport)

- [ ] **Step 1: Append failing test** to `tests/app/test_llm_client.py`
```python
def test_fill_with_ai_via_mock(monkeypatch):
    import httpx, json as _j
    from app.config import Settings
    from app.llm_client import LLMClient, build_fill_prompt

    captured = {}
    def handler(req):
        captured["body"] = _j.loads(req.content)
        return httpx.Response(200, json={"choices": [{"message":
            {"content": '{"title":"Q3","subtitle":"Review"}'}}]})
    c = LLMClient(Settings(provider="openai", model="gpt-4o"),
                  api_key="sk-x",
                  _transport=httpx.MockTransport(handler))
    prompt = build_fill_prompt(["title", "subtitle"], "Q3 board review")
    out = c.call(snapshot="", user_request=prompt)
    obj = _j.loads(out) if out.strip().startswith("{") else None
    # call() wraps in sanitize_actions which requires "actions";
    # fill uses raw chat instead — see Api.fill_with_ai which calls
    # the provider directly. This test asserts the provider is reached
    # with the slot names + brief in the body.
    sent = _j.dumps(captured["body"])
    assert "title" in sent and "subtitle" in sent
    assert "Q3 board review" in sent
```

  NOTE: `LLMClient.call` runs `sanitize_actions` (expects an `actions`
  array). Fill returns a plain slot object, so `Api.fill_with_ai` must
  NOT go through `call()`. It uses a dedicated raw path. Adjust the test
  to assert via the Api path instead:
```python
def test_fill_with_ai_returns_slot_dict(monkeypatch):
    import httpx, json as _j
    import app.main as M
    class FakeLLM:
        def __init__(self,*a,**k): pass
        def raw(self, prompt):
            return '{"title":"Q3","subtitle":"Review"}'
    monkeypatch.setattr(M, "LLMClient", FakeLLM)
    api = M.Api()
    api.settings = __import__("app.config", fromlist=["Settings"]).Settings(
        provider="openai", model="gpt-4o")
    monkeypatch.setattr(M.secrets, "get_api_key", lambda: "sk-x")
    monkeypatch.setattr(M.secrets, "has_api_key", lambda: True)
    r = api.fill_with_ai("title", "Q3 board review")
    assert r["content"] == {"title": "Q3", "subtitle": "Review"}
```
  (Use ONLY the second test; delete the first if you wrote it.)

- [ ] **Step 2: Run → fail**
Run: `python -m pytest tests/app/test_llm_client.py::test_fill_with_ai_returns_slot_dict -q`
Expected: FAIL (`AttributeError`/`raw` not found, or fill_with_ai missing)

- [ ] **Step 3: Implement**
  Add a `raw()` method to `LLMClient` in `app/llm_client.py` (a plain
  completion that does NOT run `sanitize_actions`) — insert after
  `call`:
```python
    def raw(self, prompt: str) -> str:
        if self.s.provider == "anthropic":
            url = "https://api.anthropic.com/v1/messages"
            headers = {"x-api-key": self.api_key,
                       "anthropic-version": "2023-06-01"}
            body = {"model": self.s.model, "max_tokens": 4000,
                    "messages": [{"role": "user", "content": prompt}]}
            with self._http() as h:
                r = h.post(url, json=body, headers=headers)
                r.raise_for_status()
                return r.json()["content"][0]["text"]
        base = ("https://api.openai.com/v1"
                if self.s.provider == "openai"
                else self.s.base_url.rstrip("/"))
        headers = {"authorization": f"Bearer {self.api_key}"}
        body = {"model": self.s.model,
                "messages": [{"role": "user", "content": prompt}]}
        with self._http() as h:
            r = h.post(base + "/chat/completions", json=body,
                       headers=headers)
            r.raise_for_status()
            return r.json()["choices"][0]["message"]["content"]
```
  Add `fill_with_ai` to `class Api` in `app/main.py`:
```python
    def fill_with_ai(self, template, brief):
        if not secrets.has_api_key():
            return {"error": "Save your LLM API key first."}
        from app.llm_client import build_fill_prompt
        from app.template_slots import BUILTIN_SLOTS
        slots = BUILTIN_SLOTS.get(template)
        if slots is None:
            cap = self.list_captured_templates()["templates"]
            m = next((c for c in cap if c["name"] == template), None)
            slots = m["slots"] if m else []
        if not slots:
            return {"error": f"Unknown template '{template}'."}
        llm = LLMClient(self.settings, secrets.get_api_key() or "")
        try:
            raw = llm.raw(build_fill_prompt(slots, brief))
        except Exception as e:  # noqa: BLE001
            import httpx
            if isinstance(e, httpx.HTTPStatusError):
                body = ""
                try:
                    body = e.response.text[:800]
                except Exception:
                    pass
                return {"error": f"LLM API error "
                                 f"{e.response.status_code}: {body}"}
            return {"error": f"{type(e).__name__}: {e}"}
        import json
        import re
        s = re.sub(r"```[a-zA-Z]*\n?", "", raw).replace("```", "").strip()
        i, j = s.find("{"), s.rfind("}")
        if i == -1 or j == -1:
            return {"error": "AI did not return JSON."}
        try:
            content = json.loads(s[i:j + 1])
        except Exception as e:  # noqa: BLE001
            return {"error": f"AI JSON parse failed: {e}"}
        return {"ok": True, "content": content}
```

- [ ] **Step 4: Run → pass**
Run: `python -m pytest tests/app -q`
Expected: PASS (all green incl. the new fill test)

- [ ] **Step 5: Commit**
```bash
git add app/main.py app/llm_client.py tests/app/test_llm_client.py
git commit -m "$(cat <<'EOF'
feat(app): fill_with_ai — LLM drafts slot content (raw path, mocked test)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 9: COM gate — tests/run_smoke_app_templates.py

**Files:**
- Create: `tests/run_smoke_app_templates.py`

This harness IS the templates gate. No LLM. Single isolated PowerPoint
run, transient-COM retry, kill orphan POWERPNT (SP1 pattern). It drives
the `Api` exactly like the panel will.

- [ ] **Step 1: Write the harness** — `tests/run_smoke_app_templates.py`
```python
"""templates gate: drive Api template/Deck-DNA/spec ops on a real deck
via COM (NO LLM). Exit 0 only on PASS."""
import os
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from app.main import Api                                  # noqa: E402
from app.template_slots import (TEMPLATE_NAMES,           # noqa: E402
                                default_content)


def _seed_deck(api):
    # Build a non-empty deck via apply_template title (append).
    return api.apply_template("title", default_content("title"),
                              {"mode": "append"})


def run_once() -> list[str]:
    fails: list[str] = []
    tmp = os.path.join(tempfile.mkdtemp(prefix="sp2_"), "d.pptx")
    # start from a blank deck the engine can append to
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
        # apply all 7 builtins (append)
        for name in TEMPLATE_NAMES:
            res = api.apply_template(name, default_content(name),
                                     {"mode": "append"})
            if not res.get("ok") or "applied" not in (
                    res.get("summary", "").lower()):
                fails.append(f"apply {name}: {res}")
        # capture active slide -> list -> rename -> delete
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
        # variants (append-only)
        v = api.generate_variants({"template": "title", "n": 2,
                                   "content": default_content("title")})
        if not v.get("ok"):
            fails.append(f"variants: {v}")
        # build_deck_from_spec (append)
        bs = api.build_deck_from_spec(
            [{"template": "quote",
              "content": default_content("quote")}])
        if not bs.get("ok"):
            fails.append(f"build_from_spec: {bs}")
        # extract_spec returns JSON text
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
```

- [ ] **Step 2: Run the gate**
Run: `taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2; python tests/run_smoke_app_templates.py ; echo "EXIT=$?"`
Expected: `RESULT: PASS`, `EXIT=0`.
If FAIL: read the printed `FAIL [templates] …` line; fix only
`app/*.py` or this harness — NEVER `src/`. Re-run until PASS or
BLOCKED with a diagnosis.

- [ ] **Step 3: Commit**
```bash
git add tests/run_smoke_app_templates.py
git commit -m "$(cat <<'EOF'
test(app): templates COM gate (apply7/capture/list/rename/delete/variants/spec)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 10: Add templates gate to the aggregator

**Files:**
- Modify: `tests/run_smoke_app.py`

- [ ] **Step 1: Edit** — in `tests/run_smoke_app.py`, add to the `GATES`
list (after `core_loop`, before `packaging_smoke`):
```python
    ("templates",      [sys.executable,
                        "tests/run_smoke_app_templates.py"]),
```

- [ ] **Step 2: Run the full metric**
Run: `taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2; python tests/run_smoke_app.py ; echo "EXIT=$?"`
Expected: each `=== gate ===` then `RESULT: PASS`, `EXIT=0`
(rebuilds via PyInstaller — minutes — fine).

- [ ] **Step 3: Commit**
```bash
git add tests/run_smoke_app.py
git commit -m "$(cat <<'EOF'
test(app): run_smoke_app — add templates gate to the SP2 metric

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 11: Slide-over Templates panel UI (manual-verified)

**Files:**
- Modify: `app/web/index.html`, `app/web/app.css`, `app/web/app.js`

UI is NOT in the deterministic gate (spec §6 — manual screenshot). No
TDD; build + eyeball. Do NOT remove existing chat/sidebar markup or the
`pywebviewready`/`boot` wiring.

- [ ] **Step 1: index.html** — add a Templates toggle button in the
left sidebar Deck section (after `savePptBtn`):
```html
      <button id="tplBtn" class="ghost">Templates ▸</button>
```
  and add the slide-over panel as the LAST child of `#layout` (before
  the closing `</div><script>`):
```html
  <aside id="tplPanel" class="tpl-hidden">
    <div class="tplHead"><b>Templates &amp; Deck DNA</b>
      <button id="tplClose" class="ghost">✕</button></div>

    <details open><summary>Built-in templates</summary>
      <select id="tplPick"></select>
      <div id="tplSlots"></div>
      <div class="tplRow">
        <select id="tplTarget">
          <option value="append">Append new slide</option>
          <option value="replace">Replace slide…</option>
        </select>
        <input id="tplSlideNo" type="number" min="1" value="1" hidden>
      </div>
      <input id="tplBrief" placeholder="Optional: one-line brief for Fill-with-AI">
      <div class="tplRow">
        <button id="tplFill" class="ghost">Fill with AI</button>
        <button id="tplApply">Apply</button>
      </div>
    </details>

    <details><summary>Deck DNA (captured)</summary>
      <input id="capName" placeholder="Name to capture active slide as">
      <button id="capBtn">Capture active slide</button>
      <ul id="capList"></ul>
    </details>

    <details><summary>Variants</summary>
      <select id="varTpl"></select>
      <input id="varN" type="number" min="1" value="3">
      <button id="varBtn">Generate variants (append)</button>
    </details>

    <details><summary>Decks-as-code</summary>
      <button id="specExtract" class="ghost">Extract spec from deck</button>
      <textarea id="specBox" placeholder='[{"template":"title","content":{...}}]'></textarea>
      <label class="muted"><input type="checkbox" id="specClear"> Replace whole deck</label>
      <button id="specBuild">Build deck from spec</button>
    </details>
  </aside>
```

- [ ] **Step 2: app.css** — append:
```css
#tplPanel{position:fixed;top:0;right:0;height:100vh;width:380px;
 background:#161618;border-left:2px solid #3a6ea5;padding:14px;
 overflow:auto;box-shadow:-8px 0 24px rgba(0,0,0,.4);z-index:50}
#tplPanel.tpl-hidden{display:none}
.tplHead{display:flex;justify-content:space-between;align-items:center;margin-bottom:8px}
#tplPanel details{margin:.5rem 0;border:1px solid #2a2a2e;border-radius:6px;padding:.4rem .6rem}
#tplPanel summary{cursor:pointer;font-size:.8rem;text-transform:uppercase;color:#9ab}
#tplPanel input,#tplPanel select,#tplPanel textarea,#tplPanel button{width:100%;margin:.25rem 0;padding:.45rem;background:#222;color:#eee;border:1px solid #333;border-radius:6px}
#tplPanel .tplRow{display:flex;gap:6px}
#tplPanel .tplRow>*{flex:1}
#specBox{height:120px;font-family:monospace;font-size:.74rem;white-space:pre}
#capList{list-style:none;padding:0;margin:.3rem 0}
#capList li{display:flex;gap:4px;align-items:center;background:#1d1d20;border:1px solid #2a2a2e;border-radius:6px;padding:.3rem .4rem;margin:.2rem 0;font-size:.78rem}
#capList li span{flex:1}
#capList li button{width:auto;padding:2px 6px;font-size:.7rem;margin:0}
```

- [ ] **Step 3: app.js** — append (uses the existing `$`, `bubble`,
`api`, `currentSession`, `refreshSessions` already defined; do not
redeclare them):
```javascript
let BUILTINS = [];
async function tplInit() {
  BUILTINS = (await api.list_builtin_templates()).templates;
  const pick = $("tplPick"), vt = $("varTpl");
  pick.innerHTML = ""; vt.innerHTML = "";
  BUILTINS.forEach((t) => {
    pick.add(new Option(t.name, t.name));
    vt.add(new Option(t.name, t.name));
  });
  renderSlots();
  await refreshCaptured();
}
function curTpl() {
  return BUILTINS.find((t) => t.name === $("tplPick").value);
}
function renderSlots() {
  const t = curTpl(); const box = $("tplSlots"); box.innerHTML = "";
  if (!t) return;
  t.slots.forEach((s) => {
    const i = document.createElement("input");
    i.id = "slot_" + s; i.placeholder = s;
    box.appendChild(i);
  });
}
function collectContent() {
  const t = curTpl(); const c = {};
  t.slots.forEach((s) => {
    const v = ($("slot_" + s) || {}).value || "";
    if (s === "bullets") c[s] = v ? v.split("\n") : ["Point one"];
    else if (s === "tiles")
      c[s] = [{ stat: "00", label: v || "Metric" }];
    else c[s] = v || s;
  });
  return c;
}
async function refreshCaptured() {
  const r = await api.list_captured_templates();
  const ul = $("capList"); ul.innerHTML = "";
  (r.templates || []).forEach((t) => {
    const li = document.createElement("li");
    const sp = document.createElement("span"); sp.textContent = t.name;
    const ap = document.createElement("button"); ap.textContent = "Apply";
    ap.onclick = async () => {
      const res = await api.apply_template(t.name, {}, tgt());
      bubble(res.ok ? "app" : "fail", res.ok ? res.summary : res.error);
      refreshSessions(currentSession);
    };
    const dl = document.createElement("button"); dl.textContent = "Del";
    dl.onclick = async () => {
      await api.delete_template(t.name); refreshCaptured();
    };
    li.appendChild(sp); li.appendChild(ap); li.appendChild(dl);
    ul.appendChild(li);
  });
}
function tgt() {
  return $("tplTarget").value === "replace"
    ? { mode: "replace", slide: parseInt($("tplSlideNo").value || "1") }
    : { mode: "append" };
}
$("tplBtn").onclick = () => {
  $("tplPanel").classList.toggle("tpl-hidden");
  if (!$("tplPanel").classList.contains("tpl-hidden") && api) tplInit();
};
$("tplClose").onclick = () =>
  $("tplPanel").classList.add("tpl-hidden");
$("tplPick").onchange = renderSlots;
$("tplTarget").onchange = (e) =>
  ($("tplSlideNo").hidden = e.target.value !== "replace");
$("tplApply").onclick = async () => {
  const r = await api.apply_template($("tplPick").value,
    collectContent(), tgt());
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error, true);
  refreshSessions(currentSession);
};
$("tplFill").onclick = async () => {
  const r = await api.fill_with_ai($("tplPick").value,
    $("tplBrief").value || "professional placeholder content");
  if (r.error) { bubble("fail", r.error); return; }
  Object.entries(r.content).forEach(([k, v]) => {
    const el = $("slot_" + k);
    if (el) el.value = Array.isArray(v)
      ? v.map((x) => (typeof x === "object" ? JSON.stringify(x) : x)).join("\n")
      : v;
  });
  bubble("app", "AI filled the slots — review then Apply.");
};
$("capBtn").onclick = async () => {
  const r = await api.capture_template($("capName").value);
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error);
  if (r.ok) { $("capName").value = ""; refreshCaptured();
              refreshSessions(currentSession); }
};
$("varBtn").onclick = async () => {
  const r = await api.generate_variants({
    template: $("varTpl").value,
    n: parseInt($("varN").value || "3"),
    content: {} });
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error, true);
  refreshSessions(currentSession);
};
$("specExtract").onclick = async () => {
  const r = await api.extract_spec();
  if (r.error) { bubble("fail", r.error); return; }
  $("specBox").value = r.spec;
};
$("specBuild").onclick = async () => {
  const r = await api.build_deck_from_spec($("specBox").value,
    $("specClear").checked);
  bubble(r.ok ? "app" : "fail", r.ok ? r.summary : r.error, true);
  refreshSessions(currentSession);
};
```

- [ ] **Step 4: Manual verification (no automated gate)**
Run: `taskkill //F //IM python.exe 2>/dev/null; sleep 1; python -m app.main`
Check by eye: "Templates ▸" button opens the right slide-over; pick a
builtin → slot fields appear → choose Append → Apply → chat shows the
engine summary; Capture active slide + name → appears in the list →
Apply/Del work; Generate variants; Extract spec fills the box; Build
from spec runs. Capture a screenshot for the record. (Requires a
started session + a deck open.)

- [ ] **Step 5: Commit**
```bash
git add app/web/index.html app/web/app.css app/web/app.js
git commit -m "$(cat <<'EOF'
feat(app): slide-over Templates & Deck-DNA panel (manual-verified)

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Task 12: Final gate + README + branch finish

**Files:**
- Modify: `README.md`

- [ ] **Step 1: SP2 metric + engine guard** (sequential, kill+sleep
before each COM run)
```bash
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3
python tests/run_smoke_app.py ; echo "EXIT=$?"          # expect RESULT: PASS
git checkout -- PPT_AI_Editor.pptm                       # packaging re-baked it
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke_icon_prompt.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke_spec.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke_capture.py
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 3; python tests/run_smoke_dialogs.py
```
All must report PASS / "all tests passed". Confirm engine frozen:
`git diff --stat origin/main..HEAD -- src/` MUST be empty.

- [ ] **Step 2: README** — add under the "Desktop app (SP1)" section:
```markdown
### Templates panel (SP2)

The Export/desktop app has a slide-over **Templates ▸** panel: apply
the 7 built-in layouts or your captured "Deck DNA" templates with
instant placeholders (optional Fill-with-AI), capture the active slide,
rename/delete captured templates, generate layout variants, and
build-from / extract a deck spec — all without an LLM round-trip for
the deterministic operations. Gate: `tests/run_smoke_app.py` (templates
gate); UI verified manually.
```

- [ ] **Step 3: Commit**
```bash
git add README.md
git commit -m "$(cat <<'EOF'
docs: README — SP2 Templates panel

Co-Authored-By: Claude Opus 4.7 (1M context) <noreply@anthropic.com>
EOF
)"
```

---

## Self-Review

**Spec coverage:** DC1 slide-over panel — T11. DC2 hybrid placeholders +
Fill-with-AI — T1 (`default_content`), T8 (`fill_with_ai`), T11 (form +
Fill button). DC3 full surface — T5/T6/T7 (apply/capture/rename/delete/
variants/build/extract) + T11 UI. DC4 target (apply append/replace-N;
variants append-only; spec append/clear_existing) — T5 (`target`), T7
(`generate_variants` no slide; `build_deck_from_spec` `clear_existing`).
DC5 Approach A direct Api→_com→ExecuteFromString, logged turns — T3
(`run_action`), T5 (`_log_turn`). DC6 capture active slide+name — T5.
DC7 engine frozen — every COM task forbids `src/` edits; T12 asserts
empty `src/` diff. §3 `list_captured` from registry — T4. §4 extract
seeds spec editor — T11 (`specExtract`→`specBox`). §6 testing —
T1/T2/T4/T8 unit, T9 COM gate, T10 aggregator, T11 manual. §7 metric —
T12.

**Placeholder scan:** No TBD/TODO. Every code step has complete code.
T8 explicitly tells the engineer to use ONLY the second test and delete
the first — that is an instruction, not a placeholder.

**Type consistency:** `Api` methods return `{"ok":True,"summary":…}` or
`{"error":…}` (extract_spec → `{"ok":True,"spec":…}`; fill_with_ai →
`{"ok":True,"content":…}`) — consistent across T4–T8 and consumed
correctly by the T11 JS (`r.ok`/`r.summary`/`r.error`/`r.spec`/
`r.content`). `DeckController.run_action(action:dict)->str`,
`get_deck_spec()->str`, `captured_registry_path()->str` defined in T3,
used in T5–T7. `parse_captured_registry(path)->list[{name,slots}]`
defined T4, used by `fill_with_ai` T8 and JS T11. `target` dict shape
`{"mode":"append"}|{"mode":"replace","slide":N}` consistent T5/T11
(`tgt()`). `build_fill_prompt(slots,brief)` T2 used T8. No drift.

No gaps found.

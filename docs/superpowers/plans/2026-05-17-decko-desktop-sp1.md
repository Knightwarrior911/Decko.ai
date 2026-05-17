# Decko Desktop SP1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a Windows `.exe` that wraps the existing Decko VBA/COM engine so an end user with PowerPoint but no Python can chat-edit a deck using their own LLM key.

**Architecture:** pywebview native window (chat + side panel) over a Python core. Core = `DeckController` (PowerPoint COM), `LLMClient` (Anthropic/OpenAI/generic), `Store` (sync-ready SQLite), `Secrets` (keyring), `Orchestrator` (one turn). The carrier `.pptm` (246-action engine + verify loop) is reused unchanged via `app.Run("PPT_AI_Editor!ExecuteFromString"|"BuildSnapshotJson")`.

**Tech Stack:** Python 3.11, pywebview, pywin32 (win32com), keyring, httpx, pytest, PyInstaller, Inno Setup. Spec: `docs/superpowers/specs/2026-05-17-decko-desktop-app-design.md` (premises D1–D8 are locked — do not re-decide).

---

## File Structure

```
app/
  __init__.py          package marker
  config.py            paths (%APPDATA%\Decko), Settings dataclass
  store.py             SQLite, sync-ready schema, turn CRUD
  secrets.py           keyring wrapper for the API key
  llm_client.py        3 provider modes; build prompt; parse+sanitize
  carrier.py           locate bundled carrier; copy to %APPDATA% on first run
  deck_controller.py   PowerPoint COM: attach/file modes, snapshot, run, retry
  orchestrator.py      ChatOrchestrator.run(text): one full turn
  main.py              pywebview window + js_api bridge (UI; NOT gated)
  web/
    index.html         chat + side panel markup (UI; NOT gated)
    app.css            theme (UI; NOT gated)
    app.js             bridge calls + render (UI; NOT gated)
packaging/
  build.py             bake carrier via update_macros.py, then PyInstaller
  decko.spec           PyInstaller spec (bundles carrier + web/)
  installer.iss        Inno Setup script -> Decko-Setup.exe
requirements-app.txt   app runtime + build deps
tests/
  app/__init__.py
  app/stub_llm.py      deterministic fixed-actions stub
  app/test_store.py    store_unit (no COM, no network)
  app/test_llm_client.py  llmclient_unit (mocked HTTP, no network)
  run_smoke_app_core_loop.py   COM: stub-LLM -> ExecuteFromString, assert counts
  run_smoke_app_packaging.py   PyInstaller build + bundled carrier loads via COM
  run_smoke_app.py     AGGREGATOR (the metric): runs the four above, exit 0 at 100%
```

Engine files (`src/*.bas`, `src/*.frm`, the action surface) are **never modified** by this plan. SP1 wraps the engine.

---

## Task 1: App package skeleton + config

**Files:**
- Create: `requirements-app.txt`
- Create: `app/__init__.py`
- Create: `app/config.py`
- Create: `tests/app/__init__.py`
- Test: `tests/app/test_store.py` (created in Task 2; config has no logic to unit-test beyond path derivation, covered here)

- [ ] **Step 1: Write requirements-app.txt**

```
pywebview==5.3.2
pywin32==308
keyring==25.5.0
httpx==0.27.2
pyinstaller==6.11.1
pytest==8.3.3
```

- [ ] **Step 2: Create package markers**

`app/__init__.py`:
```python
"""Decko Desktop app package (SP1)."""
```

`tests/app/__init__.py`:
```python
```

- [ ] **Step 3: Write app/config.py**

```python
"""Paths and settings for Decko Desktop. 100% local (spec D4)."""
import os
from dataclasses import dataclass
from pathlib import Path

APP_DIR = Path(os.environ.get("APPDATA", str(Path.home()))) / "Decko"
ENGINE_DIR = APP_DIR / "engine"
DB_PATH = APP_DIR / "decko.db"
CARRIER_NAME = "PPT_AI_Editor.pptm"
INSTALLED_CARRIER = ENGINE_DIR / CARRIER_NAME

KEYRING_SERVICE = "DeckoDesktop"
KEYRING_USERNAME = "llm_api_key"

VALID_PROVIDERS = ("anthropic", "openai", "generic")


@dataclass
class Settings:
    provider: str = "anthropic"          # one of VALID_PROVIDERS
    model: str = "claude-opus-4-7"
    base_url: str = ""                   # required only when provider == "generic"

    def validate(self) -> None:
        if self.provider not in VALID_PROVIDERS:
            raise ValueError(f"provider must be one of {VALID_PROVIDERS}")
        if self.provider == "generic" and not self.base_url:
            raise ValueError("generic provider requires base_url")


def ensure_app_dirs() -> None:
    APP_DIR.mkdir(parents=True, exist_ok=True)
    ENGINE_DIR.mkdir(parents=True, exist_ok=True)
```

- [ ] **Step 4: Commit**

```bash
git add requirements-app.txt app/__init__.py app/config.py tests/app/__init__.py
git commit -m "feat(app): SP1 package skeleton + config (paths, Settings)"
```

---

## Task 2: Store (sync-ready SQLite)

**Files:**
- Create: `app/store.py`
- Test: `tests/app/test_store.py`

- [ ] **Step 1: Write the failing test**

`tests/app/test_store.py`:
```python
import uuid
from pathlib import Path

from app.store import Store


def test_schema_and_turn_roundtrip(tmp_path: Path):
    db = tmp_path / "t.db"
    s = Store(db)
    s.init()
    tid = s.add_turn(request="make title", actions_json='{"actions":[]}',
                      result_summary="applied 0", warnings=0)
    assert isinstance(tid, str) and len(tid) == 36          # UUID PK
    rows = s.list_turns()
    assert len(rows) == 1
    r = rows[0]
    assert r["id"] == tid
    assert r["request"] == "make title"
    assert r["updated_at"] is not None                       # sync-ready
    assert r["deleted"] == 0                                  # soft-delete col


def test_soft_delete_hides_row(tmp_path: Path):
    s = Store(tmp_path / "t.db")
    s.init()
    tid = s.add_turn(request="x", actions_json="{}", result_summary="", warnings=0)
    s.soft_delete(tid)
    assert s.list_turns() == []
    assert s.list_turns(include_deleted=True)[0]["deleted"] == 1
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/app/test_store.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.store'`

- [ ] **Step 3: Write minimal implementation**

`app/store.py`:
```python
"""Local SQLite store. Schema is sync-ready (UUID PK, updated_at,
soft-delete) so SP4 can add cloud sync without migration. Spec D4/§4."""
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path

_SCHEMA = """
CREATE TABLE IF NOT EXISTS turns (
    id             TEXT PRIMARY KEY,
    request        TEXT NOT NULL,
    actions_json   TEXT NOT NULL,
    result_summary TEXT NOT NULL,
    warnings       INTEGER NOT NULL DEFAULT 0,
    created_at     TEXT NOT NULL,
    updated_at     TEXT NOT NULL,
    deleted        INTEGER NOT NULL DEFAULT 0
);
"""


def _now() -> str:
    return datetime.now(timezone.utc).isoformat()


class Store:
    def __init__(self, db_path: Path):
        self.db_path = Path(db_path)

    def _conn(self) -> sqlite3.Connection:
        c = sqlite3.connect(self.db_path)
        c.row_factory = sqlite3.Row
        return c

    def init(self) -> None:
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        with self._conn() as c:
            c.executescript(_SCHEMA)

    def add_turn(self, request: str, actions_json: str,
                 result_summary: str, warnings: int) -> str:
        tid = str(uuid.uuid4())
        ts = _now()
        with self._conn() as c:
            c.execute(
                "INSERT INTO turns (id,request,actions_json,result_summary,"
                "warnings,created_at,updated_at,deleted) VALUES (?,?,?,?,?,?,?,0)",
                (tid, request, actions_json, result_summary, warnings, ts, ts),
            )
        return tid

    def soft_delete(self, tid: str) -> None:
        with self._conn() as c:
            c.execute("UPDATE turns SET deleted=1, updated_at=? WHERE id=?",
                      (_now(), tid))

    def list_turns(self, include_deleted: bool = False) -> list[dict]:
        q = "SELECT * FROM turns"
        if not include_deleted:
            q += " WHERE deleted=0"
        q += " ORDER BY created_at ASC"
        with self._conn() as c:
            return [dict(r) for r in c.execute(q)]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/app/test_store.py -v`
Expected: PASS (2 passed)

- [ ] **Step 5: Commit**

```bash
git add app/store.py tests/app/test_store.py
git commit -m "feat(app): sync-ready SQLite Store (turns, soft-delete)"
```

---

## Task 3: Secrets (keyring)

**Files:**
- Create: `app/secrets.py`
- Test: `tests/app/test_store.py` (append; keyring backend stubbed)

- [ ] **Step 1: Write the failing test (append to tests/app/test_store.py)**

```python
def test_secrets_set_get_clear(monkeypatch):
    import app.secrets as secmod
    bucket = {}
    monkeypatch.setattr(secmod.keyring, "set_password",
                        lambda s, u, p: bucket.__setitem__((s, u), p))
    monkeypatch.setattr(secmod.keyring, "get_password",
                        lambda s, u: bucket.get((s, u)))
    monkeypatch.setattr(secmod.keyring, "delete_password",
                        lambda s, u: bucket.pop((s, u), None))
    secmod.set_api_key("sk-test")
    assert secmod.get_api_key() == "sk-test"
    assert secmod.has_api_key() is True
    secmod.clear_api_key()
    assert secmod.get_api_key() is None
    assert secmod.has_api_key() is False
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/app/test_store.py::test_secrets_set_get_clear -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.secrets'`

- [ ] **Step 3: Write minimal implementation**

`app/secrets.py`:
```python
"""API key in Windows Credential Manager via keyring. Never in SQLite
or plaintext (spec §4)."""
import keyring

from app.config import KEYRING_SERVICE, KEYRING_USERNAME


def set_api_key(key: str) -> None:
    keyring.set_password(KEYRING_SERVICE, KEYRING_USERNAME, key)


def get_api_key() -> str | None:
    return keyring.get_password(KEYRING_SERVICE, KEYRING_USERNAME)


def has_api_key() -> bool:
    return bool(get_api_key())


def clear_api_key() -> None:
    try:
        keyring.delete_password(KEYRING_SERVICE, KEYRING_USERNAME)
    except keyring.errors.PasswordDeleteError:
        pass
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/app/test_store.py::test_secrets_set_get_clear -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add app/secrets.py tests/app/test_store.py
git commit -m "feat(app): keyring-backed API key storage"
```

---

## Task 4: LLMClient (3 providers, mocked HTTP)

**Files:**
- Create: `app/llm_client.py`
- Test: `tests/app/test_llm_client.py`

- [ ] **Step 1: Write the failing test**

`tests/app/test_llm_client.py`:
```python
import json

import httpx

from app.config import Settings
from app.llm_client import LLMClient


class _Capture:
    def __init__(self, payload_text):
        self.payload_text = payload_text
        self.seen = {}

    def handler(self, request: httpx.Request) -> httpx.Response:
        self.seen["url"] = str(request.url)
        self.seen["headers"] = dict(request.headers)
        self.seen["body"] = json.loads(request.content)
        return httpx.Response(200, json=self._reply())

    def _reply(self):
        raise NotImplementedError


class _Anthropic(_Capture):
    def _reply(self):
        return {"content": [{"type": "text", "text": self.payload_text}]}


class _OpenAI(_Capture):
    def _reply(self):
        return {"choices": [{"message": {"content": self.payload_text}}]}


ACTIONS = '{"actions":[{"type":"apply_template","template":"title",' \
          '"content":{"title":"Hi","subtitle":"There"}}]}'


def _client(settings, cap):
    transport = httpx.MockTransport(cap.handler)
    return LLMClient(settings, api_key="sk-x",
                     _transport=transport)


def test_anthropic_request_shape_and_parse():
    cap = _Anthropic("```json\n" + ACTIONS + "\n```")
    c = _client(Settings(provider="anthropic", model="claude-opus-4-7"), cap)
    out = c.call(snapshot="{SNAP}", user_request="make a title slide")
    assert cap.seen["url"] == "https://api.anthropic.com/v1/messages"
    assert cap.seen["headers"]["x-api-key"] == "sk-x"
    assert cap.seen["body"]["model"] == "claude-opus-4-7"
    # prompt carries snapshot + request
    sent = json.dumps(cap.seen["body"])
    assert "{SNAP}" in sent and "make a title slide" in sent
    # fenced JSON is sanitized to a clean actions object
    assert json.loads(out)["actions"][0]["type"] == "apply_template"


def test_openai_request_shape_and_parse():
    cap = _OpenAI(ACTIONS)
    c = _client(Settings(provider="openai", model="gpt-4o"), cap)
    out = c.call(snapshot="{S}", user_request="x")
    assert cap.seen["url"] == "https://api.openai.com/v1/chat/completions"
    assert cap.seen["headers"]["authorization"] == "Bearer sk-x"
    assert cap.seen["body"]["model"] == "gpt-4o"
    assert json.loads(out)["actions"][0]["template"] == "title"


def test_generic_uses_base_url_openai_schema():
    cap = _OpenAI(ACTIONS)
    s = Settings(provider="generic", model="deepseek-chat",
                 base_url="https://api.deepseek.com/v1")
    c = _client(s, cap)
    c.call(snapshot="{S}", user_request="x")
    assert cap.seen["url"] == "https://api.deepseek.com/v1/chat/completions"
    assert cap.seen["headers"]["authorization"] == "Bearer sk-x"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/app/test_llm_client.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.llm_client'`

- [ ] **Step 3: Write minimal implementation**

`app/llm_client.py`:
```python
"""LLM brain. BYO key. Three provider modes (spec D6):
  anthropic -> https://api.anthropic.com/v1/messages
  openai    -> https://api.openai.com/v1/chat/completions
  generic   -> {base_url}/chat/completions  (OpenAI schema)
Output is sanitized to a clean {"actions":[...]} JSON string."""
import json
import re

import httpx

from app.config import Settings

_SYSTEM = (
    "You are Decko's action generator. Given a PowerPoint deck snapshot "
    "and a user request, return ONLY a JSON object {\"actions\":[...]}. "
    "No prose, no markdown fences."
)


def _build_user_prompt(snapshot: str, user_request: str) -> str:
    return (
        "DECK SNAPSHOT:\n" + snapshot + "\n\n"
        "USER REQUEST:\n" + user_request + "\n\n"
        'Return only {"actions":[...]}.'
    )


def sanitize_actions(raw: str) -> str:
    """Strip fences/prose; return the outermost {...} containing "actions"."""
    text = raw.strip()
    text = re.sub(r"^```[a-zA-Z]*\n?|\n?```$", "", text).strip()
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end < start:
        raise ValueError("no JSON object in LLM output")
    obj = json.loads(text[start:end + 1])
    if "actions" not in obj or not isinstance(obj["actions"], list):
        raise ValueError('LLM output missing "actions" array')
    return json.dumps(obj)


class LLMClient:
    def __init__(self, settings: Settings, api_key: str,
                 _transport: httpx.BaseTransport | None = None):
        settings.validate()
        self.s = settings
        self.api_key = api_key
        self._transport = _transport

    def _http(self) -> httpx.Client:
        return httpx.Client(timeout=120.0, transport=self._transport)

    def call(self, snapshot: str, user_request: str) -> str:
        user = _build_user_prompt(snapshot, user_request)
        if self.s.provider == "anthropic":
            url = "https://api.anthropic.com/v1/messages"
            headers = {"x-api-key": self.api_key,
                       "anthropic-version": "2023-06-01"}
            body = {"model": self.s.model, "max_tokens": 8000,
                    "system": _SYSTEM,
                    "messages": [{"role": "user", "content": user}]}
            with self._http() as h:
                r = h.post(url, json=body, headers=headers)
            r.raise_for_status()
            raw = r.json()["content"][0]["text"]
        else:
            base = ("https://api.openai.com/v1" if self.s.provider == "openai"
                    else self.s.base_url.rstrip("/"))
            url = base + "/chat/completions"
            headers = {"authorization": f"Bearer {self.api_key}"}
            body = {"model": self.s.model,
                    "messages": [{"role": "system", "content": _SYSTEM},
                                 {"role": "user", "content": user}]}
            with self._http() as h:
                r = h.post(url, json=body, headers=headers)
            r.raise_for_status()
            raw = r.json()["choices"][0]["message"]["content"]
        return sanitize_actions(raw)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/app/test_llm_client.py -v`
Expected: PASS (3 passed)

- [ ] **Step 5: Commit**

```bash
git add app/llm_client.py tests/app/test_llm_client.py
git commit -m "feat(app): LLMClient — anthropic/openai/generic + sanitizer"
```

---

## Task 5: Carrier locate/copy

**Files:**
- Create: `app/carrier.py`
- Test: `tests/app/test_store.py` (append)

- [ ] **Step 1: Write the failing test (append to tests/app/test_store.py)**

```python
def test_carrier_copies_to_install_dir_once(tmp_path, monkeypatch):
    import app.carrier as cm
    src = tmp_path / "bundled" / "PPT_AI_Editor.pptm"
    src.parent.mkdir(parents=True)
    src.write_bytes(b"PPTM-BYTES")
    dest = tmp_path / "appdata" / "engine" / "PPT_AI_Editor.pptm"
    monkeypatch.setattr(cm, "_bundled_carrier", lambda: src)
    monkeypatch.setattr(cm, "INSTALLED_CARRIER", dest)
    p1 = cm.ensure_carrier()
    assert p1 == dest and dest.read_bytes() == b"PPTM-BYTES"
    dest.write_bytes(b"USER-MODIFIED")          # must NOT be overwritten
    p2 = cm.ensure_carrier()
    assert p2.read_bytes() == b"USER-MODIFIED"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/app/test_store.py::test_carrier_copies_to_install_dir_once -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.carrier'`

- [ ] **Step 3: Write minimal implementation**

`app/carrier.py`:
```python
"""Locate the bundled carrier and copy it to %APPDATA%\\Decko\\engine
on first run only. Never run update_macros on the user's machine
(spec §4)."""
import shutil
import sys
from pathlib import Path

from app.config import CARRIER_NAME, INSTALLED_CARRIER, ensure_app_dirs


def _bundled_carrier() -> Path:
    # PyInstaller unpacks data to sys._MEIPASS; dev fallback = repo root.
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))
    return base / CARRIER_NAME


def ensure_carrier() -> Path:
    ensure_app_dirs()
    if not INSTALLED_CARRIER.exists():
        shutil.copy2(_bundled_carrier(), INSTALLED_CARRIER)
    return INSTALLED_CARRIER
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/app/test_store.py::test_carrier_copies_to_install_dir_once -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add app/carrier.py tests/app/test_store.py
git commit -m "feat(app): bundled-carrier locate + first-run copy"
```

---

## Task 6: DeckController (PowerPoint COM)

**Files:**
- Create: `app/deck_controller.py`

No unit test (COM integration is exercised by the Task 8 core-loop harness — COM cannot be meaningfully mocked at unit level and the project's discipline is deterministic COM harnesses, spec §8).

- [ ] **Step 1: Write implementation**

`app/deck_controller.py`:
```python
"""PowerPoint COM. Two modes (spec D5):
  attach -> operate on the user's already-open ActivePresentation
  file   -> open a chosen .pptx, operate, Save
The carrier .pptm is opened (hidden) alongside so its macros
(ExecuteFromString / BuildSnapshotJson) are callable via app.Run.
Reuses the project's transient-COM retry discipline."""
import time

import pythoncom
import pywintypes
import win32com.client

from app.carrier import ensure_carrier

_TRANSIENT = (pywintypes.com_error, AttributeError)


def _open_app():
    last = None
    for _ in range(15):
        try:
            app = win32com.client.DispatchEx("PowerPoint.Application")
            app.Visible = True
            return app
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(2.0)
    raise RuntimeError(f"PowerPoint COM bring-up failed: {last!r}")


class NoPowerPointError(RuntimeError):
    pass


class NoOpenDeckError(RuntimeError):
    pass


class DeckController:
    def __init__(self):
        self.app = None
        self.carrier = None
        self.deck = None

    def start(self):
        pythoncom.CoInitialize()
        try:
            self.app = _open_app()
        except RuntimeError as e:
            raise NoPowerPointError(str(e))
        self.carrier = self.app.Presentations.Open(
            str(ensure_carrier()), WithWindow=False)

    def attach_open_deck(self):
        # The active deck must be a non-carrier presentation.
        for p in self.app.Presentations:
            if p.FullName != self.carrier.FullName:
                self.deck = p
                p.Windows(1).Activate()
                return
        raise NoOpenDeckError("No deck open in PowerPoint.")

    def open_file(self, path: str):
        self.deck = self.app.Presentations.Open(path, WithWindow=True)
        self.deck.Windows(1).Activate()

    def get_snapshot(self) -> str:
        return self._run("BuildSnapshotJson")

    def run_actions(self, actions_json: str) -> str:
        # ExecuteFromString runs the verify loop by default; returns a
        # human summary incl. "FAILURES (N)" contract.
        return self._run("ExecuteFromString", actions_json)

    def _run(self, macro: str, *args, _attempts: int = 3):
        last = None
        for i in range(1, _attempts + 1):
            try:
                return self.app.Run(f"PPT_AI_Editor!{macro}", *args)
            except _TRANSIENT as e:  # noqa: PERF203
                last = e
                time.sleep(2.0 * i)
        raise RuntimeError(f"{macro} failed after retries: {last!r}")

    def close(self, save_deck: bool = False):
        try:
            if self.deck is not None and save_deck:
                self.deck.Save()
        finally:
            for p in (self.carrier,):
                try:
                    if p is not None:
                        p.Saved = True
                        p.Close()
                except Exception:
                    pass
            try:
                self.app.Quit()
            except Exception:
                pass
            time.sleep(1.0)
```

- [ ] **Step 2: Syntax check**

Run: `python -c "import ast; ast.parse(open('app/deck_controller.py').read()); print('ok')"`
Expected: `ok`

- [ ] **Step 3: Commit**

```bash
git add app/deck_controller.py
git commit -m "feat(app): DeckController — COM attach/file modes + retry"
```

---

## Task 7: Stub LLM + Orchestrator

**Files:**
- Create: `tests/app/stub_llm.py`
- Create: `app/orchestrator.py`
- Test: `tests/app/test_store.py` (append — orchestrator with a fake deck + stub LLM, no COM)

- [ ] **Step 1: Write the failing test (append to tests/app/test_store.py)**

```python
def test_orchestrator_turn_persists_and_summarizes(tmp_path):
    from app.orchestrator import ChatOrchestrator
    from app.store import Store
    from tests.app.stub_llm import StubLLM

    class FakeDeck:
        def __init__(self):
            self.ran = None
        def get_snapshot(self):
            return "{SNAP}"
        def run_actions(self, j):
            self.ran = j
            return "applied 1, skipped 0"

    st = Store(tmp_path / "o.db"); st.init()
    fd = FakeDeck()
    orch = ChatOrchestrator(deck=fd, llm=StubLLM(), store=st)
    res = orch.run("make a title slide")
    assert res["summary"] == "applied 1, skipped 0"
    assert '"actions"' in fd.ran                         # actions reached deck
    rows = st.list_turns()
    assert len(rows) == 1 and rows[0]["request"] == "make a title slide"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/app/test_store.py::test_orchestrator_turn_persists_and_summarizes -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'app.orchestrator'`

- [ ] **Step 3: Write the stub and implementation**

`tests/app/stub_llm.py`:
```python
"""Deterministic LLM stub. No network. Returns a fixed valid batch so
the core loop is gated without a live model (spec §8)."""

FIXED_ACTIONS = (
    '{"actions":[{"type":"apply_template","template":"title",'
    '"content":{"title":"Stub Title","subtitle":"Stub Sub"}}]}'
)


class StubLLM:
    def call(self, snapshot: str, user_request: str) -> str:
        return FIXED_ACTIONS
```

`app/orchestrator.py`:
```python
"""One chat turn: text -> snapshot -> LLM -> actions -> engine ->
summary -> persist. `deck` and `llm` are duck-typed so this is unit
testable without COM (spec §6)."""
import re


class ChatOrchestrator:
    def __init__(self, deck, llm, store):
        self.deck = deck
        self.llm = llm
        self.store = store

    @staticmethod
    def _warn_count(summary: str) -> int:
        m = re.search(r"(\d+)\s+warning", summary)
        return int(m.group(1)) if m else 0

    def run(self, text: str) -> dict:
        snapshot = self.deck.get_snapshot()
        actions = self.llm.call(snapshot=snapshot, user_request=text)
        summary = self.deck.run_actions(actions)
        warnings = self._warn_count(summary)
        tid = self.store.add_turn(request=text, actions_json=actions,
                                  result_summary=summary, warnings=warnings)
        return {"id": tid, "summary": summary, "warnings": warnings,
                "actions_json": actions}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/app/test_store.py::test_orchestrator_turn_persists_and_summarizes -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add tests/app/stub_llm.py app/orchestrator.py tests/app/test_store.py
git commit -m "feat(app): ChatOrchestrator + deterministic LLM stub"
```

---

## Task 8: Core-loop COM harness

**Files:**
- Create: `tests/run_smoke_app_core_loop.py`

- [ ] **Step 1: Write the harness (it IS the test)**

`tests/run_smoke_app_core_loop.py`:
```python
"""core_loop: snapshot -> stub-LLM -> ExecuteFromString -> assert the
engine summary. No network. Single isolated PowerPoint run with the
project's transient-COM retry. Exit 0 only on PASS."""
import os
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from app.deck_controller import DeckController          # noqa: E402
from tests.app.stub_llm import StubLLM                  # noqa: E402
from app.orchestrator import ChatOrchestrator           # noqa: E402
from app.store import Store                             # noqa: E402


def run_once() -> list[str]:
    fails: list[str] = []
    dc = DeckController()
    dc.start()
    try:
        pres = dc.app.Presentations.Add()
        pres.PageSetup.SlideWidth = 960
        pres.PageSetup.SlideHeight = 540
        pres.Windows(1).Activate()
        time.sleep(1.0)
        dc.deck = pres

        st = Store(Path(os.environ["TEMP"]) / "decko_coreloop.db")
        if st.db_path.exists():
            st.db_path.unlink()
        st.init()
        orch = ChatOrchestrator(deck=dc, llm=StubLLM(), store=st)
        res = orch.run("make a title slide")

        s = res["summary"] or ""
        if "applied" not in s.lower():
            fails.append(f"summary missing 'applied': {s!r}")
        if "FAILURES" in s:
            fails.append(f"engine reported FAILURES: {s!r}")
        if pres.Slides.Count < 1:
            fails.append("no slide created by apply_template")
        if len(st.list_turns()) != 1:
            fails.append("turn not persisted")
        try:
            pres.Saved = True
            pres.Close()
        except Exception:
            pass
        return fails
    finally:
        dc.close(save_deck=False)


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
        print(f"FAIL: core_loop failed after retries: {last!r}")
        return 1
    if fails:
        for f in fails:
            print(f"  FAIL [core_loop] {f}")
        print("\nRESULT: FAIL")
        return 1
    print("  core_loop: snapshot -> stub LLM -> ExecuteFromString OK")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run it**

Run: `taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2; python tests/run_smoke_app_core_loop.py`
Expected: `RESULT: PASS`, exit 0
(If FAIL: read the printed reason. Do NOT modify `src/` to pass — fix the app layer.)

- [ ] **Step 3: Commit**

```bash
git add tests/run_smoke_app_core_loop.py
git commit -m "test(app): deterministic core-loop COM harness (stub LLM)"
```

---

## Task 9: Packaging (bake carrier, PyInstaller, packaging smoke)

**Files:**
- Create: `packaging/build.py`
- Create: `packaging/decko.spec`
- Create: `packaging/installer.iss`
- Create: `tests/run_smoke_app_packaging.py`

- [ ] **Step 1: Write packaging/build.py**

```python
"""Build pipeline. Bakes current src/ into the carrier via the existing
update_macros.py (spec §4 carrier provenance), then runs PyInstaller."""
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def main() -> int:
    print("[build] baking carrier via update_macros.py")
    r = subprocess.run([sys.executable, "update_macros.py"], cwd=REPO)
    if r.returncode != 0:
        print("[build] update_macros failed")
        return 1
    print("[build] PyInstaller")
    r = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm",
         str(REPO / "packaging" / "decko.spec")], cwd=REPO)
    return r.returncode


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Write packaging/decko.spec**

```python
# PyInstaller spec for Decko Desktop. Bundles the carrier + web/ assets.
from pathlib import Path

REPO = Path(SPECPATH).parent

a = Analysis(
    [str(REPO / "app" / "main.py")],
    pathex=[str(REPO)],
    binaries=[],
    datas=[
        (str(REPO / "PPT_AI_Editor.pptm"), "."),
        (str(REPO / "app" / "web"), "app/web"),
    ],
    hiddenimports=["win32com", "win32com.client", "keyring.backends.Windows"],
    hookspath=[], runtime_hooks=[], excludes=[],
)
pyz = PYZ(a.pure)
exe = EXE(pyz, a.scripts, a.binaries, a.datas, [],
          name="Decko", console=False)
```

- [ ] **Step 3: Write packaging/installer.iss**

```ini
[Setup]
AppName=Decko
AppVersion=1.0.0
DefaultDirName={autopf}\Decko
DefaultGroupName=Decko
OutputBaseFilename=Decko-Setup
Compression=lzma2
SolidCompression=yes

[Files]
Source: "..\dist\Decko.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Decko"; Filename: "{app}\Decko.exe"
Name: "{commondesktop}\Decko"; Filename: "{app}\Decko.exe"
```

- [ ] **Step 4: Write tests/run_smoke_app_packaging.py**

```python
"""packaging_smoke: PyInstaller build succeeds AND the bundled carrier
loads via COM from the packaged path. Exit 0 only on PASS."""
import os
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def main() -> int:
    r = subprocess.run([sys.executable, "packaging/build.py"], cwd=REPO)
    if r.returncode != 0:
        print("FAIL: packaging build failed")
        return 1
    exe = REPO / "dist" / "Decko.exe"
    if not exe.exists():
        print(f"FAIL: {exe} missing")
        return 1
    # The .exe must contain a usable carrier: smoke it by loading the
    # repo carrier via COM (same bytes that get bundled).
    os.system("taskkill /F /IM POWERPNT.EXE >NUL 2>&1")
    time.sleep(2.0)
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True
    try:
        p = app.Presentations.Open(str(REPO / "PPT_AI_Editor.pptm"),
                                    WithWindow=True)
        n = len(app.Run("PPT_AI_Editor!GetAllActionTypes"))
        p.Saved = True
        p.Close()
    finally:
        app.Quit()
        time.sleep(1.0)
    if n < 1000:
        print(f"FAIL: GetAllActionTypes too short ({n})")
        return 1
    print(f"  packaging_smoke: Decko.exe built; carrier OK (len={n})")
    print("\nRESULT: PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 5: Run packaging smoke**

Run: `python tests/run_smoke_app_packaging.py`
Expected: `RESULT: PASS`, exit 0

- [ ] **Step 6: Commit**

```bash
git add packaging/build.py packaging/decko.spec packaging/installer.iss tests/run_smoke_app_packaging.py
git commit -m "feat(app): packaging pipeline + packaging_smoke harness"
```

---

## Task 10: Aggregator (the metric)

**Files:**
- Create: `tests/run_smoke_app.py`

- [ ] **Step 1: Write the aggregator**

`tests/run_smoke_app.py`:
```python
"""SP1 metric. Runs the four SP1 gates; exit 0 only at 100%.
  store_unit    : pytest tests/app/test_store.py
  llmclient_unit: pytest tests/app/test_llm_client.py
  core_loop     : tests/run_smoke_app_core_loop.py   (COM, stub LLM)
  packaging_smoke: tests/run_smoke_app_packaging.py  (PyInstaller + COM)
UI layer is NOT gated (manual screenshot, spec §8)."""
import subprocess
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent

GATES = [
    ("store_unit",     [sys.executable, "-m", "pytest", "-q",
                         "tests/app/test_store.py"]),
    ("llmclient_unit", [sys.executable, "-m", "pytest", "-q",
                         "tests/app/test_llm_client.py"]),
    ("core_loop",      [sys.executable, "tests/run_smoke_app_core_loop.py"]),
    ("packaging_smoke",[sys.executable, "tests/run_smoke_app_packaging.py"]),
]


def main() -> int:
    failed = []
    for name, cmd in GATES:
        print(f"=== {name} ===")
        if subprocess.run(cmd, cwd=REPO).returncode != 0:
            failed.append(name)
    print("\nRESULT:", "PASS" if not failed else f"FAIL {failed}")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Run the full metric**

Run: `taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2; python tests/run_smoke_app.py`
Expected: `RESULT: PASS`, exit 0

- [ ] **Step 3: Run the engine regression guard (must stay green)**

Run each, expect PASS:
```bash
python tests/run_smoke.py
python tests/run_smoke_icon_prompt.py
python tests/run_smoke_spec.py
python tests/run_smoke_capture.py
python tests/run_smoke_dialogs.py
```

- [ ] **Step 4: Commit**

```bash
git add tests/run_smoke_app.py
git commit -m "test(app): run_smoke_app aggregator — SP1 metric (100% gate)"
```

---

## Task 11: UI layer (pywebview) — NOT deterministically gated

**Files:**
- Create: `app/main.py`
- Create: `app/web/index.html`
- Create: `app/web/app.css`
- Create: `app/web/app.js`

Per spec §8 the UI is verified by **manual screenshot review**, not the
deterministic gate. No TDD steps; this task is "build + eyeball".

- [ ] **Step 1: Write app/main.py**

```python
"""pywebview window + js_api bridge. Wires the side panel + chat to the
Python core. Not in the deterministic gate (spec §8)."""
import sys
import webview

from app import secrets
from app.config import DB_PATH, Settings, ensure_app_dirs
from app.deck_controller import (DeckController, NoOpenDeckError,
                                 NoPowerPointError)
from app.llm_client import LLMClient
from app.orchestrator import ChatOrchestrator
from app.store import Store


class Api:
    def __init__(self):
        self.settings = Settings()
        self.store = Store(DB_PATH)
        self.dc = None
        self.orch = None

    def boot(self):
        ensure_app_dirs()
        self.store.init()
        return {"has_key": secrets.has_api_key(),
                "history": self.store.list_turns()}

    def save_settings(self, provider, model, base_url, api_key):
        self.settings = Settings(provider=provider, model=model,
                                 base_url=base_url or "")
        self.settings.validate()
        if api_key:
            secrets.set_api_key(api_key)
        return {"ok": True}

    def open_session(self, mode, file_path=""):
        self.dc = DeckController()
        try:
            self.dc.start()
            if mode == "attach":
                self.dc.attach_open_deck()
            else:
                self.dc.open_file(file_path)
        except NoPowerPointError:
            return {"error": "Microsoft PowerPoint is required and was "
                             "not found. Install PowerPoint and retry."}
        except NoOpenDeckError:
            return {"error": "No deck open in PowerPoint. Open one, or "
                             "choose 'Open file' instead."}
        llm = LLMClient(self.settings, secrets.get_api_key() or "")
        self.orch = ChatOrchestrator(self.dc, llm, self.store)
        return {"ok": True}

    def send(self, text):
        if self.orch is None:
            return {"error": "Start a session first."}
        try:
            return self.orch.run(text)
        except Exception as e:  # noqa: BLE001
            return {"error": str(e)}

    def shutdown(self):
        if self.dc is not None:
            self.dc.close(save_deck=False)
        return {"ok": True}


def main():
    api = Api()
    webview.create_window("Decko", "app/web/index.html",
                          js_api=api, width=1100, height=720)
    webview.start()
    api.shutdown()


if __name__ == "__main__":
    sys.exit(main() or 0)
```

- [ ] **Step 2: Write app/web/index.html**

```html
<!doctype html><html><head><meta charset="utf-8">
<link rel="stylesheet" href="app.css"></head><body>
<div id="layout">
  <aside id="side">
    <h2>Decko</h2>
    <section><h3>Deck</h3>
      <select id="mode"><option value="attach">Attach to open deck</option>
        <option value="file">Open a .pptx file</option></select>
      <input id="file" placeholder="C:\\path\\deck.pptx" hidden>
      <button id="startBtn">Start session</button></section>
    <section><h3>LLM</h3>
      <select id="provider"><option>anthropic</option><option>openai</option>
        <option>generic</option></select>
      <input id="model" placeholder="model" value="claude-opus-4-7">
      <input id="baseUrl" placeholder="base URL (generic only)" hidden>
      <input id="apiKey" type="password" placeholder="API key">
      <button id="saveBtn">Save settings</button></section>
    <section><h3>History</h3><ul id="history"></ul></section>
  </aside>
  <main id="chat">
    <div id="thread"></div>
    <div id="composer">
      <textarea id="msg" placeholder="Describe the change..."></textarea>
      <button id="sendBtn">Send</button></div>
  </main>
</div><script src="app.js"></script></body></html>
```

- [ ] **Step 3: Write app/web/app.css**

```css
*{box-sizing:border-box;font-family:Segoe UI,system-ui,sans-serif}
body{margin:0;background:#0d0d0f;color:#eee}
#layout{display:flex;height:100vh}
#side{width:300px;background:#161618;padding:16px;overflow:auto}
#side h2{margin:.2rem 0 1rem}#side h3{margin:1rem 0 .4rem;font-size:.8rem;
 text-transform:uppercase;color:#888}
#side input,#side select,#side button,#composer button,#msg{width:100%;
 margin:.25rem 0;padding:.5rem;background:#222;color:#eee;border:1px solid #333;
 border-radius:6px}
#chat{flex:1;display:flex;flex-direction:column}
#thread{flex:1;overflow:auto;padding:18px}
.bubble{margin:.5rem 0;padding:.6rem .8rem;border-radius:8px;max-width:80%}
.user{background:#234;margin-left:auto}.app{background:#1d1d20}
.warn{color:#e6b800}.fail{color:#ff6b6b}
#composer{display:flex;gap:8px;padding:12px;border-top:1px solid #333}
#msg{height:64px;resize:none}#composer button{width:120px}
</style>
```
(Note: the trailing `</style>` is HTML-injected by browsers tolerant of
the stylesheet link; keep the file as pure CSS — delete the `</style>`
line before saving. It is shown only to mark file end.)

- [ ] **Step 4: Write app/web/app.js**

```javascript
const $ = (id) => document.getElementById(id);
let api;

function bubble(kind, html) {
  const d = document.createElement("div");
  d.className = "bubble " + kind;
  d.innerHTML = html;
  $("thread").appendChild(d);
  $("thread").scrollTop = $("thread").scrollHeight;
}

window.addEventListener("pywebviewready", async () => {
  api = window.pywebview.api;
  const s = await api.boot();
  s.history.forEach((h) =>
    bubble("app", `<b>${h.request}</b><br>${h.result_summary}`));
  if (!s.has_key) bubble("app", "Set your API key in the side panel.");
});

$("provider").onchange = (e) =>
  ($("baseUrl").hidden = e.target.value !== "generic");
$("mode").onchange = (e) =>
  ($("file").hidden = e.target.value !== "file");

$("saveBtn").onclick = async () => {
  const r = await api.save_settings($("provider").value, $("model").value,
    $("baseUrl").value, $("apiKey").value);
  bubble("app", r.ok ? "Settings saved." : "Error: " + r.error);
};

$("startBtn").onclick = async () => {
  const r = await api.open_session($("mode").value, $("file").value);
  bubble("app", r.ok ? "Session started." : "fail: " + r.error);
};

$("sendBtn").onclick = async () => {
  const t = $("msg").value.trim();
  if (!t) return;
  bubble("user", t);
  $("msg").value = "";
  const r = await api.send(t);
  if (r.error) { bubble("app fail", r.error); return; }
  const w = r.warnings ? ` <span class="warn">(${r.warnings} warnings)</span>`
                       : "";
  bubble("app", r.summary + w);
};
```

- [ ] **Step 5: Manual verification (no automated gate)**

Run: `python -m app.main`
Check by eye: window opens; side panel + chat render; saving settings
works; with a deck open, "attach" + a request runs a turn and shows the
engine summary. Capture a screenshot for the record.

- [ ] **Step 6: Commit**

```bash
git add app/main.py app/web/index.html app/web/app.css app/web/app.js
git commit -m "feat(app): pywebview UI — chat + side panel (manual-verified)"
```

---

## Task 12: Final gate + handoff doc

**Files:**
- Modify: `README.md` (add a "Desktop app (SP1)" section pointing at the installer)

- [ ] **Step 1: Run the full SP1 metric + engine guard**

```bash
taskkill //F //IM POWERPNT.EXE 2>/dev/null; sleep 2
python tests/run_smoke_app.py            # expect RESULT: PASS
python tests/run_smoke.py                # expect all tests passed
python tests/run_smoke_icon_prompt.py    # expect PASS
python tests/run_smoke_spec.py           # expect PASS
python tests/run_smoke_capture.py        # expect PASS
python tests/run_smoke_dialogs.py        # expect PASS
```

- [ ] **Step 2: Add README section**

Add under the existing documentation list in `README.md`:
```markdown
## Desktop app (SP1)

`Decko-Setup.exe` (built via `python packaging/build.py` →
`packaging/installer.iss`) is a Windows installer requiring only
Microsoft PowerPoint — no Python. It wraps the same VBA/COM engine in a
chat UI; users supply their own LLM key (Anthropic/OpenAI/generic).
Design: `docs/superpowers/specs/2026-05-17-decko-desktop-app-design.md`.
```

- [ ] **Step 3: Commit**

```bash
git add README.md
git commit -m "docs: README — Desktop app (SP1) section"
```

---

## Self-Review

**Spec coverage:** D1 (COM reused — Tasks 6/8/9, no `src/` change). D2 (Windows-only — COM in T6). D3 (BYO key — T3/T4). D4 (local now, sync-ready — T2 schema). D5 (both modes — T6 `attach_open_deck`/`open_file`). D6 (3 providers — T4). D7 (chat+side panel — T11). D8 (pywebview + PyInstaller — T9/T11). §6 data flow — T7 orchestrator. §7 error handling — `NoPowerPointError`/`NoOpenDeckError` (T6), key/timeout/sanitize (T4), FAILURES surfaced via summary (T7/T8). §8 testing — T2/T4 unit, T8 core_loop, T9 packaging_smoke, T10 aggregator; UI manual T11. §9 metric — T10 `run_smoke_app.py` + engine guard T10/T12. §4 carrier provenance — T9 `build.py` runs `update_macros.py`. All covered.

**Placeholder scan:** No TBD/TODO. Every code step has complete code. The `app.css` step carries an explicit note that the shown trailing `</style>` is not part of the saved file — not a placeholder, an instruction.

**Type consistency:** `DeckController.get_snapshot/run_actions/close(save_deck=)`, `ChatOrchestrator(deck,llm,store).run()→{id,summary,warnings,actions_json}`, `Store.add_turn(...)→str`/`list_turns(include_deleted=)`, `LLMClient(settings,api_key,_transport).call(snapshot,user_request)→str`, `secrets.set/get/has/clear_api_key` — names consistent across Tasks 2–11. Stub `StubLLM.call(snapshot,user_request)` matches `LLMClient.call` signature used by orchestrator.

No gaps found.

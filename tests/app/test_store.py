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
    assert '"actions"' in fd.ran
    rows = st.list_turns()
    assert len(rows) == 1 and rows[0]["request"] == "make a title slide"


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

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

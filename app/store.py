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

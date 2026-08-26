"""SQLite source of truth for the persistent academic workbench."""

from __future__ import annotations

import json
import re
import shutil
import sqlite3
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


_SENSITIVE_KEY = re.compile(
    r"(?:password|passwd|secret|token|cookie|authorization|credential|"
    r"student[_ -]?(?:number|no|name)|real[_ -]?name|response[_ -]?(?:html|body)|"
    r"page[_ -]?source|raw[_ -]?(?:html|page))",
    re.IGNORECASE,
)
_SENSITIVE_INLINE = re.compile(
    r"(?i)(?:cookie|set-cookie|authorization|bearer|password|passwd|token|"
    r"student[_ -]?(?:number|no|name))\s*[:=]\s*[^,;\s]+"
)
_HTML_TAG = re.compile(r"<\/?[a-z][^>]*>", re.IGNORECASE)


def sanitize_for_storage(value: Any, *, key: str = "") -> Any:
    """Keep structured diagnostics useful without persisting page/auth data."""
    if _SENSITIVE_KEY.search(key):
        return "[redacted]"
    if isinstance(value, dict):
        return {str(name): sanitize_for_storage(item, key=str(name)) for name, item in value.items()}
    if isinstance(value, list):
        return [sanitize_for_storage(item) for item in value]
    if isinstance(value, tuple):
        return [sanitize_for_storage(item) for item in value]
    if isinstance(value, str):
        cleaned = _HTML_TAG.sub(" ", value)
        return _SENSITIVE_INLINE.sub("[redacted]", cleaned)
    return value


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


SCHEMA = """
create table if not exists schema_migrations(version integer primary key, applied_at text not null);
create table if not exists app_metadata(key text primary key, value text not null);
create table if not exists profiles(id text primary key, created_at text not null, payload text not null);
create table if not exists notice_versions(id text primary key, source_url text not null default '', body_hash text not null default '', status text not null, created_at text not null, payload text not null);
create table if not exists snapshots(id text primary key, kind text not null, term text not null, profile_id text, notice_id text, source text not null, source_at text not null, created_at text not null, payload text not null, complete integer not null check(complete in (0,1)));
create index if not exists snapshots_kind_created on snapshots(kind, created_at desc);
create table if not exists plans(id text primary key, term text not null, profile_id text not null, notice_id text not null, timetable_snapshot_id text not null, selection_snapshot_id text not null, created_at text not null, payload text not null);
create index if not exists plans_created on plans(created_at desc);
create table if not exists refresh_attempts(id text primary key, kind text not null, state text not null, started_at text not null, finished_at text, error text not null default '', snapshot_id text references snapshots(id));
create table if not exists observation_tasks(id text primary key, operation text not null, coalesce_key text not null, state text not null, created_at text not null, updated_at text not null, progress text not null, context text not null, error text not null default '');
"""


class WorkspaceDatabase:
    def __init__(self, root: Path, connection: sqlite3.Connection):
        self.root = root
        self.connection = connection
        self.connection.row_factory = sqlite3.Row

    @classmethod
    def open(cls, root: Path | str) -> "WorkspaceDatabase":
        root = Path(root)
        root.mkdir(parents=True, exist_ok=True)
        path = root / "workbench.sqlite3"
        existed = path.exists()
        if existed:
            shutil.copy2(path, root / "workbench.sqlite3.backup")
        connection = sqlite3.connect(path, timeout=5, check_same_thread=False)
        try:
            connection.execute("pragma foreign_keys=on")
            connection.execute("pragma busy_timeout=5000")
            connection.execute("pragma journal_mode=wal")
            with connection:
                connection.executescript(SCHEMA)
                connection.execute(
                    "insert or ignore into schema_migrations(version, applied_at) values(1, ?)",
                    (utc_now(),),
                )
            database = cls(root, connection)
            database._import_legacy_once()
            return database
        except Exception:
            connection.close()
            if existed and (root / "workbench.sqlite3.backup").exists():
                shutil.copy2(root / "workbench.sqlite3.backup", path)
            raise

    def _import_legacy_once(self) -> None:
        done = self.connection.execute(
            "select value from app_metadata where key='legacy_imported'"
        ).fetchone()
        if done:
            return
        candidates = {
            "profile": ("student-profile.json",),
            "notice": ("selection-notice.json",),
            "timetable": ("current-timetable.json",),
            "selection": ("selection-entry.json", "selection-query.json"),
        }
        with self.connection:
            profile_id = None
            notice_id = None
            for kind, names in candidates.items():
                path = next((self.root / name for name in names if (self.root / name).is_file()), None)
                if path is None:
                    continue
                payload = json.loads(path.read_text(encoding="utf-8"))
                if kind == "profile":
                    profile_id = self._insert_profile(payload)
                elif kind == "notice":
                    notice_id = self._insert_notice(payload)
                else:
                    self._insert_snapshot(
                        kind,
                        payload.get("term", ""),
                        payload,
                        "legacy-json",
                        profile_id=profile_id,
                        notice_id=notice_id if kind == "selection" else None,
                    )
            self.connection.execute(
                "insert into app_metadata(key,value) values('legacy_imported',?)", (utc_now(),)
            )

    def _insert_profile(self, payload: dict[str, Any]) -> str:
        identity = uuid.uuid4().hex
        self.connection.execute(
            "insert into profiles values(?,?,?)", (identity, utc_now(), json.dumps(sanitize_for_storage(payload), ensure_ascii=False))
        )
        return identity

    def _insert_notice(self, payload: dict[str, Any]) -> str:
        identity = payload.get("version_id") or uuid.uuid4().hex
        self.connection.execute(
            "insert or ignore into notice_versions values(?,?,?,?,?,?)",
            (identity, payload.get("source_url", ""), payload.get("content_hash", ""), payload.get("status", "candidate"), utc_now(), json.dumps(sanitize_for_storage(payload), ensure_ascii=False)),
        )
        return identity

    def _insert_snapshot(self, kind: str, term: str, payload: dict[str, Any], source: str, profile_id: str | None = None, notice_id: str | None = None) -> str:
        identity = uuid.uuid4().hex
        now = utc_now()
        self.connection.execute(
            "insert into snapshots values(?,?,?,?,?,?,?,?,?,1)",
            (identity, kind, term, profile_id, notice_id, source, payload.get("source_at", now), now, json.dumps(sanitize_for_storage(payload), ensure_ascii=False)),
        )
        return identity

    def current_profile(self) -> dict[str, Any] | None:
        row = self.connection.execute("select id,payload,created_at from profiles order by created_at desc limit 1").fetchone()
        return ({**json.loads(row["payload"]), "version_id": row["id"]} if row else None)

    def confirmed_notice(self) -> dict[str, Any] | None:
        row = self.connection.execute("select id,payload from notice_versions where status='confirmed' order by created_at desc limit 1").fetchone()
        return ({**json.loads(row["payload"]), "version_id": row["id"]} if row else None)

    def save_notice(self, payload: dict[str, Any], *, confirmed: bool = False) -> dict[str, Any]:
        copy = dict(payload)
        copy["status"] = "confirmed" if confirmed else "candidate"
        with self.connection:
            if confirmed:
                self.connection.execute("update notice_versions set status='candidate' where status='confirmed'")
            identity = self._insert_notice(copy)
        return {**copy, "version_id": identity}

    def notice(self, identity: str) -> dict[str, Any] | None:
        row = self.connection.execute("select id,payload,status from notice_versions where id=?", (identity,)).fetchone()
        return ({**json.loads(row["payload"]), "version_id": row["id"], "status": row["status"]} if row else None)

    def confirm_notice(self, identity: str) -> dict[str, Any]:
        notice = self.notice(identity)
        if not notice:
            raise ValueError("candidate notice not found")
        if not notice.get("query_eligible"):
            raise ValueError("candidate notice is incomplete and cannot drive queries")
        with self.connection:
            self.connection.execute("update notice_versions set status='candidate' where status='confirmed'")
            self.connection.execute("update notice_versions set status='confirmed' where id=?", (identity,))
        return {**notice, "status": "confirmed"}

    def publish_snapshot(self, kind: str, term: str, payload: dict[str, Any], *, source: str, profile_id: str | None = None, notice_id: str | None = None) -> dict[str, Any]:
        with self.connection:
            identity = self._insert_snapshot(kind, term, payload, source, profile_id, notice_id)
            self.connection.execute("insert into refresh_attempts values(?,?,?,?,?,?,?)", (uuid.uuid4().hex, kind, "succeeded", utc_now(), utc_now(), "", identity))
            # Timestamps have second precision, so rowid breaks ties for rapid refreshes.
            old = self.connection.execute("select id from snapshots where kind=? order by created_at desc, rowid desc limit -1 offset 20", (kind,)).fetchall()
            # Keep refresh diagnostics while allowing retention to remove the referenced snapshot.
            self.connection.executemany("update refresh_attempts set snapshot_id=null where snapshot_id=?", ((row[0],) for row in old))
            self.connection.executemany("delete from snapshots where id=?", ((row[0],) for row in old))
        return self.snapshot(identity)

    def record_failed_attempt(self, kind: str, error: str) -> None:
        with self.connection:
            safe_error = sanitize_for_storage(str(error))
            self.connection.execute("insert into refresh_attempts values(?,?,?,?,?,?,null)", (uuid.uuid4().hex, kind, "failed", utc_now(), utc_now(), str(safe_error)[:1000]))

    def snapshot(self, identity: str) -> dict[str, Any]:
        row = self.connection.execute("select * from snapshots where id=?", (identity,)).fetchone()
        return self._snapshot_dict(row)

    def latest_snapshot(self, kind: str) -> dict[str, Any] | None:
        row = self.connection.execute("select * from snapshots where kind=? and complete=1 order by created_at desc limit 1", (kind,)).fetchone()
        return self._snapshot_dict(row) if row else None

    def save_plan(self, payload: dict[str, Any]) -> dict[str, Any]:
        identity = uuid.uuid4().hex
        now = utc_now()
        safe_payload = sanitize_for_storage(payload)
        with self.connection:
            self.connection.execute(
                "insert into plans values(?,?,?,?,?,?,?,?)",
                (identity, str(payload.get("term", "")), str(payload.get("profile_id", "")),
                 str(payload.get("notice_id", "")), str(payload.get("timetable_snapshot_id", "")),
                 str(payload.get("selection_snapshot_id", "")), now,
                 json.dumps(safe_payload, ensure_ascii=False)),
            )
        return {"id": identity, "created_at": now, **safe_payload}

    def latest_plan(self) -> dict[str, Any] | None:
        row = self.connection.execute("select id,created_at,payload from plans order by created_at desc, rowid desc limit 1").fetchone()
        if not row:
            return None
        return {"id": row["id"], "created_at": row["created_at"], **json.loads(row["payload"])}

    @staticmethod
    def _snapshot_dict(row: sqlite3.Row) -> dict[str, Any]:
        return {"id": row["id"], "kind": row["kind"], "term": row["term"], "profile_id": row["profile_id"], "notice_id": row["notice_id"], "source": row["source"], "source_at": row["source_at"], "created_at": row["created_at"], "payload": json.loads(row["payload"])}

    def close(self) -> None:
        self.connection.close()

"""CRUD audit logging backed by SQLite."""

from __future__ import annotations

import json
import logging
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from quan_ly_ho_so.state import current_database_path
from quan_ly_ho_so.utils.text import fold


ACTION_ALIASES: Dict[str, str] = {
    "tao moi": "create",
    "tao_moi": "create",
    "tao": "create",
    "cap nhat": "update",
    "cap_nhat": "update",
    "sua": "update",
    "xoa mem": "delete",
    "xoa_mem": "delete",
    "xoa": "delete",
    "khoi phuc": "restore",
    "khoi_phuc": "restore",
    "xuat file": "export",
    "xuat_file": "export",
    "xuat excel": "export",
    "xuat_excel": "export",
    "xuat": "export",
}


def escaped_like(value: str) -> str:
    """Escape special characters in SQL LIKE pattern."""
    return str(value).replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def _normalize_date(val: Optional[str]) -> Optional[str]:
    """Normalize date strings to YYYY-MM-DD format."""
    if not val:
        return None
    val = val.strip()
    if not val:
        return None
    if "/" in val:
        parts = val.split("/")
        if len(parts) == 3 and len(parts[2]) == 4:
            return f"{parts[2]}-{parts[1].zfill(2)}-{parts[0].zfill(2)}"
    if len(val) >= 10 and val[4] == "-" and val[7] == "-":
        return val[:10]
    return val


def _audit_database_path() -> Path:
    """Audit log lives next to the workbook cache database."""
    return current_database_path().with_name("audit.sqlite3")


def _audit_connection() -> sqlite3.Connection:
    path = _audit_database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.create_function("fold", 1, fold)
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS audit_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT NOT NULL,
            username TEXT NOT NULL,
            action TEXT NOT NULL,
            stt TEXT NOT NULL DEFAULT '',
            record_name TEXT NOT NULL DEFAULT '',
            details TEXT NOT NULL DEFAULT '{}'
        )
        """
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp DESC)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_audit_stt ON audit_log(stt)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_audit_username ON audit_log(username)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_log(action)"
    )
    connection.commit()
    return connection


def log_action(
    username: str,
    action: str,
    stt: str = "",
    record_name: str = "",
    details: Optional[Dict[str, Any]] = None,
) -> None:
    """Record a CRUD action in the audit log.
    
    Args:
        username: Who performed the action.
        action: One of 'create', 'update', 'delete', 'restore', 'export', etc.
        stt: The record's STT identifier.
        record_name: The record's display name (TÊN THƯỜNG DÙNG).
        details: Optional dict with field-level change details.
                 For updates: {"changes": [{"field": "Tên", "from": "A", "to": "B"}, ...]}
                 For creates: {"fields": {"Tên": "value", ...}}
                 For deletes: {"record_name": "...", "reason": "soft_delete"}
    """
    try:
        connection = _audit_connection()
        try:
            connection.execute(
                """
                INSERT INTO audit_log (timestamp, username, action, stt, record_name, details)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    datetime.now().isoformat(),
                    username,
                    action,
                    stt,
                    record_name,
                    json.dumps(details or {}, ensure_ascii=False, separators=(",", ":")),
                ),
            )
            connection.commit()
        finally:
            connection.close()
    except Exception as error:
        logging.warning("Could not write audit log: %s", error)


def compute_changes(
    old_values: Dict[str, str],
    new_values: Dict[str, str],
    field_labels: Dict[str, str],
) -> List[Dict[str, str]]:
    """Compare old and new field values and return a list of changes.
    
    Args:
        old_values: Dict mapping column number (as string) to old value.
        new_values: Dict mapping column number (as string) to new value.
        field_labels: Dict mapping column number (as string) to human-readable label.
    
    Returns:
        List of dicts: [{"field": "label", "from": "old", "to": "new"}, ...]
    """
    changes = []
    for key in new_values:
        old = (old_values.get(key) or "").strip()
        new = (new_values.get(key) or "").strip()
        if old != new:
            label = field_labels.get(key, key)
            changes.append({"field": label, "from": old, "to": new})
    return changes


def get_audit_performers() -> List[str]:
    """Retrieve distinct usernames that have entries in the audit log or exist in the system."""
    usernames = set()
    try:
        from quan_ly_ho_so.auth.db import list_users
        for u in list_users():
            if u.get("username"):
                usernames.add(u["username"])
    except Exception:
        pass

    try:
        connection = _audit_connection()
        try:
            rows = connection.execute(
                "SELECT DISTINCT username FROM audit_log WHERE username != ''"
            ).fetchall()
            for row in rows:
                if row[0]:
                    usernames.add(row[0])
        finally:
            connection.close()
    except Exception:
        pass

    return sorted(usernames, key=lambda s: s.casefold())


def get_audit_logs(
    limit: int = 50,
    offset: int = 0,
    stt: Optional[str] = None,
    username: Optional[str] = None,
    action: Optional[str] = None,
    date_from: Optional[str] = None,
    date_to: Optional[str] = None,
    query: Optional[str] = None,
    record_name: Optional[str] = None,
) -> tuple[list[dict], int]:
    """Retrieve audit log entries, newest first, with filtering and search.
    
    Args:
        limit: Number of records to return (default 50).
        offset: Record offset for pagination.
        stt: Filter by record STT.
        username: Filter by performer username.
        action: Filter by action (e.g. create, update, delete, restore, export).
        date_from: Filter by date (YYYY-MM-DD), inclusive.
        date_to: Filter by date (YYYY-MM-DD), inclusive.
        query: Unified search across username (Người thực hiện), stt (STT), record_name (Họ tên).
        record_name: Specific filter/search for record_name (Họ tên).

    Returns:
        Tuple of (list of log entries, total count).
    """
    conditions: List[str] = []
    parameters: List[Any] = []

    clean_stt = str(stt).strip() if stt is not None and str(stt).strip() else None
    if clean_stt:
        conditions.append("stt = ?")
        parameters.append(clean_stt)

    clean_user = str(username).strip() if username is not None and str(username).strip() else None
    if clean_user:
        conditions.append("fold(username) = fold(?)")
        parameters.append(clean_user)

    clean_action = str(action).strip() if action is not None and str(action).strip() else None
    if clean_action:
        canonical = ACTION_ALIASES.get(fold(clean_action), clean_action.lower())
        conditions.append("(action = ? OR action = ?)")
        parameters.extend([clean_action, canonical])

    clean_date_from = _normalize_date(date_from)
    if clean_date_from:
        conditions.append("substr(timestamp, 1, 10) >= ?")
        parameters.append(clean_date_from)

    clean_date_to = _normalize_date(date_to)
    if clean_date_to:
        conditions.append("substr(timestamp, 1, 10) <= ?")
        parameters.append(clean_date_to)

    clean_record_name = str(record_name).strip() if record_name is not None and str(record_name).strip() else None
    if clean_record_name:
        conditions.append("fold(record_name) LIKE ? ESCAPE '\\'")
        parameters.append(f"%{escaped_like(fold(clean_record_name))}%")

    clean_query = str(query).strip() if query is not None and str(query).strip() else None
    if clean_query:
        q_like = f"%{escaped_like(fold(clean_query))}%"
        conditions.append(
            "(fold(username) LIKE ? ESCAPE '\\' OR fold(stt) LIKE ? ESCAPE '\\' OR fold(record_name) LIKE ? ESCAPE '\\')"
        )
        parameters.extend([q_like, q_like, q_like])

    where_clause = f" WHERE {' AND '.join(conditions)}" if conditions else ""

    connection = _audit_connection()
    try:
        count_query = f"SELECT COUNT(*) FROM audit_log{where_clause}"
        total = connection.execute(count_query, parameters).fetchone()[0]

        select_query = f"SELECT * FROM audit_log{where_clause} ORDER BY id DESC LIMIT ? OFFSET ?"
        rows = connection.execute(select_query, [*parameters, limit, offset]).fetchall()

        entries = []
        for row in rows:
            entry = dict(row)
            try:
                entry["details"] = json.loads(entry["details"])
            except (json.JSONDecodeError, TypeError):
                entry["details"] = {}
            entries.append(entry)
        return entries, total
    finally:
        connection.close()

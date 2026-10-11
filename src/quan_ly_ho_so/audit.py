"""CRUD audit logging backed by SQLite."""

from __future__ import annotations

import json
import logging
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

from quan_ly_ho_so.state import current_database_path


def _audit_database_path() -> Path:
    """Audit log lives next to the workbook cache database."""
    return current_database_path().with_name("audit.sqlite3")


def _audit_connection() -> sqlite3.Connection:
    path = _audit_database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
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
        action: One of 'create', 'update', 'delete'.
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


def get_audit_logs(limit: int = 100, offset: int = 0, stt: Optional[str] = None) -> tuple[list[dict], int]:
    """Retrieve audit log entries, newest first, optionally filtered by STT.
    
    Returns:
        Tuple of (list of log entries, total count).
    """
    connection = _audit_connection()
    try:
        if stt:
            total = connection.execute(
                "SELECT COUNT(*) FROM audit_log WHERE stt = ?", (str(stt),)
            ).fetchone()[0]
            rows = connection.execute(
                "SELECT * FROM audit_log WHERE stt = ? ORDER BY id DESC LIMIT ? OFFSET ?",
                (str(stt), limit, offset),
            ).fetchall()
        else:
            total = connection.execute("SELECT COUNT(*) FROM audit_log").fetchone()[0]
            rows = connection.execute(
                "SELECT * FROM audit_log ORDER BY id DESC LIMIT ? OFFSET ?",
                (limit, offset),
            ).fetchall()
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

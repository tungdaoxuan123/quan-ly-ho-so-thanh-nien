"""Pending approval workflow for profile edits and spam lockout management."""

from __future__ import annotations

import json
import logging
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from quan_ly_ho_so.audit import _audit_database_path, comparable_field_value, compute_changes, log_action
from quan_ly_ho_so.forms.fields import (
    coerce_cell_value,
    field_definitions,
    form_values,
)
from quan_ly_ho_so.state import state
from quan_ly_ho_so.workbook.manager import refresh_state
from quan_ly_ho_so.utils.text import display_value
from quan_ly_ho_so.workbook.cache import find_record_by_citizen_id, signature_token
from quan_ly_ho_so.workbook.writer import save_person

# Identity-critical fields that users are strictly prohibited from editing on /tra-cuu
IDENTITY_CRITICAL_HEADERS = {
    "TÊN THƯỜNG DÙNG",
    "Tên khai sinh",
    "Ngày sinh",
    "Tháng sinh",
    "Năm sinh",
    "Căn cước",
}


def is_identity_critical(header: str) -> bool:
    """Return True if header is an immutable identity-critical field."""
    return header.strip() in IDENTITY_CRITICAL_HEADERS


def _request_source() -> Tuple[str, str]:
    workbook = state["workbook"]
    return (str(Path(workbook).resolve()) if workbook else "", state["sheet"] or "")


def _approvals_connection() -> sqlite3.Connection:
    """Open or create connection to the approvals & audit SQLite database."""
    path = _audit_database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS profile_edit_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            cccd TEXT NOT NULL,
            workbook TEXT NOT NULL DEFAULT '',
            sheet TEXT NOT NULL DEFAULT '',
            stt TEXT NOT NULL DEFAULT '',
            record_name TEXT NOT NULL DEFAULT '',
            status TEXT NOT NULL DEFAULT 'pending',
            submitted_at TEXT NOT NULL,
            submitted_data TEXT NOT NULL DEFAULT '{}',
            original_data TEXT NOT NULL DEFAULT '{}',
            changes TEXT NOT NULL DEFAULT '[]',
            reviewer TEXT NOT NULL DEFAULT '',
            reviewed_at TEXT NOT NULL DEFAULT '',
            reviewer_note TEXT NOT NULL DEFAULT ''
        )
        """
    )
    # Old requests have no trustworthy source. Keep their history, but never infer a
    # workbook from whichever file happens to be open when the upgrade runs.
    columns = {row["name"] for row in connection.execute("PRAGMA table_info(profile_edit_requests)")}
    for column in ("workbook", "sheet"):
        if column not in columns:
            connection.execute(f"ALTER TABLE profile_edit_requests ADD COLUMN {column} TEXT NOT NULL DEFAULT ''")
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_requests_cccd_status ON profile_edit_requests(cccd, status)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_requests_status ON profile_edit_requests(status)"
    )
    connection.execute(
        "CREATE INDEX IF NOT EXISTS idx_requests_submitted ON profile_edit_requests(submitted_at DESC)"
    )
    connection.commit()
    return connection


def _row_to_request(row: sqlite3.Row) -> Dict[str, Any]:
    req = dict(row)
    for json_field in ("submitted_data", "original_data", "changes"):
        try:
            req[json_field] = json.loads(req.get(json_field) or ("[]" if json_field == "changes" else "{}"))
        except (json.JSONDecodeError, TypeError):
            req[json_field] = [] if json_field == "changes" else {}
    return req


def has_pending_request(cccd: str) -> bool:
    """Check if there is an active pending update request for the given CCCD (Spam prevention)."""
    clean_cccd = str(cccd).strip()
    if not clean_cccd:
        return False
    connection = _approvals_connection()
    try:
        row = connection.execute(
            "SELECT COUNT(*) FROM profile_edit_requests WHERE cccd = ? AND workbook = ? AND sheet = ? AND status = 'pending'",
            (clean_cccd, *_request_source()),
        ).fetchone()
        return bool(row and row[0] > 0)
    finally:
        connection.close()


def get_pending_request_by_cccd(cccd: str) -> Optional[Dict[str, Any]]:
    """Retrieve the pending request for the given CCCD if one exists."""
    clean_cccd = str(cccd).strip()
    if not clean_cccd:
        return None
    connection = _approvals_connection()
    try:
        row = connection.execute(
            "SELECT * FROM profile_edit_requests WHERE cccd = ? AND workbook = ? AND sheet = ? AND status = 'pending' ORDER BY id DESC LIMIT 1",
            (clean_cccd, *_request_source()),
        ).fetchone()
        return _row_to_request(row) if row else None
    finally:
        connection.close()


def get_latest_request_by_cccd(cccd: str) -> Optional[Dict[str, Any]]:
    """Retrieve the most recent request for the given CCCD (any status)."""
    clean_cccd = str(cccd).strip()
    if not clean_cccd:
        return None
    connection = _approvals_connection()
    try:
        row = connection.execute(
            "SELECT * FROM profile_edit_requests WHERE cccd = ? AND workbook = ? AND sheet = ? ORDER BY id DESC LIMIT 1",
            (clean_cccd, *_request_source()),
        ).fetchone()
        return _row_to_request(row) if row else None
    finally:
        connection.close()


def get_request_by_id(request_id: int) -> Optional[Dict[str, Any]]:
    """Retrieve a single request by its ID."""
    connection = _approvals_connection()
    try:
        row = connection.execute(
            "SELECT * FROM profile_edit_requests WHERE id = ?",
            (request_id,),
        ).fetchone()
        return _row_to_request(row) if row else None
    finally:
        connection.close()


def count_pending_requests() -> int:
    """Return the total number of pending requests in the review queue."""
    connection = _approvals_connection()
    try:
        row = connection.execute(
            "SELECT COUNT(*) FROM profile_edit_requests WHERE status = 'pending'"
        ).fetchone()
        return row[0] if row else 0
    finally:
        connection.close()


def list_requests(
    status: Optional[str] = None,
    query: Optional[str] = None,
    limit: int = 50,
    offset: int = 0,
) -> Tuple[List[Dict[str, Any]], int]:
    """Retrieve requests with filtering by status, search query, and pagination."""
    conditions: List[str] = []
    params: List[Any] = []

    clean_status = (status or "").strip().lower()
    if clean_status and clean_status != "all":
        conditions.append("status = ?")
        params.append(clean_status)

    clean_query = (query or "").strip()
    if clean_query:
        like_term = f"%{clean_query}%"
        conditions.append("(cccd LIKE ? OR record_name LIKE ? OR stt LIKE ?)")
        params.extend([like_term, like_term, like_term])

    where_clause = f" WHERE {' AND '.join(conditions)}" if conditions else ""

    connection = _approvals_connection()
    try:
        count_sql = f"SELECT COUNT(*) FROM profile_edit_requests{where_clause}"
        total = connection.execute(count_sql, params).fetchone()[0]

        select_sql = f"""
            SELECT * FROM profile_edit_requests{where_clause}
            ORDER BY CASE WHEN status = 'pending' THEN 0 ELSE 1 END, id DESC
            LIMIT ? OFFSET ?
        """
        rows = connection.execute(select_sql, [*params, limit, offset]).fetchall()
        return [_row_to_request(r) for r in rows], total
    finally:
        connection.close()


def create_profile_edit_request(
    cccd: str,
    stt: str,
    record_name: str,
    submitted_values: Dict[str, str],
    original_values: Dict[str, str],
    changes: List[Dict[str, str]],
) -> int:
    """Store a pending request with snapshots keyed by header, bound to its source.
    
    Returns the created request ID.
    """
    clean_cccd = cccd.strip()
    clean_stt = str(stt).strip()
    clean_name = record_name.strip()
    workbook, sheet = _request_source()
    if not workbook or not sheet:
        raise ValueError("Hãy mở tệp Excel trước khi gửi yêu cầu cập nhật.")
    now = datetime.now().isoformat()

    connection = _approvals_connection()
    try:
        cursor = connection.execute(
            """
            INSERT INTO profile_edit_requests (
                cccd, workbook, sheet, stt, record_name, status, submitted_at, submitted_data, original_data, changes
            ) VALUES (?, ?, ?, ?, ?, 'pending', ?, ?, ?, ?)
            """,
            (
                clean_cccd,
                workbook,
                sheet,
                clean_stt,
                clean_name,
                now,
                json.dumps(submitted_values, ensure_ascii=False),
                json.dumps(original_values, ensure_ascii=False),
                json.dumps(changes, ensure_ascii=False),
            ),
        )
        connection.commit()
        req_id = cursor.lastrowid
    finally:
        connection.close()

    # Immutable audit log entry for submission lifecycle event
    log_action(
        username=f"Khách tra cứu (CCCD {clean_cccd})",
        action="request_update",
        stt=clean_stt,
        record_name=clean_name,
        details={
            "request_id": req_id,
            "cccd": clean_cccd,
            "workbook": workbook,
            "sheet": sheet,
            "lifecycle": "submitted",
            "submitted_at": now,
            "changes": changes,
        },
    )

    return req_id


def approve_request(
    request_id: int,
    reviewer: str,
    admin_note: str = "",
) -> Dict[str, Any]:
    """Approve a pending request, apply diff to the Excel workbook, unlock user, and log audit."""
    req = get_request_by_id(request_id)
    if not req:
        raise ValueError("Không tìm thấy yêu cầu xét duyệt.")
    if req["status"] != "pending":
        raise ValueError("Yêu cầu này đã được xử lý trước đó.")

    source = (req["workbook"], req["sheet"])
    if not all(source):
        raise ValueError("Yêu cầu cũ chưa ghi nhận tệp Excel gốc. Hãy từ chối yêu cầu và đề nghị gửi lại.")
    if _request_source() != source:
        raise ValueError("Hãy mở đúng tệp Excel và trang tính gốc của yêu cầu trước khi phê duyệt.")
    refresh_state()
    if state["error"]:
        raise ValueError(state["error"])
    if _request_source() != source:
        raise ValueError("Trang tính gốc của yêu cầu đã thay đổi. Không thể phê duyệt.")

    cccd = req["cccd"]
    record = find_record_by_citizen_id(cccd)
    if record is None:
        raise ValueError("Không tìm thấy hồ sơ tương ứng trong tệp Excel hiện tại.")

    # Apply exactly the displayed diff, resolving headers against the refreshed workbook.
    # Column numbers in an old form are not stable when Excel columns are rearranged.
    columns_by_header = {d["header"]: d["column"] for d in field_definitions()}
    updates = {}
    for change in req["changes"]:
        header = change["field"]
        if header in IDENTITY_CRITICAL_HEADERS:
            raise ValueError(f"Không được thay đổi trường định danh: {header}.")
        if header not in columns_by_header:
            raise ValueError(f"Tệp Excel không còn cột {header}. Không thể phê duyệt yêu cầu này.")
        if comparable_field_value(display_value(record["data"].get(header))) != comparable_field_value(display_value(change["from"])):
            raise ValueError(f"Thông tin {header} đã thay đổi sau khi gửi yêu cầu. Hãy từ chối và đề nghị gửi lại.")
        updates[columns_by_header[header]] = coerce_cell_value(header, display_value(change["to"]))
    if not updates:
        raise ValueError("Yêu cầu không có thay đổi nào để phê duyệt.")

    # Save to Excel workbook atomically
    row, stt, backup = save_person(
        state["workbook"],
        state["sheet"],
        updates,
        signature_token(state["signature"]),
        original_stt=record["stt"],
    )
    refresh_state(force=True)

    # Update request row to approved
    now = datetime.now().isoformat()
    clean_note = admin_note.strip()
    connection = _approvals_connection()
    try:
        connection.execute(
            """
            UPDATE profile_edit_requests
            SET status = 'approved', reviewer = ?, reviewed_at = ?, reviewer_note = ?
            WHERE id = ?
            """,
            (reviewer, now, clean_note, request_id),
        )
        connection.commit()
    finally:
        connection.close()

    # Immutable audit log entry for approval lifecycle event
    log_action(
        username=reviewer,
        action="approve_update",
        stt=str(stt),
        record_name=req["record_name"],
        details={
            "request_id": request_id,
            "cccd": cccd,
            "reviewer": reviewer,
            "reviewed_at": now,
            "admin_note": clean_note,
            "lifecycle": "approved",
            "changes": req["changes"],
        },
    )

    return get_request_by_id(request_id)


def reject_request(
    request_id: int,
    reviewer: str,
    rejection_reason: str = "",
) -> Dict[str, Any]:
    """Reject a pending request, unlock user without modifying workbook, and log audit."""
    req = get_request_by_id(request_id)
    if not req:
        raise ValueError("Không tìm thấy yêu cầu xét duyệt.")
    if req["status"] != "pending":
        raise ValueError("Yêu cầu này đã được xử lý trước đó.")

    now = datetime.now().isoformat()
    clean_reason = rejection_reason.strip()

    connection = _approvals_connection()
    try:
        connection.execute(
            """
            UPDATE profile_edit_requests
            SET status = 'rejected', reviewer = ?, reviewed_at = ?, reviewer_note = ?
            WHERE id = ?
            """,
            (reviewer, now, clean_reason, request_id),
        )
        connection.commit()
    finally:
        connection.close()

    # Immutable audit log entry for rejection lifecycle event
    log_action(
        username=reviewer,
        action="reject_update",
        stt=req["stt"],
        record_name=req["record_name"],
        details={
            "request_id": request_id,
            "cccd": req["cccd"],
            "reviewer": reviewer,
            "reviewed_at": now,
            "rejection_reason": clean_reason,
            "admin_note": clean_reason,
            "lifecycle": "rejected",
            "changes": req["changes"],
        },
    )

    return get_request_by_id(request_id)

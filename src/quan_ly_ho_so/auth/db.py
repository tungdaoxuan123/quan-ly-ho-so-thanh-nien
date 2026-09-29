"""SQLite user store, schema migrations, and user CRUD operations."""

from __future__ import annotations

import logging
import sqlite3
from datetime import datetime
from typing import Any, Dict, List, Optional
from werkzeug.security import generate_password_hash

from quan_ly_ho_so.auth.config import (
    DEFAULT_ADMIN_PASSWORD,
    DEFAULT_ADMIN_USERNAME,
    DEFAULT_USER_PASSWORD,
    USERNAME_REGEX,
    auth_database_path,
)


def auth_database_connection() -> sqlite3.Connection:
    """Open or create the SQLite database for user authentication."""
    path = auth_database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL COLLATE NOCASE,
            password_hash TEXT NOT NULL,
            role TEXT NOT NULL,
            can_create INTEGER NOT NULL DEFAULT 0,
            can_read INTEGER NOT NULL DEFAULT 1,
            can_update INTEGER NOT NULL DEFAULT 0,
            can_delete INTEGER NOT NULL DEFAULT 0,
            must_change_password INTEGER NOT NULL DEFAULT 0,
            created_at TEXT NOT NULL,
            updated_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS idx_users_username ON users(username);
        """
    )
    # Automatically migrate existing SQLite DB if must_change_password column is missing
    columns = [row[1] for row in connection.execute("PRAGMA table_info(users)").fetchall()]
    if "must_change_password" not in columns:
        connection.execute("ALTER TABLE users ADD COLUMN must_change_password INTEGER NOT NULL DEFAULT 0")
        connection.commit()

    connection.commit()
    _ensure_single_admin(connection)
    return connection


def _ensure_single_admin(connection: sqlite3.Connection) -> None:
    row = connection.execute("SELECT id FROM users WHERE role = 'admin' LIMIT 1").fetchone()
    if row is None:
        now = datetime.now().isoformat()
        admin_hash = generate_password_hash(DEFAULT_ADMIN_PASSWORD)
        connection.execute(
            """
            INSERT INTO users (
                username, password_hash, role, can_create, can_read, can_update, can_delete, must_change_password, created_at, updated_at
            ) VALUES (?, ?, 'admin', 1, 1, 1, 1, 0, ?, ?)
            """,
            (DEFAULT_ADMIN_USERNAME, admin_hash, now, now),
        )
        connection.commit()
        logging.info("Initialized default admin account '%s'", DEFAULT_ADMIN_USERNAME)


def get_user_by_id(user_id: int) -> Optional[Dict[str, Any]]:
    conn = auth_database_connection()
    try:
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def get_user_by_username(username: str) -> Optional[Dict[str, Any]]:
    conn = auth_database_connection()
    try:
        row = conn.execute("SELECT * FROM users WHERE username = ?", (username.strip(),)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def list_users() -> List[Dict[str, Any]]:
    conn = auth_database_connection()
    try:
        cursor = conn.execute("SELECT * FROM users ORDER BY CASE WHEN role = 'admin' THEN 0 ELSE 1 END, id ASC")
        return [dict(row) for row in cursor.fetchall()]
    finally:
        conn.close()


def create_user(
    username: str,
    password: Optional[str] = None,
    can_create: bool = False,
    can_read: bool = True,
    can_update: bool = False,
    can_delete: bool = False,
    must_change_password: bool = True,
) -> Dict[str, Any]:
    username = (username or "").strip()
    if not USERNAME_REGEX.match(username):
        raise ValueError("Tên đăng nhập phải từ 3 đến 30 ký tự, chỉ chứa chữ cái, chữ số, dấu gạch dưới hoặc gạch nối.")
    password = (password or "").strip() if password is not None else ""
    if not password:
        password = DEFAULT_USER_PASSWORD
    if len(password) < 4:
        raise ValueError("Mật khẩu phải có ít nhất 4 ký tự.")

    conn = auth_database_connection()
    try:
        existing = conn.execute("SELECT id FROM users WHERE username = ?", (username,)).fetchone()
        if existing:
            raise ValueError(f"Tên đăng nhập '{username}' đã tồn tại trong hệ thống.")

        now = datetime.now().isoformat()
        password_hash = generate_password_hash(password)
        cursor = conn.execute(
            """
            INSERT INTO users (
                username, password_hash, role, can_create, can_read, can_update, can_delete, must_change_password, created_at, updated_at
            ) VALUES (?, ?, 'user', ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                password_hash,
                1 if can_create else 0,
                1 if can_read else 0,
                1 if can_update else 0,
                1 if can_delete else 0,
                1 if must_change_password else 0,
                now,
                now,
            ),
        )
        conn.commit()
        user_id = cursor.lastrowid
        row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return dict(row)
    finally:
        conn.close()


def update_user(
    user_id: int,
    password: Optional[str] = None,
    can_create: Optional[bool] = None,
    can_read: Optional[bool] = None,
    can_update: Optional[bool] = None,
    can_delete: Optional[bool] = None,
    must_change_password: Optional[bool] = None,
) -> Dict[str, Any]:
    conn = auth_database_connection()
    try:
        user = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not user:
            raise ValueError("Không tìm thấy người dùng.")

        now = datetime.now().isoformat()
        is_admin = (user["role"] == "admin")

        updates = ["updated_at = ?"]
        params: List[Any] = [now]

        if password:
            if len(password) < 4:
                raise ValueError("Mật khẩu mới phải có ít nhất 4 ký tự.")
            updates.append("password_hash = ?")
            params.append(generate_password_hash(password))

        if must_change_password is not None:
            updates.append("must_change_password = ?")
            params.append(1 if must_change_password else 0)

        if not is_admin:
            if can_create is not None:
                updates.append("can_create = ?")
                params.append(1 if can_create else 0)
            if can_read is not None:
                updates.append("can_read = ?")
                params.append(1 if can_read else 0)
            if can_update is not None:
                updates.append("can_update = ?")
                params.append(1 if can_update else 0)
            if can_delete is not None:
                updates.append("can_delete = ?")
                params.append(1 if can_delete else 0)

        params.append(user_id)
        conn.execute(f"UPDATE users SET {', '.join(updates)} WHERE id = ?", params)
        conn.commit()

        updated_row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return dict(updated_row)
    finally:
        conn.close()


def reset_user_password(user_id: int) -> Dict[str, Any]:
    """Reset a user's password to DEFAULT_USER_PASSWORD ('123456') and require change on next login."""
    conn = auth_database_connection()
    try:
        user = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not user:
            raise ValueError("Không tìm thấy người dùng.")
        if user["role"] == "admin":
            raise ValueError("Hệ thống chỉ có 1 Quản trị viên duy nhất, không thể đặt lại mật khẩu theo cách này.")

        now = datetime.now().isoformat()
        password_hash = generate_password_hash(DEFAULT_USER_PASSWORD)
        conn.execute(
            "UPDATE users SET password_hash = ?, must_change_password = 1, updated_at = ? WHERE id = ?",
            (password_hash, now, user_id),
        )
        conn.commit()
        updated_row = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        return dict(updated_row)
    finally:
        conn.close()


def delete_user(user_id: int) -> None:
    conn = auth_database_connection()
    try:
        user = conn.execute("SELECT * FROM users WHERE id = ?", (user_id,)).fetchone()
        if not user:
            raise ValueError("Không tìm thấy người dùng cần xóa.")
        if user["role"] == "admin":
            raise ValueError("Hệ thống chỉ có 1 Quản trị viên duy nhất, không thể xóa tài khoản Quản trị viên.")

        conn.execute("DELETE FROM users WHERE id = ?", (user_id,))
        conn.commit()
    finally:
        conn.close()

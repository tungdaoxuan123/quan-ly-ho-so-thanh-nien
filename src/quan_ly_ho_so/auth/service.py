"""Authentication business logic, session management, and current user context."""

from __future__ import annotations

from typing import Any, Dict, Optional
from flask import session
from werkzeug.security import check_password_hash

from quan_ly_ho_so.auth.db import get_user_by_username


def authenticate(username: str, password: str) -> Optional[Dict[str, Any]]:
    """Validate username and password credentials against the SQLite database."""
    user = get_user_by_username(username)
    if not user:
        return None
    if not check_password_hash(user["password_hash"], password):
        return None
    return user


def login_user(user: Dict[str, Any]) -> None:
    """Store the authenticated user and their permissions directly into the Flask session.

    This ensures subsequent requests read permissions straight from the session
    cookie without needing repeated DB lookups.
    """
    session.clear()
    session.permanent = True
    is_admin = user["role"] == "admin"
    session["user_id"] = user["id"]
    session["username"] = user["username"]
    session["role"] = user["role"]
    session["must_change_password"] = bool(user.get("must_change_password", 0))
    session["permissions"] = {
        "create": True if is_admin else bool(user["can_create"]),
        "read": True if is_admin else bool(user["can_read"]),
        "update": True if is_admin else bool(user["can_update"]),
        "delete": True if is_admin else bool(user["can_delete"]),
    }


def logout_user() -> None:
    """Clear user authentication data from the current session."""
    session.pop("user_id", None)
    session.pop("username", None)
    session.pop("role", None)
    session.pop("permissions", None)
    session.pop("must_change_password", None)


def current_user() -> Optional[Dict[str, Any]]:
    """Return user info and cached permissions directly from the session.

    Does NOT query SQLite database on each request to check permissions.
    """
    user_id = session.get("user_id")
    if not user_id:
        return None
    role = session.get("role", "user")
    is_admin = role == "admin"
    permissions = session.get("permissions", {})
    return {
        "id": user_id,
        "username": session.get("username", ""),
        "role": role,
        "is_admin": is_admin,
        "can_create": is_admin or bool(permissions.get("create")),
        "can_read": is_admin or bool(permissions.get("read")),
        "can_update": is_admin or bool(permissions.get("update")),
        "can_delete": is_admin or bool(permissions.get("delete")),
        "must_change_password": bool(session.get("must_change_password", False)),
    }

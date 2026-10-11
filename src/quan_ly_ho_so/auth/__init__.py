"""User authentication, authorization, SQLite user store, and session helpers."""

from __future__ import annotations

from quan_ly_ho_so.auth.config import (
    DEFAULT_ADMIN_PASSWORD,
    DEFAULT_ADMIN_USERNAME,
    DEFAULT_SESSION_LIFETIME_DAYS,
    DEFAULT_USER_PASSWORD,
    USERNAME_REGEX,
    auth_database_path,
    get_persistent_secret_key,
    load_env,
    set_auth_database_path,
)
from quan_ly_ho_so.auth.db import (
    auth_database_connection,
    create_user,
    delete_user,
    get_user_by_id,
    get_user_by_username,
    list_users,
    reset_user_password,
    update_user,
)
from quan_ly_ho_so.auth.decorators import (
    admin_required,
    auth_required,
    login_required,
    permission_required,
)
from quan_ly_ho_so.auth.service import (
    authenticate,
    current_user,
    login_user,
    logout_user,
)

__all__ = [
    "DEFAULT_ADMIN_PASSWORD",
    "DEFAULT_ADMIN_USERNAME",
    "DEFAULT_SESSION_LIFETIME_DAYS",
    "DEFAULT_USER_PASSWORD",
    "USERNAME_REGEX",
    "admin_required",
    "auth_database_connection",
    "auth_database_path",
    "auth_required",
    "authenticate",
    "create_user",
    "current_user",
    "delete_user",
    "get_persistent_secret_key",
    "get_user_by_id",
    "get_user_by_username",
    "list_users",
    "load_env",
    "login_required",
    "login_user",
    "logout_user",
    "permission_required",
    "reset_user_password",
    "set_auth_database_path",
    "update_user",
]

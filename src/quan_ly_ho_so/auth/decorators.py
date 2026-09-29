"""Route and endpoint decorators for authentication, roles, and CRUD permissions."""

from __future__ import annotations

import functools
from typing import Any, Callable, Iterable, Optional, Union
from flask import flash, jsonify, redirect, request, session, url_for

from quan_ly_ho_so.auth.service import current_user

PERMISSION_LABELS = {
    "create": "thêm mới",
    "read": "xem",
    "update": "chỉnh sửa",
    "delete": "xóa",
}


def _is_api_request() -> bool:
    """Detect whether the incoming request expects a JSON response."""
    if request.path.startswith("/api/"):
        return True
    accept = request.headers.get("Accept", "")
    return "application/json" in accept and not request.path.startswith("/downloads")


def auth_required(
    func: Optional[Callable[..., Any]] = None,
    *,
    permission: Optional[Union[str, Iterable[str]]] = None,
    permissions: Optional[Iterable[str]] = None,
    role: Optional[str] = None,
    require_all: bool = True,
):
    """Unified decorator for endpoint authentication and authorization.

    Enforces login, forced password change restrictions, role verification,
    and granular CRUD data permissions.

    Usage examples:
        @app.get("/dashboard")
        @auth_required
        def dashboard(): ...

        @app.get("/admin")
        @auth_required(role="admin")
        def admin(): ...

        @app.get("/person/new")
        @auth_required(permission="create")
        def new_person(): ...

        @app.get("/records")
        @auth_required(permissions=["read", "update"], require_all=False)
        def view_or_edit(): ...
    """
    def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(fn)
        def wrapper(*args, **kwargs):
            is_api = _is_api_request()

            # 1. Verify authenticated session
            user_id = session.get("user_id")
            if not user_id:
                if is_api:
                    return jsonify({"error": "Unauthorized", "message": "Vui lòng đăng nhập để tiếp tục."}), 401
                flash("Vui lòng đăng nhập để tiếp tục.", "error")
                next_url = request.full_path if request.query_string else request.path
                try:
                    login_url = url_for("login", next=next_url)
                except Exception:
                    login_url = "/login"
                return redirect(login_url)

            user = current_user()
            if not user:
                if is_api:
                    return jsonify({"error": "Unauthorized", "message": "Phiên làm việc không hợp lệ."}), 401
                flash("Vui lòng đăng nhập để tiếp tục.", "error")
                try:
                    login_url = url_for("login")
                except Exception:
                    login_url = "/login"
                return redirect(login_url)

            # 2. Enforce password change if required
            force_endpoints = {"force_change_password", "force_change_password_post", "logout"}
            if user.get("must_change_password") and request.endpoint not in force_endpoints:
                if is_api:
                    return jsonify({"error": "Forbidden", "message": "Yêu cầu đổi mật khẩu mới trước khi tiếp tục."}), 403
                try:
                    force_url = url_for("force_change_password")
                except Exception:
                    force_url = "/force-change-password"
                return redirect(force_url)

            # 3. Enforce role requirement (e.g. admin)
            if role is not None:
                if user.get("role") != role:
                    if is_api:
                        return jsonify({"error": "Forbidden", "message": f"Chỉ {role} mới có quyền truy cập."}), 403
                    if role == "admin":
                        flash("Chỉ Quản trị viên mới có quyền truy cập trang quản trị.", "error")
                    else:
                        flash(f"Chỉ người dùng có vai trò '{role}' mới có quyền truy cập.", "error")
                    return redirect(url_for("index"))

            # 4. Enforce permission requirements (create, read, update, delete)
            req_perms: list[str] = []
            if permission is not None:
                if isinstance(permission, str):
                    req_perms.append(permission)
                else:
                    req_perms.extend(permission)
            if permissions is not None:
                req_perms.extend(permissions)

            if req_perms:
                if require_all:
                    denied = [p for p in req_perms if not user.get(f"can_{p}")]
                    if denied:
                        if is_api:
                            return jsonify({"error": "Forbidden", "message": f"Thiếu quyền: {', '.join(denied)}"}), 403
                        if len(denied) == 1:
                            action_text = PERMISSION_LABELS.get(denied[0], denied[0])
                            flash(f"Bạn không có quyền {action_text} dữ liệu hồ sơ.", "error")
                        else:
                            actions = [PERMISSION_LABELS.get(p, p) for p in denied]
                            flash(f"Bạn không có quyền {' và '.join(actions)} dữ liệu hồ sơ.", "error")
                        return redirect(url_for("index"))
                else:
                    allowed = [p for p in req_perms if user.get(f"can_{p}")]
                    if not allowed:
                        if is_api:
                            return jsonify({"error": "Forbidden", "message": f"Yêu cầu ít nhất một trong các quyền: {', '.join(req_perms)}"}), 403
                        actions = [PERMISSION_LABELS.get(p, p) for p in req_perms]
                        flash(f"Bạn cần quyền {' hoặc '.join(actions)} dữ liệu hồ sơ.", "error")
                        return redirect(url_for("index"))

            return fn(*args, **kwargs)

        return wrapper

    if func is not None:
        return decorator(func)
    return decorator


def login_required(func: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator requiring a logged-in user session."""
    return auth_required(func)


def admin_required(func: Callable[..., Any]) -> Callable[..., Any]:
    """Decorator requiring admin role."""
    return auth_required(role="admin")(func)


def permission_required(permission_name: str) -> Callable[..., Any]:
    """Decorator requiring a specific CRUD permission (create, read, update, delete)."""
    return auth_required(permission=permission_name)

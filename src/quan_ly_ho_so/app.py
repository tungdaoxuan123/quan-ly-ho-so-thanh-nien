#!/usr/bin/env python3
"""Local browser UI for Excel-backed youth records."""

from __future__ import annotations

import hashlib
import io
import logging
import getpass
import os
import secrets
import socket
import subprocess
import sys
import threading
import webbrowser
import zipfile
from datetime import datetime
from pathlib import Path

from flask import (
    Flask,
    Response,
    abort,
    flash,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    send_from_directory,
    url_for,
)
from datetime import timedelta

from werkzeug.security import check_password_hash
from werkzeug.utils import secure_filename

from quan_ly_ho_so.auth import (
    DEFAULT_ADMIN_PASSWORD,
    DEFAULT_ADMIN_USERNAME,
    DEFAULT_SESSION_LIFETIME_DAYS,
    DEFAULT_USER_PASSWORD,
    admin_required,
    auth_database_connection,
    auth_required,
    authenticate,
    create_user,
    current_user,
    delete_user,
    get_persistent_secret_key,
    get_user_by_id,
    get_user_by_username,
    list_users,
    login_required,
    login_user,
    logout_user,
    permission_required,
    reset_user_password,
    update_user,
)
from quan_ly_ho_so.config import (
    DEFAULT_OUTPUT_NAME,
    DIEN_TEMPLATE_DIR,
    FILTER_DB_COLUMNS,
    FORM_GROUPS,
    HTTPS_DIR,
    LIST_TEMPLATE,
    MAX_EXPORT_RECORDS,
    MAX_RECORDS_PER_PAGE,
    SUPPORTED_SUFFIXES,
    TEMPLATE,
    UNSET_DIEN_FILTER,
    QUARTER_CHOICES,
    UPLOAD_DIR,
)
from quan_ly_ho_so import documents
from quan_ly_ho_so.attachments import AttachmentError, document_thumbnail, prepare_document
from quan_ly_ho_so.audit import compute_changes, get_audit_logs, log_action
from quan_ly_ho_so.errors import StaleWorkbookError
from quan_ly_ho_so.forms.enums import NvqsStatus
from quan_ly_ho_so.forms.fields import (
    CITIZEN_ID_ERROR,
    coerce_cell_value,
    field_definitions,
    form_values,
    valid_citizen_id,
    validate_form_values,
)
from quan_ly_ho_so.security import client_ip, csrf_token, is_local_network, online_mode, valid_csrf
from quan_ly_ho_so.throttle import Throttle
from quan_ly_ho_so.tls import lan_certificate, lan_ssl_context
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import display_value, fold
from quan_ly_ho_so.web.templates import (
    ADMIN_PAGE,
    AUDIT_LOG_PAGE,
    FORCE_CHANGE_PASSWORD_PAGE,
    GUEST_PAGE,
    LOGIN_PAGE,
    PAGE,
    PROFILE_PAGE,
    page_links,
    render_form,
    render_guest_form,
    render_preview,
)
from quan_ly_ho_so.word.export import build_document, combine_booklet, unique_output_path
from quan_ly_ho_so.workbook.cache import (
    active_people,
    records_by_citizen_id,
    database_record_count,
    delete_record_file,
    find_record,
    move_record_owner,
    record_file_image,
    find_record_by_citizen_id,
    query_filter_options,
    query_records,
    set_record_attribute,
    set_records_deleted,
    set_record_type,
    signature_token,
)
from quan_ly_ho_so.workbook.ca_import import plan_import, read_ca_files
from quan_ly_ho_so.workbook.dien_export import DIEN_TEMPLATES, build_dien_workbook
from quan_ly_ho_so.workbook.list_export import build_list_workbook
from quan_ly_ho_so.workbook.manager import (
    load_saved_workbook,
    native_workbook_dialog,
    refresh_state,
    set_workbook,
)
from quan_ly_ho_so.workbook.writer import append_people, save_person

app = Flask(__name__)
app.secret_key = get_persistent_secret_key()
app.permanent_session_lifetime = timedelta(days=DEFAULT_SESSION_LIFETIME_DAYS)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
if online_mode():
    # Published through an https tunnel: keep the session cookie off plain http.
    app.config.update(SESSION_COOKIE_SECURE=True, SESSION_COOKIE_SAMESITE="Lax")

# Guards against guessing passwords and mass-scraping or spamming the public guest pages.
login_failures = Throttle(limit=5, window_seconds=15 * 60)
guest_lookups = Throttle(limit=30, window_seconds=10 * 60)
guest_saves = Throttle(limit=5, window_seconds=60 * 60)
TOO_MANY_ATTEMPTS = "Bạn thao tác quá nhiều lần. Vui lòng thử lại sau ít phút."


@app.before_request
def guard_network():
    """Keep the app on the local network: refuse anything that is not this computer or a private address.
    Online mode publishes the app on purpose, so login, throttling and https cookies protect it instead."""
    if not online_mode() and not is_local_network(request.remote_addr):
        abort(403)


@app.context_processor
def inject_user():
    user = current_user()
    pending_count = 0
    if user and (user.get("is_admin") or user.get("can_update")):
        try:
            from quan_ly_ho_so.approvals import count_pending_requests
            pending_count = count_pending_requests()
        except Exception:
            pending_count = 0
    return {"current_user": user, "pending_request_count": pending_count}


@app.before_request
def require_authentication():
    exempt_endpoints = {"login", "login_post", "logout", "static", "guest_lookup", "guest_lookup_post", "guest_start_create", "guest_save", "guest_update_request", "https_ca_certificate"}
    if request.endpoint in exempt_endpoints or request.endpoint is None:
        return None
    user = current_user()
    if not user:
        # The bare address (https://<máy>:3000) is what guests type: send them to the CCCD lookup, not the login form.
        if request.endpoint == "index" and request.method == "GET" and not request.query_string:
            return redirect(url_for("guest_lookup"))
        if request.method == "GET":
            next_url = request.full_path if request.query_string else request.path
            return redirect(url_for("login", next=next_url))
        return redirect(url_for("login"))

    force_endpoints = {"force_change_password", "force_change_password_post", "logout"}
    if user.get("must_change_password") and request.endpoint not in force_endpoints:
        return redirect(url_for("force_change_password"))


@app.get("/login")
def login():
    if current_user():
        return redirect(url_for("index"))
    next_url = request.args.get("next", "")
    return render_template_string(
        LOGIN_PAGE,
        csrf=csrf_token(),
        next_url=next_url,
        username="",
    )


@app.post("/login")
def login_post():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử đăng nhập lại.", "error")
        return redirect(url_for("login"))
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    next_url = request.form.get("next", "")
    attempt_key = (client_ip(), username.lower())
    if login_failures.blocked(attempt_key):
        flash(TOO_MANY_ATTEMPTS, "error")
        return render_template_string(LOGIN_PAGE, csrf=csrf_token(), next_url=next_url, username=username), 429
    user = authenticate(username, password)
    if user is None:
        login_failures.record(attempt_key)
        flash("Tên đăng nhập hoặc mật khẩu không chính xác.", "error")
        return render_template_string(
            LOGIN_PAGE,
            csrf=csrf_token(),
            next_url=next_url,
            username=username,
        )
    login_failures.clear(attempt_key)
    login_user(user)
    if user.get("must_change_password"):
        flash("Đây là lần đầu đăng nhập hoặc mật khẩu vừa được đặt lại. Vui lòng tự cài đặt mật khẩu mới.", "error")
        return redirect(url_for("force_change_password"))
    flash(f"Đăng nhập thành công. Xin chào {user['username']}!", "success")
    if next_url and next_url.startswith("/") and not next_url.startswith("//"):
        return redirect(next_url)
    return redirect(url_for("index"))


@app.route("/logout", methods=["GET", "POST"])
def logout():
    logout_user()
    flash("Đã đăng xuất khỏi hệ thống.", "success")
    return redirect(url_for("login"))


@app.get("/force-change-password")
@auth_required
def force_change_password():
    user = current_user()
    if not user or not user.get("must_change_password"):
        return redirect(url_for("index"))
    return render_template_string(
        FORCE_CHANGE_PASSWORD_PAGE,
        csrf=csrf_token(),
        username=user["username"],
    )


@app.post("/force-change-password")
@auth_required
def force_change_password_post():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("force_change_password"))
    user = current_user()
    if not user:
        return redirect(url_for("login"))
    new_pw = request.form.get("new_password", "").strip()
    confirm_pw = request.form.get("confirm_password", "").strip()
    if len(new_pw) < 4:
        flash("Mật khẩu mới phải có ít nhất 4 ký tự.", "error")
        return render_template_string(FORCE_CHANGE_PASSWORD_PAGE, csrf=csrf_token(), username=user["username"])
    if new_pw == DEFAULT_USER_PASSWORD:
        flash("Mật khẩu mới không được trùng với mật khẩu mặc định (123456). Vui lòng chọn mật khẩu mới khác an toàn hơn.", "error")
        return render_template_string(FORCE_CHANGE_PASSWORD_PAGE, csrf=csrf_token(), username=user["username"])
    if new_pw != confirm_pw:
        flash("Mật khẩu mới và xác nhận mật khẩu không khớp nhau.", "error")
        return render_template_string(FORCE_CHANGE_PASSWORD_PAGE, csrf=csrf_token(), username=user["username"])
    try:
        update_user(user["id"], password=new_pw, must_change_password=False)
        session["must_change_password"] = False
        flash("Cài đặt mật khẩu mới thành công! Bạn có thể bắt đầu sử dụng hệ thống.", "success")
        return redirect(url_for("index"))
    except Exception as error:
        flash(f"Không thể cập nhật mật khẩu: {error}", "error")
        return render_template_string(FORCE_CHANGE_PASSWORD_PAGE, csrf=csrf_token(), username=user["username"])


@app.get("/profile")
@auth_required
def user_profile():
    user_info = current_user()
    user_db = get_user_by_id(user_info["id"])
    return render_template_string(
        PROFILE_PAGE,
        user=user_db or user_info,
        current_user=user_info,
        csrf=csrf_token(),
    )


@app.post("/profile/change-password")
@auth_required
def profile_change_password():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("user_profile"))
    user_info = current_user()
    user_db = get_user_by_id(user_info["id"])
    current_pw = request.form.get("current_password", "")
    new_pw = request.form.get("new_password", "").strip()
    confirm_pw = request.form.get("confirm_password", "").strip()
    if not user_db or not check_password_hash(user_db["password_hash"], current_pw):
        flash("Mật khẩu hiện tại không chính xác.", "error")
        return redirect(url_for("user_profile"))
    if len(new_pw) < 4:
        flash("Mật khẩu mới phải có ít nhất 4 ký tự.", "error")
        return redirect(url_for("user_profile"))
    if new_pw == DEFAULT_USER_PASSWORD:
        flash("Mật khẩu mới không được trùng với mật khẩu mặc định (123456). Vui lòng chọn mật khẩu mới khác an toàn hơn.", "error")
        return redirect(url_for("user_profile"))
    if new_pw != confirm_pw:
        flash("Mật khẩu mới và xác nhận mật khẩu không khớp nhau.", "error")
        return redirect(url_for("user_profile"))
    try:
        update_user(user_info["id"], password=new_pw, must_change_password=False)
        session["must_change_password"] = False
        flash("Đổi mật khẩu thành công!", "success")
    except Exception as error:
        flash(f"Không thể đổi mật khẩu: {error}", "error")
    return redirect(url_for("user_profile"))


@app.get("/admin")
@auth_required(role="admin")
def admin_dashboard():
    users = list_users()
    admin_user = next((u for u in users if u["role"] == "admin"), None)
    return render_template_string(
        ADMIN_PAGE,
        users=users,
        admin_user=admin_user,
        csrf=csrf_token(),
        current_user=current_user(),
    )


@app.get("/admin/requests")
@app.get("/admin/approvals")
@auth_required(permissions=["update"], require_all=False)
def admin_approvals_queue():
    from quan_ly_ho_so.approvals import count_pending_requests, list_requests
    from quan_ly_ho_so.web.approvals_templates import APPROVALS_PAGE
    status_filter = request.args.get("status", "pending").strip().lower()
    query = request.args.get("q", "").strip()
    try:
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        page = 1
    limit = 20
    offset = (page - 1) * limit
    reqs, total_filtered = list_requests(status=status_filter, query=query, limit=limit, offset=offset)
    pending_count = count_pending_requests()
    _, total_all = list_requests(status="all", limit=1, offset=0)
    total_pages = max(1, (total_filtered + limit - 1) // limit)
    return render_template_string(
        APPROVALS_PAGE,
        requests=reqs,
        pending_count=pending_count,
        total_requests=total_all,
        total_filtered=total_filtered,
        current_status=status_filter,
        current_query=query,
        current_page=page,
        total_pages=total_pages,
        csrf=csrf_token(),
        current_user=current_user(),
    )


@app.post("/admin/requests/<int:request_id>/review")
@auth_required(permissions=["update"], require_all=False)
def admin_review_request(request_id):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("admin_approvals_queue"))
    action = request.form.get("action", "").strip().lower()
    note = request.form.get("note", "").strip()
    return_status = request.form.get("return_status", "pending")
    return_q = request.form.get("return_q", "")
    user = current_user()
    reviewer = user.get("username", "admin")

    from quan_ly_ho_so.approvals import approve_request, reject_request
    try:
        if action == "approve":
            approve_request(request_id, reviewer=reviewer, admin_note=note)
            flash(f"Đã PHÊ DUYỆT yêu cầu #{request_id} và đồng bộ thông tin hồ sơ thành công.", "success")
        elif action == "reject":
            reject_request(request_id, reviewer=reviewer, rejection_reason=note)
            flash(f"Đã TỪ CHỐI yêu cầu #{request_id}. Hồ sơ đã được tự động mở khóa.", "info")
        else:
            flash("Hành động xét duyệt không hợp lệ.", "error")
    except Exception as error:
        logging.exception("Error reviewing request %s: %s", request_id, error)
        flash(f"Lỗi khi xét duyệt yêu cầu #{request_id}: {error}", "error")

    return redirect(url_for("admin_approvals_queue", status=return_status, q=return_q))


@app.post("/admin/users/create")
@auth_required(role="admin")
def admin_create_user():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("admin_dashboard"))
    username = request.form.get("username", "").strip()
    can_create = bool(request.form.get("can_create"))
    can_read = bool(request.form.get("can_read"))
    can_update = bool(request.form.get("can_update"))
    can_delete = bool(request.form.get("can_delete"))
    can_export = bool(request.form.get("can_export"))
    try:
        create_user(
            username,
            password=DEFAULT_USER_PASSWORD,
            can_create=can_create,
            can_read=can_read,
            can_update=can_update,
            can_delete=can_delete,
            can_export=can_export,
            must_change_password=True,
        )
        flash(f"Đã tạo người dùng '{username}' thành công (mật khẩu mặc định: '{DEFAULT_USER_PASSWORD}'). Người dùng sẽ phải tự cài mật khẩu mới khi đăng nhập lần đầu.", "success")
    except Exception as error:
        flash(f"Không thể tạo người dùng: {error}", "error")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/users/<int:user_id>/update")
@auth_required(role="admin")
def admin_update_user(user_id):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("admin_dashboard"))
    can_create = bool(request.form.get("can_create"))
    can_read = bool(request.form.get("can_read"))
    can_update = bool(request.form.get("can_update"))
    can_delete = bool(request.form.get("can_delete"))
    can_export = bool(request.form.get("can_export"))
    try:
        updated = update_user(
            user_id,
            password=None,
            can_create=can_create,
            can_read=can_read,
            can_update=can_update,
            can_delete=can_delete,
            can_export=can_export,
        )
        flash(f"Đã cập nhật quyền người dùng '{updated['username']}' thành công.", "success")
    except Exception as error:
        flash(f"Không thể cập nhật người dùng: {error}", "error")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/users/<int:user_id>/reset-password")
@auth_required(role="admin")
def admin_reset_user_password(user_id):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("admin_dashboard"))
    try:
        updated = reset_user_password(user_id)
        flash(f"Đã đặt lại mật khẩu cho người dùng '{updated['username']}' về mặc định ('{DEFAULT_USER_PASSWORD}'). Người dùng sẽ bắt buộc phải tự cài đặt mật khẩu mới khi đăng nhập.", "success")
    except Exception as error:
        flash(f"Không thể đặt lại mật khẩu: {error}", "error")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/users/<int:user_id>/delete")
@auth_required(role="admin")
def admin_delete_user(user_id):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("admin_dashboard"))
    try:
        delete_user(user_id)
        flash("Đã xóa người dùng thành công.", "success")
    except Exception as error:
        flash(f"Không thể xóa người dùng: {error}", "error")
    return redirect(url_for("admin_dashboard"))


@app.post("/admin/change-password")
@auth_required(role="admin")
def admin_change_password():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("admin_dashboard"))
    current_user_obj = current_user()
    admin_record = get_user_by_id(current_user_obj["id"])
    current_pw = request.form.get("current_password", "")
    new_pw = request.form.get("new_password", "")
    confirm_pw = request.form.get("confirm_password", "")
    if not admin_record or not check_password_hash(admin_record["password_hash"], current_pw):
        flash("Mật khẩu hiện tại không chính xác.", "error")
        return redirect(url_for("admin_dashboard"))
    if len(new_pw) < 4:
        flash("Mật khẩu mới phải có ít nhất 4 ký tự.", "error")
        return redirect(url_for("admin_dashboard"))
    if new_pw == DEFAULT_USER_PASSWORD:
        flash("Mật khẩu mới không được trùng với mật khẩu mặc định (123456). Vui lòng chọn mật khẩu mới khác an toàn hơn.", "error")
        return redirect(url_for("admin_dashboard"))
    if new_pw != confirm_pw:
        flash("Mật khẩu mới và xác nhận mật khẩu không khớp nhau.", "error")
        return redirect(url_for("admin_dashboard"))
    try:
        update_user(current_user_obj["id"], password=new_pw, must_change_password=False)
        session["must_change_password"] = False
        flash("Đã đổi mật khẩu Quản trị viên thành công.", "success")
    except Exception as error:
        flash(f"Không thể đổi mật khẩu: {error}", "error")
    return redirect(url_for("admin_dashboard"))


def valid_date(value):
    """Return `value` if it is a YYYY-MM-DD date (what <input type="date"> sends), else ""."""
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def audit(action, record=None, details=None):
    """Write one entry of the update history under whoever is signed in (guests have no account)."""
    username = (current_user() or {}).get("username", "khách")
    log_action(username, action, record["stt"] if record else "", record["name"] if record else "", details)


def pending_ca_report():
    pending = state.get("import_pending")
    if pending is None:
        return None
    report = ca_report(pending, plan_import(pending["people"], records_by_citizen_id()), preview=True)
    # People already in the workbook that none of the chosen files lists (matched by CCCD).
    listed = {person["citizen_id"] for person in pending["people"] if person["citizen_id"]}
    missing = [
        {**person, "file": "Excel tổng", "notes": [] if person["citizen_id"] else ["chưa có CCCD, không đối chiếu được"]}
        for person in active_people()
        if person["citizen_id"] not in listed
    ]
    report["groups"].append(("Có trong Excel tổng nhưng không có trong file CA", missing))
    report["counts"]["missing"] = len(missing)
    return report


@app.get("/")
@auth_required
def index():
    refresh_state()
    user = current_user()
    if state["workbook"] is None:
        return render_template_string(PAGE, loaded=False, error=state["error"], csrf=csrf_token(), current_user=user)
    if not user.get("can_read"):
        return render_template_string(
            PAGE,
            loaded=True,
            error=state["error"],
            workbook_name=state["workbook"].name,
            sheet_name=state["sheet"],
            record_count=state["record_count"],
            loaded_at=state["loaded_at"].strftime("%Y-%m-%d %H:%M:%S") if state["loaded_at"] else "",
            records=[],
            filtered_count=0,
            shown_start=0,
            shown_end=0,
            page=1,
            pages=1,
            pagination=None,
            query="",
            filters={},
            filters_active=False,
            options={},
            signature=signature_token(state["signature"]),
            csrf=csrf_token(),
            read_only=state["workbook"].suffix.lower() == ".xlsm",
            dien_options=NvqsStatus.options(),
            unset_dien=UNSET_DIEN_FILTER,
            selection_scope="",
            quarter_choices=QUARTER_CHOICES,
            import_report=None,
            current_user=user,
        )
    query = request.args.get("q", "").strip()
    options = {key: query_filter_options(key) for key in FILTER_DB_COLUMNS}
    # Every option filter can carry several values; keep the canonical spelling of each and drop strangers.
    filters = {}
    for key, values in options.items():
        wanted = {fold(value) for value in request.args.getlist(key)}
        filters[key] = [value for value in values if fold(value) in wanted]
    filters["address"] = request.args.get("address", "").strip()
    for key in ("created_from", "created_to", "updated_from", "updated_to"):
        filters[key] = valid_date(request.args.get(key, ""))
    filters["deleted"] = "1" if request.args.get("deleted") == "1" else ""
    filters["dien"] = [
        code for code in dict.fromkeys(request.args.getlist("dien"))
        if code in NvqsStatus.__members__ or code == UNSET_DIEN_FILTER
    ]
    try:
        page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        page = 1
    start = (page - 1) * MAX_RECORDS_PER_PAGE
    visible, filtered_count = query_records(query, filters, MAX_RECORDS_PER_PAGE, start)
    pages = max(1, (filtered_count + MAX_RECORDS_PER_PAGE - 1) // MAX_RECORDS_PER_PAGE)
    page = min(page, pages)
    corrected_start = (page - 1) * MAX_RECORDS_PER_PAGE
    if corrected_start != start:
        start = corrected_start
        visible, filtered_count = query_records(query, filters, MAX_RECORDS_PER_PAGE, start)
    params = {"q": query, **filters}
    return render_template_string(
        PAGE,
        loaded=True,
        error=state["error"],
        workbook_name=state["workbook"].name,
        sheet_name=state["sheet"],
        record_count=state["record_count"],
        loaded_at=state["loaded_at"].strftime("%Y-%m-%d %H:%M:%S") if state["loaded_at"] else "",
        records=visible,
        filtered_count=filtered_count,
        shown_start=start + 1 if visible else 0,
        shown_end=start + len(visible),
        page=page,
        pages=pages,
        pagination=page_links(page, pages, params),
        query=query,
        filters=filters,
        filters_active=any(filters.values()),
        options=options,
        signature=signature_token(state["signature"]),
        csrf=csrf_token(),
        read_only=state["workbook"].suffix.lower() == ".xlsm",
        dien_options=NvqsStatus.options(),
        unset_dien=UNSET_DIEN_FILTER,
        selection_scope=f"{state['workbook']}:{state['sheet']}",
        quarter_choices=QUARTER_CHOICES,
        import_report=state.pop("import_report", None) or pending_ca_report(),
        current_user=user,
    )


@app.get("/api/status")
@auth_required
def api_status():
    refresh_state()
    return jsonify({
        "loaded": state["workbook"] is not None,
        "signature": signature_token(state["signature"]),
        "record_count": state["record_count"],
        "error": state["error"],
    })


@app.get("/refresh")
@auth_required
def refresh():
    refresh_state(force=True)
    if state["error"]:
        flash(state["error"], "error")
    else:
        flash("Đã làm mới dữ liệu từ tệp Excel.", "success")
    return redirect(url_for("index"))


@app.get("/choose")
@auth_required
def choose_workbook():
    selected = native_workbook_dialog()
    if selected is None:
        flash("Chưa chọn tệp Excel. Hãy dùng mục tải tệp lên nếu hộp chọn tệp không mở được.", "error")
        return redirect(url_for("index"))
    try:
        set_workbook(selected)
        flash(f"Đã tải tệp {selected.name}.", "success")
    except Exception as error:
        logging.exception("Could not select workbook: %s", error)
        flash(str(error), "error")
    return redirect(url_for("index"))


@app.post("/load")
@auth_required
def load_uploaded_workbook():
    upload = request.files.get("workbook")
    if upload is None or not upload.filename:
        flash("Hãy chọn tệp Excel trước.", "error")
        return redirect(url_for("index"))
    filename = secure_filename(upload.filename)
    if Path(filename).suffix.lower() not in SUPPORTED_SUFFIXES:
        flash("Hãy chọn tệp có đuôi .xlsx hoặc .xlsm.", "error")
        return redirect(url_for("index"))
    try:
        UPLOAD_DIR.mkdir(exist_ok=True)
        destination = UPLOAD_DIR / filename
        upload.save(destination)
        set_workbook(destination)
        flash(f"Đã tải bản sao tệp {filename}.", "success")
    except Exception as error:
        logging.exception("Could not load uploaded workbook: %s", error)
        flash(f"Không thể tải tệp Excel: {error}", "error")
    return redirect(url_for("index"))


@app.get("/open-excel")
@auth_required
def open_excel():
    if state["workbook"] is None:
        return redirect(url_for("index"))
    try:
        if os.name == "nt":
            os.startfile(str(state["workbook"]))
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(state["workbook"])])
        else:
            subprocess.Popen(["xdg-open", str(state["workbook"])])
        flash("Đã mở tệp bằng Excel hoặc ứng dụng bảng tính mặc định.", "success")
    except (OSError, FileNotFoundError) as error:
        flash(f"Không thể mở tệp Excel: {error}", "error")
    return redirect(url_for("index"))


@app.get("/person/new")
@auth_required(permission="create")
def new_person():
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    return render_form("Thêm hồ sơ")


@app.get("/person/<stt>/edit")
@auth_required(permission="update")
def edit_person(stt):
    refresh_state()
    record = find_record(stt)
    if record is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    return render_form("Cập nhật hồ sơ", record=record)


@app.get("/person/new/preview")
@auth_required(permission="create")
def new_person_preview():
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    return render_preview("Thêm hồ sơ")


@app.get("/person/<stt>/preview")
@auth_required(permission="read")
def preview_person(stt):
    refresh_state()
    record = find_record(stt)
    if record is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    return render_preview("Xem trước hồ sơ", record=record)


@app.post("/person/save")
@auth_required
def save_person_route():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy mở lại biểu mẫu.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    user = current_user()
    definitions = field_definitions()
    values = {definition["name"]: request.form.get(definition["name"], "") for definition in definitions}
    original_stt = request.form.get("original_stt", "").strip() or None
    if original_stt and not user.get("can_update"):
        flash("Bạn không có quyền chỉnh sửa hồ sơ.", "error")
        return redirect(url_for("index"))
    if not original_stt and not user.get("can_create"):
        flash("Bạn không có quyền thêm mới hồ sơ.", "error")
        return redirect(url_for("index"))
    values_by_column = {definition["column"]: values[definition["name"]] for definition in definitions}
    values_for_validation = {column: raw for column, raw in values_by_column.items()}
    values_for_validation["__original_stt"] = original_stt or ""
    validation_errors = validate_form_values(values_for_validation)
    record = find_record(original_stt) if original_stt else None
    if original_stt and record is None:
        validation_errors.append("Hồ sơ đã thay đổi bên ngoài ứng dụng. Hãy làm mới và mở lại hồ sơ.")
    if state["workbook"].suffix.lower() == ".xlsm":
        validation_errors.append("Ứng dụng không hỗ trợ lưu biểu mẫu vào tệp .xlsm.")
    # Absent (rather than blank) means the posted form had no dropdown at all, so leave the stored status alone.
    nvqs_type = request.form.get("nvqs_type")
    if nvqs_type and nvqs_type not in NvqsStatus.__members__:
        validation_errors.append("Diện nghĩa vụ quân sự không hợp lệ.")
    citizen_column = state["headers"].get("Căn cước")
    citizen_id = values_by_column.get(citizen_column, "").strip() if citizen_column else ""
    if nvqs_type and not citizen_id:
        validation_errors.append("Hãy nhập số CCCD trước khi chọn Diện, vì Diện được lưu theo số CCCD.")
    is_preview = request.form.get("view") == "preview"
    render_page = render_preview if is_preview else render_form
    status_argument = {} if is_preview else {"nvqs_type": nvqs_type or ""}
    if validation_errors:
        return render_page("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=validation_errors, **status_argument)

    headers_by_column = {definition["column"]: definition["header"] for definition in definitions}
    labels_by_name = {definition["name"]: definition["header"] for definition in definitions}
    old_form_values = form_values(record) if record else {}
    updates = {column: coerce_cell_value(headers_by_column[column], raw) for column, raw in values_by_column.items()}
    try:
        row, stt, backup = save_person(
            state["workbook"], state["sheet"], updates, request.form.get("signature", ""), original_stt=original_stt,
        )
        if record is not None:
            move_record_owner(state["workbook"], record["citizen_id"], citizen_id)
            documents.move_documents(record["citizen_id"], citizen_id)
        if nvqs_type is not None and citizen_id:
            set_record_type(state["workbook"], citizen_id, nvqs_type)
        # Audit logging
        username = user.get("username", "unknown")
        name_column = state["headers"].get("TÊN THƯỜNG DÙNG")
        record_name = values.get(f"column_{name_column}", "") if name_column else ""
        if original_stt:
            changes = compute_changes(old_form_values, values, labels_by_name)
            if changes:
                log_action(username, "update", str(stt), record_name, {"changes": changes})
        else:
            log_action(username, "create", str(stt), record_name, {"fields": {labels_by_name.get(k, k): v for k, v in values.items() if v}})
        refresh_state(force=True)
        flash(f"Đã lưu hồ sơ STT {stt}. Bản sao lưu: {backup.name}", "success")
        return redirect(url_for("index"))
    except Exception as error:
        logging.exception("Could not save workbook: %s", error)
        return render_page("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=[str(error)], **status_argument)


@app.post("/generate/<stt>")
@auth_required(permission="read")
def generate(stt):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    record = find_record(stt)
    if state["workbook"] is None or record is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    output_dir = state["workbook"].parent / DEFAULT_OUTPUT_NAME
    try:
        document, values = build_document(state["workbook"], state["sheet"], record["row"], TEMPLATE)
        # Even a single person's form gets its own V-VI (Part B) moved to the
        # front: printed on A3 and folded, the reader hits Part B first.
        booklet = combine_booklet([document], TEMPLATE)
        output_dir.mkdir(parents=True, exist_ok=True)
        output_name = f"{values['record_stt']} - {values['common_name']}" if values["record_stt"] else values["common_name"]
        output = unique_output_path(output_dir, output_name)
        booklet.save(output)
    except Exception as error:
        logging.exception("Could not generate document: %s", error)
        flash(f"Không thể tạo tệp Word: {error}", "error")
        return redirect(url_for("index"))
    state["last_output_dir"] = output_dir
    flash(f"Đã tạo tệp Word: {output.name}.", "success")
    return redirect(url_for("download", filename=output.name))


@app.post("/generate/batch")
@auth_required(permission="read")
def generate_batch():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    selected_stt = request.form.getlist("stt")
    if not selected_stt:
        flash("Hãy chọn ít nhất một hồ sơ để tạo Word hàng loạt.", "error")
        return redirect(url_for("index"))

    records = []
    missing = []
    for stt in selected_stt:
        record = find_record(stt)
        (records if record is not None else missing).append(record if record is not None else stt)
    if missing:
        flash(f"Không tìm thấy hồ sơ STT: {', '.join(missing)}. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))

    output_dir = state["workbook"].parent / DEFAULT_OUTPUT_NAME
    try:
        documents = [
            build_document(state["workbook"], state["sheet"], record["row"], TEMPLATE)[0] for record in records
        ]
        combined = combine_booklet(documents, TEMPLATE)
        output_dir.mkdir(parents=True, exist_ok=True)
        first_stt, last_stt = records[0]["stt"], records[-1]["stt"]
        label = f"{first_stt}-{last_stt}" if first_stt != last_stt else first_stt
        output_name = f"Lô hồ sơ {label} ({len(records)} người)"
        output = unique_output_path(output_dir, output_name)
        combined.save(output)
    except Exception as error:
        logging.exception("Could not generate batch document: %s", error)
        flash(f"Không thể tạo tệp Word hàng loạt: {error}", "error")
        return redirect(url_for("index"))
    state["last_output_dir"] = output_dir
    flash(f"Đã tạo tệp Word gộp cho {len(records)} hồ sơ: {output.name}.", "success")
    return redirect(url_for("download", filename=output.name))


def owner_citizen_id(stt):
    """Resolve the record's CCCD, which is what app-owned data is filed under."""
    record = find_record(stt)
    return record["citizen_id"] if record else None


@app.post("/person/<stt>/files")
@auth_required(permission="update")
def upload_record_file(stt):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    citizen_id = owner_citizen_id(stt)
    if citizen_id is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    if not citizen_id:
        flash("Hồ sơ chưa có số CCCD. Hãy nhập số CCCD và lưu hồ sơ trước khi thêm tài liệu.", "error")
        return redirect(url_for("edit_person", stt=stt))
    upload = request.files.get("document")
    if upload is None or not upload.filename:
        flash("Hãy chọn tệp tài liệu để tải lên.", "error")
        return redirect(url_for("edit_person", stt=stt))
    try:
        # Path().name drops any folder a browser may have sent along; the accents in the name stay.
        filename, data = prepare_document(upload.read(), Path(upload.filename).name)
        filename = documents.save_document(citizen_id, filename, data)
    except AttachmentError as error:
        flash(str(error), "error")
        return redirect(url_for("edit_person", stt=stt))
    except Exception as error:
        logging.exception("Could not store uploaded document: %s", error)
        flash(f"Không thể lưu tài liệu: {error}", "error")
        return redirect(url_for("edit_person", stt=stt))
    audit("update", find_record(stt), {"reason": f"Thêm tài liệu liên quan: {filename}"})
    flash(f"Đã thêm tài liệu {filename} vào hồ sơ liên quan.", "success")
    return redirect(url_for("edit_person", stt=stt))


@app.get("/person/<stt>/documents/<name>")
@auth_required(permission="read")
def record_document_file(stt, name):
    if state["workbook"] is None:
        return redirect(url_for("index"))
    citizen_id = owner_citizen_id(stt)
    path = documents.document_path(citizen_id, name) if citizen_id else None
    if path is None:
        abort(404)
    if request.args.get("size") == "thumb":
        try:
            thumbnail = document_thumbnail(path)
        except AttachmentError as error:
            logging.warning("No thumbnail for %s: %s", path.name, error)
            abort(404)
        response = send_file(io.BytesIO(thumbnail), mimetype="image/jpeg", etag=hashlib.sha1(thumbnail).hexdigest())
    else:
        response = send_file(path, as_attachment=request.args.get("download") == "1", download_name=path.name)
    # Never let the browser reuse what a since-deleted document left behind at this address.
    response.headers["Cache-Control"] = "private, no-cache"
    return response


@app.post("/person/<stt>/documents/<name>/delete")
@auth_required(permission="update")
def remove_record_document(stt, name):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    citizen_id = owner_citizen_id(stt)
    if citizen_id and documents.delete_document(citizen_id, name):
        audit("update", find_record(stt), {"reason": f"Xóa tài liệu liên quan: {name}"})
        flash("Đã xóa tài liệu khỏi hồ sơ liên quan.", "success")
    else:
        flash("Không tìm thấy tài liệu cần xóa.", "error")
    return redirect(url_for("edit_person", stt=stt))


@app.get("/person/<stt>/files/<int:file_id>")
@auth_required(permission="read")
def record_file(stt, file_id):
    if state["workbook"] is None:
        return redirect(url_for("index"))
    citizen_id = owner_citizen_id(stt)
    data = record_file_image(state["workbook"], citizen_id, file_id, thumbnail=request.args.get("size") == "thumb") if citizen_id else None
    if data is None:
        abort(404)
    response = Response(data, mimetype="image/jpeg")
    # Tag by content and force revalidation: the browser must never reuse the picture that a
    # since-deleted document left behind at this address.
    response.set_etag(hashlib.sha1(data).hexdigest())
    response.headers["Cache-Control"] = "private, no-cache"
    return response.make_conditional(request)


@app.post("/person/<stt>/files/<int:file_id>/delete")
@auth_required(permission="update")
def remove_record_file(stt, file_id):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    citizen_id = owner_citizen_id(stt)
    if citizen_id and delete_record_file(state["workbook"], citizen_id, file_id):
        flash("Đã xóa một trang tài liệu khỏi hồ sơ liên quan.", "success")
    else:
        flash("Không tìm thấy trang tài liệu cần xóa.", "error")
    return redirect(url_for("edit_person", stt=stt))


def apply_to_selection(field, value, label, shown_value):
    """Set one manager-chosen field on every ticked record, filed under each owner's CCCD."""
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    selected_stt = request.form.getlist("stt")
    if not selected_stt:
        flash("Hãy chọn ít nhất một hồ sơ để sửa hàng loạt.", "error")
        return redirect(url_for("index"))

    changed, missing, without_citizen_id = 0, [], []
    for stt in selected_stt:
        record = find_record(stt)
        if record is None:
            missing.append(stt)
        elif not record["citizen_id"]:
            without_citizen_id.append(stt)
        else:
            previous = record["type" if field == "type" else field]
            set_record_attribute(state["workbook"], record["citizen_id"], field, value)
            audit("update", record, {"changes": [{"field": label, "from": previous or "", "to": shown_value or ""}]})
            changed += 1
    if changed:
        refresh_state(force=True)
        if shown_value:
            flash(f"Đã đặt {label} \"{shown_value}\" cho {changed} hồ sơ.", "success")
        else:
            flash(f"Đã xóa {label} của {changed} hồ sơ.", "success")
    if without_citizen_id:
        flash(
            f"Bỏ qua hồ sơ STT {', '.join(without_citizen_id)} vì chưa có số CCCD — {label} được lưu theo số CCCD.",
            "error",
        )
    if missing:
        flash(f"Không tìm thấy hồ sơ STT: {', '.join(missing)}. Hãy làm mới tệp Excel và thử lại.", "error")
    return redirect(url_for("index"))


@app.post("/bulk/dien")
@auth_required(permission="update")
def bulk_update_dien():
    value = request.form.get("bulk_dien", "").strip()
    if value and value not in NvqsStatus.__members__:
        flash("Diện nghĩa vụ quân sự không hợp lệ.", "error")
        return redirect(url_for("index"))
    return apply_to_selection("type", value, "Diện", NvqsStatus.value_for(value))


@app.post("/bulk/khu-pho")
@auth_required(permission="update")
def bulk_update_quarter():
    value = request.form.get("bulk_quarter", "").strip()
    if value and value not in QUARTER_CHOICES:
        flash("Khu phố không hợp lệ.", "error")
        return redirect(url_for("index"))
    return apply_to_selection("quarter", value, "Khu phố", value)


def ca_report(pending, plan, backup="", preview=False):
    """What the import page shows, both as a preview before adding and as the result afterwards."""
    return {
        "preview": preview,
        "files": pending["files"],
        "read": len(pending["people"]),
        "backup": backup,
        "errors": pending["errors"],
        "counts": {key: len(value) for key, value in plan.items()},
        "groups": [
            ("Sẽ được thêm mới vào Excel tổng" if preview else "Đã thêm mới vào Excel tổng", plan["add"]),
            ("Sẽ được cập nhật theo file CA (đã có trong Excel tổng)" if preview else "Đã cập nhật theo file CA", plan["update"]),
            ("Đã có trong Excel tổng, dữ liệu không khác — giữ nguyên", plan["existing"]),
            ("Không có số CCCD — bỏ qua, hãy bổ sung CCCD rồi nhập lại", plan["no_citizen_id"]),
            ("Trùng CCCD ngay trong các tệp vừa chọn — chỉ lấy dòng đầu", plan["repeated"]),
            ("Có thay đổi trong file CA nhưng không cập nhật — đã chọn chỉ thêm mới", plan.get("not_updated", [])),
        ],
    }


@app.post("/import/ca")
@auth_required(permission="create")
def import_ca_lists():
    """Read the Ministry's khu phố lists and show who would be added; nothing is written yet."""
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    uploads = [upload for upload in request.files.getlist("ca_files") if upload.filename]
    if not uploads:
        flash("Hãy chọn ít nhất một tệp Excel của bộ CA.", "error")
        return redirect(url_for("index"))
    try:
        people, file_errors = read_ca_files([(Path(upload.filename).name, upload.read()) for upload in uploads])
    except Exception as error:
        logging.exception("Could not read CA lists: %s", error)
        flash(f"Không thể đọc file CA: {error}", "error")
        return redirect(url_for("index"))
    state["import_pending"] = {"files": len(uploads), "people": people, "errors": file_errors}
    return redirect(url_for("index"))


@app.post("/import/ca/confirm")
@auth_required(permission="create")
def import_ca_confirm():
    """Add the people from the previewed CA lists that the workbook still does not have."""
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    pending = state.get("import_pending")
    if state["workbook"] is None or pending is None:
        flash("Không có file CA nào đang chờ thêm. Hãy chọn lại tệp.", "error")
        return redirect(url_for("index"))
    try:
        # Plan again: the workbook may have gained some of these people since the preview.
        plan = plan_import(pending["people"], records_by_citizen_id())
        if request.form.get("only_add"):
            # "Chỉ thêm mới": people already held stay exactly as they are; the report still lists what differed.
            plan["not_updated"], plan["update"] = plan["update"], []
        backup = None
        if plan["add"] or plan["update"]:
            updates = [
                (person["row"], person["citizen_id"], {c["header"]: c["value"] for c in person["changes"] if c["header"]})
                for person in plan["update"]
            ]
            stts, backup = append_people(
                state["workbook"], state["sheet"], [person["values"] for person in plan["add"]],
                signature_token(state["signature"]), updates,
            )
            for person in plan["update"]:
                audit(
                    "update", {"stt": person["stt"], "name": person["name"]},
                    {"changes": [{"field": c["field"], "from": c["old"], "to": c["new"]} for c in person["changes"]],
                     "reason": f"Cập nhật từ file CA ({person['file']})"},
                )
                for change in person["changes"]:
                    if not change["header"]:
                        set_record_attribute(state["workbook"], person["citizen_id"], "quarter", change["value"])
            for person, stt in zip(plan["add"], stts):
                person["stt"] = stt
                audit("create", {"stt": stt, "name": person["values"]["TÊN THƯỜNG DÙNG"]}, {"reason": f"Nhập từ file CA ({person['file']})"})
                if person["quarter"]:
                    # Saved before the refresh below, so the new records pick their khu phố up as they load.
                    set_record_attribute(state["workbook"], person["citizen_id"], "quarter", person["quarter"])
            refresh_state(force=True)
    except Exception as error:
        logging.exception("Could not import CA lists: %s", error)
        flash(f"Không thể thêm người từ file CA: {error}", "error")
        return redirect(url_for("index"))
    state.pop("import_pending", None)
    state["import_report"] = ca_report(pending, plan, backup.name if backup else "")
    return redirect(url_for("index"))


@app.post("/import/ca/cancel")
@auth_required(permission="create")
def import_ca_cancel():
    if valid_csrf(request.form.get("csrf_token")):
        state.pop("import_pending", None)
    return redirect(url_for("index"))


def change_deleted_flag(deleted):
    """Soft-delete or restore every ticked record, filed under each owner's CCCD."""
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    selected_stt = request.form.getlist("stt")
    if not selected_stt:
        flash("Hãy chọn ít nhất một hồ sơ.", "error")
        return redirect(url_for("index"))

    citizen_ids, missing, without_citizen_id, records = [], [], [], []
    for stt in selected_stt:
        record = find_record(stt)
        if record is None:
            missing.append(stt)
        elif not record["citizen_id"]:
            without_citizen_id.append(stt)
        else:
            citizen_ids.append(record["citizen_id"])
            records.append(record)
    if citizen_ids:
        set_records_deleted(state["workbook"], citizen_ids, deleted)
        for record in records:
            audit("delete" if deleted else "restore", record, {"reason": "Xóa mềm" if deleted else "Khôi phục"})
        refresh_state(force=True)
        action = "Đã xóa" if deleted else "Đã khôi phục"
        flash(f"{action} {len(citizen_ids)} hồ sơ." + (" Hồ sơ vẫn còn trong tệp Excel và có thể khôi phục." if deleted else ""), "success")
    if without_citizen_id:
        flash(f"Bỏ qua hồ sơ STT {', '.join(without_citizen_id)} vì chưa có số CCCD — trạng thái xóa được lưu theo số CCCD.", "error")
    if missing:
        flash(f"Không tìm thấy hồ sơ STT: {', '.join(missing)}. Hãy làm mới tệp Excel và thử lại.", "error")
    return redirect(url_for("index"))


@app.post("/bulk/delete")
@auth_required(permission="delete")
def bulk_delete():
    return change_deleted_flag(True)


@app.post("/bulk/restore")
@auth_required(permission="delete")
def bulk_restore():
    return change_deleted_flag(False)


@app.post("/export/batch")
@auth_required(permission="export")
def export_batch_excel():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    selected_stt = request.form.getlist("stt")
    if not selected_stt:
        flash("Hãy chọn ít nhất một hồ sơ để tạo Excel hàng loạt.", "error")
        return redirect(url_for("index"))
    if not LIST_TEMPLATE.is_file():
        flash(f"Không tìm thấy tệp mẫu {LIST_TEMPLATE.name} cạnh ứng dụng.", "error")
        return redirect(url_for("index"))

    records = []
    missing = []
    for stt in selected_stt:
        record = find_record(stt)
        (records if record is not None else missing).append(record if record is not None else stt)
    if missing:
        flash(f"Không tìm thấy hồ sơ STT: {', '.join(missing)}. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))

    output_dir = state["workbook"].parent / DEFAULT_OUTPUT_NAME
    try:
        listing = build_list_workbook(records, LIST_TEMPLATE)
        output_dir.mkdir(parents=True, exist_ok=True)
        first_stt, last_stt = records[0]["stt"], records[-1]["stt"]
        label = f"{first_stt}-{last_stt}" if first_stt != last_stt else first_stt
        output = unique_output_path(output_dir, f"Danh sách thanh niên {label} ({len(records)} người)", suffix=".xlsx")
        listing.save(output)
    except Exception as error:
        logging.exception("Could not export record list: %s", error)
        flash(f"Không thể tạo tệp Excel hàng loạt: {error}", "error")
        return redirect(url_for("index"))
    state["last_output_dir"] = output_dir
    audit("export", None, {"reason": f"Xuất Excel hàng loạt {len(records)} hồ sơ: {output.name}"})
    flash(f"Đã tạo tệp Excel danh sách cho {len(records)} hồ sơ: {output.name}.", "success")
    return redirect(url_for("download", filename=output.name))


@app.post("/export/dien")
@auth_required(permission="export")
def export_by_dien():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    chosen = [code for code in request.form.getlist("export_dien") if code in NvqsStatus.__members__]
    if not chosen:
        flash("Hãy chọn ít nhất một diện để xuất Excel.", "error")
        return redirect(url_for("index"))

    output_dir = state["workbook"].parent / DEFAULT_OUTPUT_NAME
    written, empty, missing_template = [], [], []
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        for code in chosen:
            label = NvqsStatus.value_for(code)
            template = DIEN_TEMPLATE_DIR / DIEN_TEMPLATES[code]
            if not template.is_file():
                missing_template.append(template.name)
                continue
            records, _ = query_records("", {"dien": [code]}, MAX_EXPORT_RECORDS, 0)
            if not records:
                empty.append(label)
                continue
            listing = build_dien_workbook(records, template)
            output = unique_output_path(output_dir, f"{label} ({len(records)} người)", suffix=".xlsx")
            listing.save(output)
            written.append(output)
    except Exception as error:
        logging.exception("Could not export by diện: %s", error)
        flash(f"Không thể tạo tệp Excel theo diện: {error}", "error")
        return redirect(url_for("index"))

    for name in missing_template:
        flash(f"Không tìm thấy tệp mẫu {name} trong thư mục template.", "error")
    if empty:
        flash(f"Không có hồ sơ nào thuộc diện: {', '.join(empty)}.", "error")
    if not written:
        return redirect(url_for("index"))

    audit("export", None, {"reason": f"Xuất Excel theo diện: {', '.join(path.name for path in written)}"})
    state["last_output_dir"] = output_dir
    if len(written) == 1:
        flash(f"Đã tạo tệp Excel: {written[0].name}.", "success")
        return redirect(url_for("download", filename=written[0].name))

    # One file per diện, handed over together so the browser can deliver them in a single download.
    bundle = unique_output_path(output_dir, f"Danh sách theo diện ({len(written)} tệp)", suffix=".zip")
    with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as archive:
        for path in written:
            archive.write(path, path.name)
    flash(f"Đã tạo {len(written)} tệp Excel theo diện, tải về trong {bundle.name}.", "success")
    return redirect(url_for("download", filename=bundle.name))


@app.get("/downloads/<path:filename>")
@auth_required(permission="read")
def download(filename):
    # Word files are "in" (print); anything else a person can download is an exported Excel/ZIP file.
    if Path(filename).suffix.lower() != ".docx" and not current_user().get("can_export"):
        flash("Bạn không có quyền xuất tệp Excel.", "error")
        return redirect(url_for("index"))
    output_dir = state.get("last_output_dir")
    if output_dir is None or not output_dir.is_dir():
        flash("Không tìm thấy thư mục chứa tệp Word đã tạo.", "error")
        return redirect(url_for("index"))
    return send_from_directory(output_dir, filename, as_attachment=True)


@app.get("/chung-chi")
def https_ca_certificate():
    """The LAN CA certificate, so a phone can install it once and stop showing the HTTPS warning."""
    ca_path = HTTPS_DIR / "ca.pem"
    if not ca_path.is_file():
        abort(404)
    return send_file(ca_path, mimetype="application/x-x509-ca-cert", as_attachment=True, download_name="QuanLyHoSo-CA.crt")


@app.get("/tra-cuu")
def guest_lookup():
    refresh_state()
    if state["workbook"] is None:
        return render_template_string(
            GUEST_PAGE,
            mode=None, cccd="", record=None, sections=[], values={}, errors=[], 
            signature="", csrf=csrf_token(), cccd_field_name="",
        )
    return render_guest_form(mode=None, cccd="")


@app.post("/tra-cuu")
def guest_lookup_post():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("guest_lookup"))
    refresh_state()
    if state["workbook"] is None:
        flash("Hệ thống chưa sẵn sàng. Vui lòng liên hệ quản trị viên.", "error")
        return redirect(url_for("guest_lookup"))
    cccd = request.form.get("cccd", "").strip()
    if not cccd:
        flash("Vui lòng nhập số CCCD.", "error")
        return redirect(url_for("guest_lookup"))
    if not valid_citizen_id(cccd):
        flash(CITIZEN_ID_ERROR, "error")
        return redirect(url_for("guest_lookup"))
    if guest_lookups.blocked(client_ip()):
        flash(TOO_MANY_ATTEMPTS, "error")
        return redirect(url_for("guest_lookup"))
    guest_lookups.record(client_ip())
    record = find_record_by_citizen_id(cccd)
    if record:
        return render_guest_form(mode="readonly", cccd=cccd, record=record)
    # No profile yet: ask first, so nobody lands in a long form by accident.
    return render_guest_form(mode="confirm", cccd=cccd)


@app.post("/tra-cuu/them")
def guest_start_create():
    """The guest said yes to adding their own profile: show the form."""
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("guest_lookup"))
    refresh_state()
    if state["workbook"] is None:
        flash("Hệ thống chưa sẵn sàng. Vui lòng liên hệ quản trị viên.", "error")
        return redirect(url_for("guest_lookup"))
    cccd = request.form.get("cccd", "").strip()
    if not cccd:
        flash("Vui lòng nhập số CCCD.", "error")
        return redirect(url_for("guest_lookup"))
    if not valid_citizen_id(cccd):
        flash(CITIZEN_ID_ERROR, "error")
        return redirect(url_for("guest_lookup"))
    if guest_lookups.blocked(client_ip()):
        flash(TOO_MANY_ATTEMPTS, "error")
        return redirect(url_for("guest_lookup"))
    guest_lookups.record(client_ip())
    record = find_record_by_citizen_id(cccd)
    if record:
        return render_guest_form(mode="readonly", cccd=cccd, record=record)
    return render_guest_form(mode="create", cccd=cccd)


@app.post("/tra-cuu/save")
def guest_save():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("guest_lookup"))
    refresh_state()
    if state["workbook"] is None:
        flash("Hệ thống chưa sẵn sàng. Vui lòng liên hệ quản trị viên.", "error")
        return redirect(url_for("guest_lookup"))
    if state["workbook"].suffix.lower() == ".xlsm":
        flash("Hệ thống không hỗ trợ lưu hồ sơ vào tệp .xlsm.", "error")
        return redirect(url_for("guest_lookup"))
    cccd = request.form.get("cccd", "").strip()
    if not cccd:
        flash("Thiếu số CCCD.", "error")
        return redirect(url_for("guest_lookup"))
    if not valid_citizen_id(cccd):
        flash(CITIZEN_ID_ERROR, "error")
        return redirect(url_for("guest_lookup"))
    if guest_saves.blocked(client_ip()):
        flash(TOO_MANY_ATTEMPTS, "error")
        return redirect(url_for("guest_lookup"))
    # Check if CCCD already exists (prevent duplicate)
    existing = find_record_by_citizen_id(cccd)
    if existing:
        flash("Hồ sơ với số CCCD này đã tồn tại.", "error")
        return render_guest_form(mode="readonly", cccd=cccd, record=existing)
    definitions = field_definitions()
    values = {definition["name"]: request.form.get(definition["name"], "") for definition in definitions}
    values_by_column = {definition["column"]: values[definition["name"]] for definition in definitions}
    values_for_validation = {column: raw for column, raw in values_by_column.items()}
    values_for_validation["__original_stt"] = ""
    validation_errors = validate_form_values(values_for_validation)
    if validation_errors:
        return render_guest_form(mode="create", cccd=cccd, values=values, errors=validation_errors)
    headers_by_column = {definition["column"]: definition["header"] for definition in definitions}
    updates = {column: coerce_cell_value(headers_by_column[column], raw) for column, raw in values_by_column.items()}
    try:
        row, stt, backup = save_person(
            state["workbook"], state["sheet"], updates, request.form.get("signature", ""),
        )
        refresh_state(force=True)
        guest_saves.record(client_ip())
        audit("create", find_record_by_citizen_id(cccd), {"reason": f"Khách tự nhập (CCCD {cccd})"})
        flash(f"Đã tạo hồ sơ thành công (STT {stt}). Cảm ơn bạn!", "success")
        # After saving, look up the new record and show it in readonly mode
        new_record = find_record_by_citizen_id(cccd)
        if new_record:
            return render_guest_form(mode="readonly", cccd=cccd, record=new_record)
        return redirect(url_for("guest_lookup"))
    except Exception as error:
        logging.exception("Guest could not save record: %s", error)
        return render_guest_form(mode="create", cccd=cccd, values=values, errors=[str(error)])


@app.post("/tra-cuu/cap-nhat")
@app.post("/tra-cuu/update")
def guest_update_request():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy thử lại.", "error")
        return redirect(url_for("guest_lookup"))
    refresh_state()
    if state["workbook"] is None:
        flash("Hệ thống chưa sẵn sàng. Vui lòng liên hệ quản trị viên.", "error")
        return redirect(url_for("guest_lookup"))
    cccd = request.form.get("cccd", "").strip()
    if not cccd:
        flash("Thiếu số CCCD.", "error")
        return redirect(url_for("guest_lookup"))
    if not valid_citizen_id(cccd):
        flash(CITIZEN_ID_ERROR, "error")
        return redirect(url_for("guest_lookup"))
    record = find_record_by_citizen_id(cccd)
    if not record:
        flash("Không tìm thấy hồ sơ thanh niên với số CCCD này.", "error")
        return redirect(url_for("guest_lookup"))

    # Temporary Lockout (Spam Prevention)
    from quan_ly_ho_so.approvals import (
        IDENTITY_CRITICAL_HEADERS,
        create_profile_edit_request,
        has_pending_request,
    )
    if has_pending_request(cccd):
        flash("Hồ sơ đang có một yêu cầu cập nhật chờ xét duyệt. Bạn tạm thời không thể gửi thêm yêu cầu mới.", "error")
        return render_guest_form(mode="readonly", cccd=cccd, record=record)

    definitions = field_definitions()
    labels_by_name = {definition["name"]: definition["header"] for definition in definitions}
    raw_values = {}
    for definition in definitions:
        name = definition["name"]
        header = definition["header"]
        if name in request.form:
            raw_values[name] = request.form[name]
        else:
            raw_values[name] = display_value(record["data"].get(header))

    # Strictly prohibit user changes to identity-critical fields:
    # Always force them to match the existing record's original data!
    for definition in definitions:
        header = definition["header"]
        if header in IDENTITY_CRITICAL_HEADERS:
            raw_values[definition["name"]] = display_value(record["data"].get(header))

    # Validate general fields
    values_by_column = {definition["column"]: raw_values[definition["name"]] for definition in definitions}
    values_for_validation = {column: raw for column, raw in values_by_column.items()}
    values_for_validation["__original_stt"] = record["stt"]
    validation_errors = validate_form_values(values_for_validation)
    if validation_errors:
        return render_guest_form(mode="readonly", cccd=cccd, record=record, values=raw_values, errors=validation_errors)

    # Compute changes (diff view)
    old_values = form_values(record)
    changes = compute_changes(old_values, raw_values, labels_by_name)

    # Ensure no identity-critical fields are in the change diff
    changes = [c for c in changes if c["field"] not in IDENTITY_CRITICAL_HEADERS]

    if not changes:
        flash("Không có thông tin nào thay đổi để gửi yêu cầu cập nhật.", "error")
        return render_guest_form(mode="readonly", cccd=cccd, record=record, values=raw_values)

    try:
        name_header = "TÊN THƯỜNG DÙNG"
        record_name = record["data"].get(name_header, "") or record["name"]
        req_id = create_profile_edit_request(
            cccd=cccd,
            stt=str(record["stt"]),
            record_name=record_name,
            submitted_values={labels_by_name[name]: value for name, value in raw_values.items()},
            original_values={labels_by_name[name]: value for name, value in old_values.items()},
            changes=changes,
        )
        flash("Yêu cầu cập nhật thông tin đã được gửi thành công đến Quản trị viên để xét duyệt.", "success")
        return render_guest_form(mode="readonly", cccd=cccd, record=record)
    except Exception as error:
        logging.exception("Could not create profile edit request: %s", error)
        flash(f"Lỗi khi gửi yêu cầu cập nhật: {error}", "error")
        return render_guest_form(mode="readonly", cccd=cccd, record=record, values=raw_values)


@app.errorhandler(413)
def request_too_large(error):
    return redirect(url_for("index"))


def lan_address():
    """This computer's address on the local network, as other computers would reach it."""
    probe = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        probe.connect(("10.255.255.255", 1))  # UDP: nothing is sent, the OS just picks the outgoing interface
        return probe.getsockname()[0]
    except OSError as error:
        logging.warning("Could not work out the LAN address: %s", error)
        return "<địa chỉ IP của máy này>"
    finally:
        probe.close()


def lan_hostname():
    """This computer's mDNS name (name.local). Unlike the LAN address it stays the same after reconnecting to the Wi-Fi."""
    if sys.platform == "darwin":
        # macOS keeps the Bonjour name apart from the hostname, which the network's DHCP may override.
        try:
            name = subprocess.run(["scutil", "--get", "LocalHostName"], capture_output=True, text=True, check=True).stdout.strip()
            if name:
                return f"{name}.local"
        except (OSError, subprocess.CalledProcessError) as error:
            logging.warning("Could not read the Bonjour name: %s", error)
    return f"{socket.gethostname().split('.')[0]}.local"


@app.get("/audit-log")
@app.get("/admin/audit-log")
@auth_required(role="admin")
def admin_audit_log():
    from quan_ly_ho_so.audit import get_audit_logs, get_audit_performers
    import urllib.parse

    try:
        page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        page = 1

    q = (request.args.get("q") or request.args.get("search") or request.args.get("query") or "").strip()
    stt = (request.args.get("stt") or "").strip() or None
    username = (request.args.get("username") or request.args.get("user") or request.args.get("nguoi_thuc_hien") or "").strip() or None
    action = (request.args.get("action") or request.args.get("hanh_dong") or "").strip() or None
    date_from = (request.args.get("date_from") or request.args.get("from_date") or request.args.get("tu_ngay") or "").strip() or None
    date_to = (request.args.get("date_to") or request.args.get("to_date") or request.args.get("den_ngay") or "").strip() or None
    record_name = (request.args.get("record_name") or request.args.get("ho_ten") or request.args.get("name") or "").strip() or None

    per_page = 50
    entries, total = get_audit_logs(
        limit=per_page,
        offset=(page - 1) * per_page,
        stt=stt,
        username=username,
        action=action,
        date_from=date_from,
        date_to=date_to,
        query=q or None,
        record_name=record_name,
    )
    pages = max(1, (total + per_page - 1) // per_page)
    performers = get_audit_performers()

    has_filters = bool(q or stt or username or action or date_from or date_to or record_name)

    def audit_page_url(target_page: int) -> str:
        params = dict(request.args)
        params["page"] = str(target_page)
        cleaned = {k: v for k, v in params.items() if v}
        return f"{request.path}?{urllib.parse.urlencode(cleaned)}"

    if pages <= 9:
        page_range = list(range(1, pages + 1))
    else:
        raw_range = sorted({1, 2, pages - 1, pages, max(1, page - 1), page, min(pages, page + 1)})
        page_range = []
        prev = 0
        for p in raw_range:
            if prev and p - prev == 2:
                page_range.append(prev + 1)
            elif prev and p - prev > 2:
                page_range.append("...")
            page_range.append(p)
            prev = p

    return render_template_string(
        AUDIT_LOG_PAGE,
        entries=entries,
        total=total,
        page=page,
        pages=pages,
        page_range=page_range,
        per_page=per_page,
        stt=stt or "",
        q=q,
        username=username or "",
        action=action or "",
        date_from=date_from or "",
        date_to=date_to or "",
        record_name=record_name or "",
        performers=performers,
        has_filters=has_filters,
        audit_page_url=audit_page_url,
        reset_url=request.path,
        csrf=csrf_token(),
        current_user=current_user(),
    )


def ensure_admin_password():
    """Opened to the LAN or the internet, the admin must not keep the default password: ask for a new one here."""
    admin = get_user_by_username(DEFAULT_ADMIN_USERNAME)
    if not admin or not check_password_hash(admin["password_hash"], DEFAULT_ADMIN_PASSWORD):
        return
    if not sys.stdin.isatty():
        raise SystemExit(
            "Tài khoản quản trị vẫn dùng mật khẩu mặc định nên chưa thể mở cho người khác. "
            "Hãy chạy lại trong cửa sổ dòng lệnh để đặt mật khẩu mới."
        )
    print(f"Tài khoản quản trị '{DEFAULT_ADMIN_USERNAME}' đang dùng mật khẩu mặc định. Hãy đặt mật khẩu mới (ít nhất 8 ký tự).")
    while True:
        new_password = getpass.getpass("Mật khẩu mới: ")
        if len(new_password) < 8:
            print("Mật khẩu phải có ít nhất 8 ký tự.")
        elif new_password == DEFAULT_ADMIN_PASSWORD:
            print("Không được dùng lại mật khẩu mặc định.")
        elif new_password != getpass.getpass("Nhập lại mật khẩu: "):
            print("Hai lần nhập không khớp.")
        else:
            break
    update_user(admin["id"], password=new_password, must_change_password=False)
    print("Đã đổi mật khẩu quản trị.")


def main():
    if not TEMPLATE.is_file():
        raise SystemExit("Mau_Ho_So_Thanh_Nien.docx phải nằm cùng thư mục với ứng dụng.")
    auth_database_connection().close()
    if "--setup" in sys.argv[1:]:
        ensure_admin_password()
        return
    # QLHS_HOST=0.0.0.0 opens the app to the local network; by default only this computer can reach it.
    host = os.environ.get("QLHS_HOST") or os.environ.get("FLASK_RUN_HOST", "127.0.0.1")
    port = int(os.environ.get("FLASK_RUN_PORT", "3000"))
    debug = os.environ.get("FLASK_DEBUG", "0").lower() in ("1", "true", "yes")
    open_browser = os.environ.get("OPEN_BROWSER", "1").lower() in ("1", "true", "yes")
    shared = host not in ("127.0.0.1", "localhost", "::1")
    if shared or online_mode():
        ensure_admin_password()
    load_saved_workbook()
    # LAN mode serves https (Chrome on phones insists on it); online mode stays http because the tunnel adds https.
    ssl_context = None
    scheme = "http"
    if shared and not online_mode():
        address, hostname = lan_address(), lan_hostname()
        ssl_context = lan_ssl_context(*lan_certificate(HTTPS_DIR, address, hostname))
        scheme = "https"
        app.config.update(SESSION_COOKIE_SECURE=True)
        logging.info("Máy khác trong mạng nội bộ vào bằng: https://%s:%d (cần tài khoản đăng nhập)", hostname, port)
        logging.info("Địa chỉ trên không đổi khi bật lại wifi. Nếu thiết bị không mở được thì dùng: https://%s:%d", address, port)
        logging.info("Lần đầu trình duyệt báo 'không riêng tư': bấm Nâng cao -> Tiếp tục. Muốn hết cảnh báo, cài chứng chỉ tại https://%s:%d/chung-chi", hostname, port)
    if open_browser:
        local_host = "127.0.0.1" if shared else host
        threading.Timer(0.7, lambda: webbrowser.open(f"{scheme}://{local_host}:{port}")).start()
    app.run(host=host, port=port, debug=debug, ssl_context=ssl_context)



if __name__ == "__main__":
    main()

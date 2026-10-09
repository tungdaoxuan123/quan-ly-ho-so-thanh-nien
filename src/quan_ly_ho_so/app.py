#!/usr/bin/env python3
"""Local browser UI for Excel-backed youth records."""

from __future__ import annotations

import logging
import os
import secrets
import subprocess
import sys
import threading
import webbrowser
from pathlib import Path

from datetime import timedelta

from flask import Flask, flash, jsonify, redirect, render_template_string, request, send_from_directory, session, url_for
from werkzeug.security import check_password_hash
from werkzeug.utils import secure_filename

from quan_ly_ho_so.auth import (
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
    FILTER_DB_COLUMNS,
    FORM_GROUPS,
    MAX_RECORDS_PER_PAGE,
    SUPPORTED_SUFFIXES,
    TEMPLATE,
    UPLOAD_DIR,
)
from quan_ly_ho_so.audit import compute_changes, get_audit_logs, log_action
from quan_ly_ho_so.errors import StaleWorkbookError
from quan_ly_ho_so.forms.fields import coerce_cell_value, field_definitions, form_values, validate_form_values
from quan_ly_ho_so.security import csrf_token, valid_csrf
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import fold
from quan_ly_ho_so.web.templates import (
    ADMIN_PAGE,
    AUDIT_LOG_PAGE,
    FORCE_CHANGE_PASSWORD_PAGE,
    GUEST_PAGE,
    LOGIN_PAGE,
    PAGE,
    PROFILE_PAGE,
    page_url,
    render_form,
    render_guest_form,
    render_preview,
)
from quan_ly_ho_so.word.export import build_document, combine_booklet, unique_output_path
from quan_ly_ho_so.workbook.cache import (
    database_record_count,
    find_record,
    find_record_by_citizen_id,
    query_filter_options,
    query_records,
    signature_token,
)
from quan_ly_ho_so.workbook.manager import (
    load_saved_workbook,
    native_workbook_dialog,
    refresh_state,
    set_workbook,
)
from quan_ly_ho_so.workbook.writer import delete_person, save_person

app = Flask(__name__)
app.secret_key = get_persistent_secret_key()
app.permanent_session_lifetime = timedelta(days=DEFAULT_SESSION_LIFETIME_DAYS)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


@app.context_processor
def inject_user():
    return {"current_user": current_user()}


@app.before_request
def require_authentication():
    exempt_endpoints = {"login", "login_post", "logout", "static", "guest_lookup", "guest_lookup_post", "guest_save"}
    if request.endpoint in exempt_endpoints or request.endpoint is None:
        return None
    user = current_user()
    if not user:
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
    user = authenticate(username, password)
    if user is None:
        flash("Tên đăng nhập hoặc mật khẩu không chính xác.", "error")
        return render_template_string(
            LOGIN_PAGE,
            csrf=csrf_token(),
            next_url=next_url,
            username=username,
        )
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
    try:
        create_user(
            username,
            password=DEFAULT_USER_PASSWORD,
            can_create=can_create,
            can_read=can_read,
            can_update=can_update,
            can_delete=can_delete,
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
    try:
        updated = update_user(
            user_id,
            password=None,
            can_create=can_create,
            can_read=can_read,
            can_update=can_update,
            can_delete=can_delete,
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
            page_urls={},
            query="",
            filters={},
            filters_active=False,
            options={},
            signature=signature_token(state["signature"]),
            csrf=csrf_token(),
            read_only=state["workbook"].suffix.lower() == ".xlsm",
            current_user=user,
        )
    query = request.args.get("q", "").strip()
    filters = {
        "birth_year": request.args.get("birth_year", "").strip(),
        "occupation": request.args.get("occupation", "").strip(),
        "education": request.args.get("education", "").strip(),
        "ethnicity": request.args.get("ethnicity", "").strip(),
        "religion": request.args.get("religion", "").strip(),
        "address": request.args.get("address", "").strip(),
    }
    options = {key: query_filter_options(key) for key in FILTER_DB_COLUMNS}
    for key, values in options.items():
        filters[key] = next((value for value in values if fold(value) == fold(filters[key])), filters[key])
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
        page_urls={number: page_url(number, params) for number in range(1, pages + 1)},
        query=query,
        filters=filters,
        filters_active=any(filters.values()),
        options=options,
        signature=signature_token(state["signature"]),
        csrf=csrf_token(),
        read_only=state["workbook"].suffix.lower() == ".xlsm",
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
    render_page = render_preview if request.form.get("view") == "preview" else render_form
    if validation_errors:
        return render_page("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=validation_errors)

    headers_by_column = {definition["column"]: definition["header"] for definition in definitions}
    labels_by_name = {definition["name"]: definition["header"] for definition in definitions}
    old_form_values = form_values(record) if record else {}
    updates = {column: coerce_cell_value(headers_by_column[column], raw) for column, raw in values_by_column.items()}
    try:
        row, stt, backup = save_person(
            state["workbook"], state["sheet"], updates, request.form.get("signature", ""), original_stt=original_stt,
        )
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
        return render_page("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=[str(error)])


@app.post("/person/<stt>/delete")
@auth_required(permission="delete")
def delete_person_route(stt):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    if state["workbook"].suffix.lower() == ".xlsm":
        flash("Ứng dụng không hỗ trợ xóa hồ sơ trên tệp .xlsm. Hãy dùng Excel để chỉnh sửa.", "error")
        return redirect(url_for("index"))
    record = find_record(stt)
    if record is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    try:
        row, deleted_stt, backup = delete_person(
            state["workbook"], state["sheet"], stt, request.form.get("signature", "")
        )
        # Audit logging
        username = current_user().get("username", "unknown")
        log_action(username, "delete", str(deleted_stt), record["name"], {"reason": "soft_delete"})
        refresh_state(force=True)
        flash(f"Đã xóa hồ sơ STT {deleted_stt} (xóa mềm). Bản sao lưu: {backup.name}", "success")
        return redirect(url_for("index"))
    except Exception as error:
        logging.exception("Could not delete record: %s", error)
        flash(f"Không thể xóa hồ sơ: {error}", "error")
        return redirect(url_for("index"))


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


@app.get("/downloads/<path:filename>")
@auth_required(permission="read")
def download(filename):
    output_dir = state.get("last_output_dir")
    if output_dir is None or not output_dir.is_dir():
        flash("Không tìm thấy thư mục chứa tệp Word đã tạo.", "error")
        return redirect(url_for("index"))
    return send_from_directory(output_dir, filename, as_attachment=True)


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
    record = find_record_by_citizen_id(cccd)
    if record:
        return render_guest_form(mode="readonly", cccd=cccd, record=record)
    else:
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
        flash(f"Đã tạo hồ sơ thành công (STT {stt}). Cảm ơn bạn!", "success")
        # After saving, look up the new record and show it in readonly mode
        new_record = find_record_by_citizen_id(cccd)
        if new_record:
            return render_guest_form(mode="readonly", cccd=cccd, record=new_record)
        return redirect(url_for("guest_lookup"))
    except Exception as error:
        logging.exception("Guest could not save record: %s", error)
        return render_guest_form(mode="create", cccd=cccd, values=values, errors=[str(error)])


@app.errorhandler(413)
def request_too_large(error):
    return redirect(url_for("index"))


@app.get("/admin/audit-log")
@auth_required(role="admin")
def admin_audit_log():
    from quan_ly_ho_so.audit import get_audit_logs
    try:
        page = max(1, int(request.args.get("page", "1")))
    except ValueError:
        page = 1
    stt = request.args.get("stt", "").strip() or None
    per_page = 50
    entries, total = get_audit_logs(limit=per_page, offset=(page - 1) * per_page, stt=stt)
    pages = max(1, (total + per_page - 1) // per_page)
    return render_template_string(
        AUDIT_LOG_PAGE,
        entries=entries,
        total=total,
        page=page,
        pages=pages,
        stt=stt,
        csrf=csrf_token(),
        current_user=current_user(),
    )


def main():
    if not TEMPLATE.is_file():
        raise SystemExit("Mau_Ho_So_Thanh_Nien.docx phải nằm cùng thư mục với ứng dụng.")
    auth_database_connection().close()
    load_saved_workbook()
    host = os.environ.get("FLASK_RUN_HOST", "127.0.0.1")
    port = int(os.environ.get("FLASK_RUN_PORT", "8765"))
    debug = os.environ.get("FLASK_DEBUG", "0").lower() in ("1", "true", "yes")
    open_browser = os.environ.get("OPEN_BROWSER", "1").lower() in ("1", "true", "yes")
    if open_browser:
        threading.Timer(0.7, lambda: webbrowser.open(f"http://{host}:{port}")).start()
    app.run(host=host, port=port, debug=debug)



if __name__ == "__main__":
    main()

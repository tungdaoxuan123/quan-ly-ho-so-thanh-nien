#!/usr/bin/env python3
"""Local browser UI for Excel-backed youth records."""

from __future__ import annotations

import hashlib
import io
import logging
import os
import secrets
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
    send_from_directory,
    url_for,
)
from werkzeug.utils import secure_filename

from quan_ly_ho_so.config import (
    DEFAULT_OUTPUT_NAME,
    DIEN_TEMPLATE_DIR,
    FILTER_DB_COLUMNS,
    FORM_GROUPS,
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
from quan_ly_ho_so.errors import StaleWorkbookError
from quan_ly_ho_so.forms.enums import NvqsStatus
from quan_ly_ho_so.forms.fields import coerce_cell_value, field_definitions, validate_form_values
from quan_ly_ho_so.security import csrf_token, valid_csrf
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import fold
from quan_ly_ho_so.web.templates import PAGE, page_url, render_form, render_preview
from quan_ly_ho_so.word.export import build_document, combine_booklet, unique_output_path
from quan_ly_ho_so.workbook.cache import (
    all_citizen_ids,
    database_record_count,
    delete_record_file,
    find_record,
    move_record_owner,
    record_file_image,
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
app.secret_key = os.environ.get("QUAN_LY_HO_SO_SECRET") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


def valid_date(value):
    """Return `value` if it is a YYYY-MM-DD date (what <input type="date"> sends), else ""."""
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def pending_ca_report():
    pending = state.get("import_pending")
    if pending is None:
        return None
    return ca_report(pending, plan_import(pending["people"], all_citizen_ids()), preview=True)


@app.get("/")
def index():
    refresh_state()
    if state["workbook"] is None:
        return render_template_string(PAGE, loaded=False, error=state["error"], csrf=csrf_token())
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
        page_urls={number: page_url(number, params) for number in range(1, pages + 1)},
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
    )


@app.get("/api/status")
def api_status():
    refresh_state()
    return jsonify({
        "loaded": state["workbook"] is not None,
        "signature": signature_token(state["signature"]),
        "record_count": state["record_count"],
        "error": state["error"],
    })


@app.get("/refresh")
def refresh():
    refresh_state(force=True)
    if state["error"]:
        flash(state["error"], "error")
    else:
        flash("Đã làm mới dữ liệu từ tệp Excel.", "success")
    return redirect(url_for("index"))


@app.get("/choose")
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
def new_person():
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    return render_form("Thêm hồ sơ")


@app.get("/person/<stt>/edit")
def edit_person(stt):
    refresh_state()
    record = find_record(stt)
    if record is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    return render_form("Cập nhật hồ sơ", record=record)


@app.get("/person/new/preview")
def new_person_preview():
    refresh_state()
    if state["workbook"] is None:
        return redirect(url_for("index"))
    return render_preview("Thêm hồ sơ")


@app.get("/person/<stt>/preview")
def preview_person(stt):
    refresh_state()
    record = find_record(stt)
    if record is None:
        flash("Không tìm thấy hồ sơ. Hãy làm mới tệp Excel và thử lại.", "error")
        return redirect(url_for("index"))
    return render_preview("Xem trước hồ sơ", record=record)


@app.post("/person/save")
def save_person_route():
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Biểu mẫu đã hết hạn. Hãy mở lại biểu mẫu.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    definitions = field_definitions()
    values = {definition["name"]: request.form.get(definition["name"], "") for definition in definitions}
    original_stt = request.form.get("original_stt", "").strip() or None
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
        refresh_state(force=True)
        flash(f"Đã lưu hồ sơ STT {stt}. Bản sao lưu: {backup.name}", "success")
        return redirect(url_for("index"))
    except Exception as error:
        logging.exception("Could not save workbook: %s", error)
        return render_page("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=[str(error)], **status_argument)


@app.post("/generate/<stt>")
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
    flash(f"Đã thêm tài liệu {filename} vào hồ sơ liên quan.", "success")
    return redirect(url_for("edit_person", stt=stt))


@app.get("/person/<stt>/documents/<name>")
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
def remove_record_document(stt, name):
    if not valid_csrf(request.form.get("csrf_token")):
        flash("Phiên làm việc đã hết hạn. Hãy làm mới trang và thử lại.", "error")
        return redirect(url_for("index"))
    if state["workbook"] is None:
        return redirect(url_for("index"))
    citizen_id = owner_citizen_id(stt)
    if citizen_id and documents.delete_document(citizen_id, name):
        flash("Đã xóa tài liệu khỏi hồ sơ liên quan.", "success")
    else:
        flash("Không tìm thấy tài liệu cần xóa.", "error")
    return redirect(url_for("edit_person", stt=stt))


@app.get("/person/<stt>/files/<int:file_id>")
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
            set_record_attribute(state["workbook"], record["citizen_id"], field, value)
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
def bulk_update_dien():
    value = request.form.get("bulk_dien", "").strip()
    if value and value not in NvqsStatus.__members__:
        flash("Diện nghĩa vụ quân sự không hợp lệ.", "error")
        return redirect(url_for("index"))
    return apply_to_selection("type", value, "Diện", NvqsStatus.value_for(value))


@app.post("/bulk/khu-pho")
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
            ("Đã có trong Excel tổng — giữ nguyên, không cập nhật", plan["existing"]),
            ("Không có số CCCD — bỏ qua, hãy bổ sung CCCD rồi nhập lại", plan["no_citizen_id"]),
            ("Trùng CCCD ngay trong các tệp vừa chọn — chỉ lấy dòng đầu", plan["repeated"]),
        ],
    }


@app.post("/import/ca")
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
        plan = plan_import(pending["people"], all_citizen_ids())
        backup = None
        if plan["add"]:
            stts, backup = append_people(
                state["workbook"], state["sheet"], [person["values"] for person in plan["add"]],
                signature_token(state["signature"]),
            )
            for person, stt in zip(plan["add"], stts):
                person["stt"] = stt
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

    citizen_ids, missing, without_citizen_id = [], [], []
    for stt in selected_stt:
        record = find_record(stt)
        if record is None:
            missing.append(stt)
        elif not record["citizen_id"]:
            without_citizen_id.append(stt)
        else:
            citizen_ids.append(record["citizen_id"])
    if citizen_ids:
        set_records_deleted(state["workbook"], citizen_ids, deleted)
        refresh_state(force=True)
        action = "Đã xóa" if deleted else "Đã khôi phục"
        flash(f"{action} {len(citizen_ids)} hồ sơ." + (" Hồ sơ vẫn còn trong tệp Excel và có thể khôi phục." if deleted else ""), "success")
    if without_citizen_id:
        flash(f"Bỏ qua hồ sơ STT {', '.join(without_citizen_id)} vì chưa có số CCCD — trạng thái xóa được lưu theo số CCCD.", "error")
    if missing:
        flash(f"Không tìm thấy hồ sơ STT: {', '.join(missing)}. Hãy làm mới tệp Excel và thử lại.", "error")
    return redirect(url_for("index"))


@app.post("/bulk/delete")
def bulk_delete():
    return change_deleted_flag(True)


@app.post("/bulk/restore")
def bulk_restore():
    return change_deleted_flag(False)


@app.post("/export/batch")
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
    flash(f"Đã tạo tệp Excel danh sách cho {len(records)} hồ sơ: {output.name}.", "success")
    return redirect(url_for("download", filename=output.name))


@app.post("/export/dien")
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
def download(filename):
    output_dir = state.get("last_output_dir")
    if output_dir is None or not output_dir.is_dir():
        flash("Không tìm thấy thư mục chứa tệp Word đã tạo.", "error")
        return redirect(url_for("index"))
    return send_from_directory(output_dir, filename, as_attachment=True)


@app.errorhandler(413)
def request_too_large(error):
    return redirect(url_for("index"))


def main():
    if not TEMPLATE.is_file():
        raise SystemExit("Mau_Ho_So_Thanh_Nien.docx phải nằm cùng thư mục với ứng dụng.")
    load_saved_workbook()
    threading.Timer(0.7, lambda: webbrowser.open("http://127.0.0.1:8765")).start()
    app.run(host="127.0.0.1", port=8765, debug=False)


if __name__ == "__main__":
    main()

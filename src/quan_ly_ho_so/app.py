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

from flask import Flask, flash, jsonify, redirect, render_template_string, request, send_from_directory, url_for
from werkzeug.utils import secure_filename

from quan_ly_ho_so.config import (
    DEFAULT_OUTPUT_NAME,
    FILTER_DB_COLUMNS,
    FORM_GROUPS,
    MAX_RECORDS_PER_PAGE,
    SUPPORTED_SUFFIXES,
    TEMPLATE,
    UNSET_DIEN_FILTER,
    UPLOAD_DIR,
)
from quan_ly_ho_so.errors import StaleWorkbookError
from quan_ly_ho_so.forms.enums import NvqsStatus
from quan_ly_ho_so.forms.fields import coerce_cell_value, field_definitions, validate_form_values
from quan_ly_ho_so.security import csrf_token, valid_csrf
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import fold
from quan_ly_ho_so.web.templates import PAGE, page_url, render_form, render_preview
from quan_ly_ho_so.word.export import build_document, combine_booklet, unique_output_path
from quan_ly_ho_so.workbook.cache import (
    database_record_count,
    find_record,
    query_filter_options,
    query_records,
    set_record_type,
    signature_token,
)
from quan_ly_ho_so.workbook.manager import (
    load_saved_workbook,
    native_workbook_dialog,
    refresh_state,
    set_workbook,
)
from quan_ly_ho_so.workbook.writer import save_person

app = Flask(__name__)
app.secret_key = os.environ.get("QUAN_LY_HO_SO_SECRET") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")


@app.get("/")
def index():
    refresh_state()
    if state["workbook"] is None:
        return render_template_string(PAGE, loaded=False, error=state["error"], csrf=csrf_token())
    query = request.args.get("q", "").strip()
    filters = {
        "birth_year": request.args.get("birth_year", "").strip(),
        "occupation": request.args.get("occupation", "").strip(),
        "education": request.args.get("education", "").strip(),
        "ethnicity": request.args.get("ethnicity", "").strip(),
        "religion": request.args.get("religion", "").strip(),
        "address": request.args.get("address", "").strip(),
        "dien": request.args.get("dien", "").strip(),
    }
    if filters["dien"] not in NvqsStatus.__members__ and filters["dien"] != UNSET_DIEN_FILTER:
        filters["dien"] = ""
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
        dien_options=NvqsStatus.options(),
        unset_dien=UNSET_DIEN_FILTER,
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
        if nvqs_type is not None:
            set_record_type(state["workbook"], stt, nvqs_type)
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

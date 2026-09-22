#!/usr/bin/env python3
"""Local browser UI for Excel-backed youth records."""

from __future__ import annotations

import copy
import hmac
import json
import logging
import os
import secrets
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import threading
import unicodedata
import webbrowser
from datetime import datetime
from pathlib import Path

from flask import Flask, flash, jsonify, redirect, render_template_string, request, send_from_directory, session, url_for
from openpyxl import load_workbook
from werkzeug.utils import secure_filename

from tao_ho_so_word import generate_document, text


def resource_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = resource_dir()
UPLOAD_DIR = BASE_DIR / "uploads"
TEMPLATE = BASE_DIR / "Mau_Ho_So_Thanh_Nien.docx"
DEFAULT_OUTPUT_NAME = "Hồ sơ thanh niên đã tạo"
MAX_RECORDS_PER_PAGE = 50
CACHE_SEARCH_SCOPE = "identity-v1"
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
FILTER_DB_COLUMNS = {
    "birth_year": ("birth_year", "birth_year_key"),
    "occupation": ("occupation", "occupation_key"),
    "education": ("education", "education_key"),
    "ethnicity": ("ethnicity", "ethnicity_key"),
    "religion": ("religion", "religion_key"),
}

NUMERIC_HEADERS = {
    "Ngày sinh", "Tháng sinh", "Năm sinh", "Năm tốt nghiệp",
    "Ngày sinh cha", "tháng sinh cha", "năm sinh cha",
    "Ngày sinh mẹ", "Tháng sinh mẹ", "Năm sinh mẹ", "năm sinh vợ",
    "Số người con", "cha mẹ có bao nhiêu con", "mấy trai", "mấy gái",
}

LONG_TEXT_HEADERS = {
    "Nơi đăng ký khai sinh (2 cấp)", "Quê quán (2 cấp)", "Thường trú", "Nơi ở hiện nay",
    "Bản thân", "Nơi làm việc, học tập", "Trước 30-4 (cha)\nghi 3 cấp và 2 cấp",
    "sau 30-4 (cha)\nghi 2 cấp", "hiện nay (cha)\nghi 2 cấp",
    "Trước 30-4 (mẹ)\nghi 3 cấp và 2 cấp", "sau 30-4 (mẹ)\nghi 2 cấp",
    "hiện nay (mẹ)\nghi 2 cấp", "Anh chị em 1", "Anh chị em 2", "Anh chị em 3",
    "Anh chị em 4", "Anh chị em 5", "Anh chị em 6", "Anh chị em 7", "Cấp 1", "Cấp 2",
    "Cấp 3", "Hiện nay", "GHI CHÚ\n(Về Sức khỏe; gia cảnh)",
    "TẤT CẢ THÔNG TIN KHÁC\nLIÊN QUAN ĐẾN THANH NIÊN",
}

FORM_GROUPS = [
    ("Thông tin cá nhân", [
        "TÊN THƯỜNG DÙNG", "Tên khai sinh", "Ngày sinh", "Tháng sinh", "Năm sinh", "Căn cước",
        "Nơi đăng ký khai sinh (2 cấp)", "Quê quán (2 cấp)", "Dân tộc", "Tôn giáo", "Thành phần", "Bản thân",
    ]),
    ("Địa chỉ và công việc", [
        "Thường trú", "Nơi ở hiện nay", "Nghề nghiệp", "Văn hóa", "Chuyên môn", "Năm tốt nghiệp",
        "Ngành đào tạo", "Nơi làm việc, học tập",
    ]),
    ("Thông tin cha", [
        "Tên cha", "Sống / Chết", "Ngày sinh cha", "tháng sinh cha", "năm sinh cha", "Nghề nghiệp cha",
        "Trước 30-4 (cha)\nghi 3 cấp và 2 cấp", "sau 30-4 (cha)\nghi 2 cấp", "hiện nay (cha)\nghi 2 cấp",
    ]),
    ("Thông tin mẹ", [
        "Tên mẹ", "Sống / chết (mẹ)", "Ngày sinh mẹ", "Tháng sinh mẹ", "Năm sinh mẹ", "Nghề nghiệp mẹ",
        "Trước 30-4 (mẹ)\nghi 3 cấp và 2 cấp", "sau 30-4 (mẹ)\nghi 2 cấp", "hiện nay (mẹ)\nghi 2 cấp",
    ]),
    ("Vợ, con và anh chị em", [
        "Tên vợ", "năm sinh vợ", "nghề nghiệp vợ", "Số người con", "cha mẹ có bao nhiêu con", "mấy trai",
        "mấy gái", "là Con thứ", "Anh chị em 1", "Anh chị em 2", "Anh chị em 3", "Anh chị em 4",
        "Anh chị em 5", "Anh chị em 6", "Anh chị em 7",
    ]),
    ("Học tập và ghi chú", [
        "Cấp 1", "Cấp 2", "Cấp 3", "Hiện nay", "GHI CHÚ\n(Về Sức khỏe; gia cảnh)",
        "TẤT CẢ THÔNG TIN KHÁC\nLIÊN QUAN ĐẾN THANH NIÊN",
    ]),
]

app = Flask(__name__)
app.secret_key = os.environ.get("QUAN_LY_HO_SO_SECRET") or secrets.token_hex(32)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

state = {
    "workbook": None,
    "sheet": None,
    "headers": {},
    "database": None,
    "record_count": 0,
    "signature": None,
    "loaded_at": None,
    "error": None,
    "last_output_dir": None,
}
state_lock = threading.RLock()
_settings_file = None


class WorkbookError(Exception):
    """Expected workbook operation failure shown to the user."""


class WorkbookBusyError(WorkbookError):
    pass


class StaleWorkbookError(WorkbookError):
    pass


def settings_path():
    """Return a writable path for settings and the SQLite cache.

    LocalAppData is normally the right place on Windows, but it can be
    unavailable on managed computers.  In that case keep the app's local
    state beside the application rather than preventing the workbook from
    loading altogether.
    """
    global _settings_file
    if _settings_file is not None:
        return _settings_file

    if os.name == "nt":
        root = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    elif sys.platform == "darwin":
        root = Path.home() / "Library" / "Application Support"
    else:
        root = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))

    preferred = root / "Quan_Ly_Ho_So_Thanh_Nien"
    fallbacks = (BASE_DIR / ".quan_ly_ho_so_data", Path(tempfile.gettempdir()) / "Quan_Ly_Ho_So_Thanh_Nien")
    for directory in (preferred, *fallbacks):
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as error:
            logging.warning("Cannot use application-data folder %s: %s", directory, error)
            continue
        if directory != preferred:
            logging.warning("Using fallback application-data folder: %s", directory)
        _settings_file = directory / "settings.json"
        return _settings_file

    raise WorkbookError("Không thể tạo thư mục lưu dữ liệu ứng dụng.")


def default_database_path():
    return settings_path().with_name("cache.sqlite3")


def current_database_path():
    return Path(state["database"] or default_database_path())


def normalized_header(value):
    return text(value).replace("\r\n", "\n")


def display_value(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return text(value)


def fold(value):
    """Case- and accent-insensitive text used for Vietnamese search."""
    value = display_value(value).casefold().replace("đ", "d")
    return "".join(char for char in unicodedata.normalize("NFD", value) if unicodedata.category(char) != "Mn")


def file_signature(path):
    stat = path.stat()
    return (stat.st_mtime_ns, stat.st_size)


def signature_token(signature):
    if not signature:
        return ""
    return f"{signature[0]}:{signature[1]}"


def workbook_lock_path(path):
    return path.with_name(f"~${path.name}")


def workbook_appears_open(path):
    lock = workbook_lock_path(path)
    if not lock.is_file():
        return False
    try:
        age = datetime.now().timestamp() - lock.stat().st_mtime
    except OSError:
        return False
    # Excel lock files can be left behind after a crash. Treat only a current lock as active.
    return age < 2 * 60 * 60


def validate_workbook_path(path):
    path = Path(path).expanduser().resolve()
    if path.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise WorkbookError("Hãy chọn tệp Excel có đuôi .xlsx hoặc .xlsm.")
    if not path.is_file():
        raise WorkbookError("Không tìm thấy tệp Excel.")
    return path


def database_connection():
    path = current_database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS cache_metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS records (
            id INTEGER PRIMARY KEY,
            row_number INTEGER NOT NULL,
            stt TEXT NOT NULL,
            name TEXT NOT NULL,
            birth_year TEXT NOT NULL,
            citizen_id TEXT NOT NULL,
            occupation TEXT NOT NULL,
            education TEXT NOT NULL,
            ethnicity TEXT NOT NULL,
            religion TEXT NOT NULL,
            address TEXT NOT NULL,
            search_text TEXT NOT NULL,
            birth_year_key TEXT NOT NULL,
            occupation_key TEXT NOT NULL,
            education_key TEXT NOT NULL,
            ethnicity_key TEXT NOT NULL,
            religion_key TEXT NOT NULL,
            address_key TEXT NOT NULL,
            citizen_id_key TEXT NOT NULL,
            data_json TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS records_stt_idx ON records(stt);
        CREATE INDEX IF NOT EXISTS records_birth_year_idx ON records(birth_year_key);
        CREATE INDEX IF NOT EXISTS records_occupation_idx ON records(occupation_key);
        CREATE INDEX IF NOT EXISTS records_education_idx ON records(education_key);
        CREATE INDEX IF NOT EXISTS records_ethnicity_idx ON records(ethnicity_key);
        CREATE INDEX IF NOT EXISTS records_religion_idx ON records(religion_key);
        CREATE INDEX IF NOT EXISTS records_citizen_id_idx ON records(citizen_id_key);
        """
    )
    connection.commit()
    return connection


def metadata_values(connection):
    return {row["key"]: row["value"] for row in connection.execute("SELECT key, value FROM cache_metadata")}


def cached_workbook_info(path, signature):
    try:
        connection = database_connection()
        try:
            metadata = metadata_values(connection)
        finally:
            connection.close()
        if (
            metadata.get("workbook") != str(path)
            or metadata.get("signature") != signature_token(signature)
            or metadata.get("search_scope") != CACHE_SEARCH_SCOPE
        ):
            return None
        headers = json.loads(metadata["headers"])
        return metadata["sheet"], headers, int(metadata["record_count"]), datetime.fromisoformat(metadata["loaded_at"])
    except (OSError, sqlite3.Error, KeyError, TypeError, ValueError, json.JSONDecodeError) as error:
        logging.warning("Could not reuse SQLite cache: %s", error)
        return None


def sync_workbook_to_database(path):
    """Stream one workbook snapshot into SQLite without retaining all rows in RAM."""
    path = validate_workbook_path(path)
    signature_before = file_signature(path)
    keep_vba = path.suffix.lower() == ".xlsm"
    workbook = load_workbook(path, read_only=True, data_only=False, keep_vba=keep_vba)
    connection = None
    try:
        sheet = workbook.active
        first_row = next(sheet.iter_rows(min_row=1, max_row=1), ())
        headers = {}
        for column, cell in enumerate(first_row, start=1):
            header = normalized_header(cell.value)
            if header and header not in headers:
                headers[header] = column
        if "STT" not in headers or "TÊN THƯỜNG DÙNG" not in headers:
            raise WorkbookError("Hàng 1 của tệp Excel phải có cột STT và TÊN THƯỜNG DÙNG.")

        connection = database_connection()
        insert_sql = """
            INSERT INTO records (
                row_number, stt, name, birth_year, citizen_id, occupation, education, ethnicity, religion,
                address, search_text, birth_year_key, occupation_key, education_key, ethnicity_key,
                religion_key, address_key, citizen_id_key, data_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        count = 0
        batch = []
        loaded_at = datetime.now()
        with connection:
            connection.execute("DELETE FROM records")
            for row_number, values in enumerate(sheet.iter_rows(min_row=3, values_only=True), start=3):
                data = {
                    header: display_value(values[column - 1] if column <= len(values) else None)
                    for header, column in headers.items()
                }
                stt = data.get("STT", "")
                name = data.get("TÊN THƯỜNG DÙNG", "")
                if not stt or not name:
                    continue
                birth_year = data.get("Năm sinh", "")
                citizen_id = data.get("Căn cước", "")
                occupation = data.get("Nghề nghiệp", "")
                education = data.get("Văn hóa", "")
                ethnicity = data.get("Dân tộc", "")
                religion = data.get("Tôn giáo", "")
                address = f"{data.get('Thường trú', '')} {data.get('Nơi ở hiện nay', '')}".strip()
                # Quick search intentionally matches only the identity fields advertised by the UI.
                # Advanced filters cover occupation, education, address, and the remaining profile data.
                searchable = " ".join([stt, name, citizen_id])
                batch.append((
                    row_number, stt, name, birth_year, citizen_id, occupation, education, ethnicity, religion,
                    address, fold(searchable), fold(birth_year), fold(occupation), fold(education), fold(ethnicity),
                    fold(religion), fold(address), fold(citizen_id),
                    json.dumps(data, ensure_ascii=False, separators=(",", ":")),
                ))
                count += 1
                if len(batch) >= 500:
                    connection.executemany(insert_sql, batch)
                    batch.clear()
            if batch:
                connection.executemany(insert_sql, batch)
            signature_after = file_signature(path)
            if signature_after != signature_before:
                raise StaleWorkbookError("Tệp Excel thay đổi trong lúc đồng bộ. Hãy thử làm mới lại.")
            metadata = {
                "workbook": str(path),
                "signature": signature_token(signature_after),
                "sheet": sheet.title,
                "headers": json.dumps(headers, ensure_ascii=False, separators=(",", ":")),
                "record_count": str(count),
                "loaded_at": loaded_at.isoformat(),
                "search_scope": CACHE_SEARCH_SCOPE,
            }
            connection.executemany(
                "INSERT OR REPLACE INTO cache_metadata(key, value) VALUES (?, ?)", metadata.items()
            )
        return sheet.title, headers, count, signature_after, loaded_at
    finally:
        workbook.close()
        if connection is not None:
            connection.close()


def database_record_count():
    connection = database_connection()
    try:
        return int(connection.execute("SELECT COUNT(*) FROM records").fetchone()[0])
    finally:
        connection.close()


def row_to_record(row):
    return {
        "row": row["row_number"],
        "stt": row["stt"],
        "name": row["name"],
        "birth_year": row["birth_year"],
        "citizen_id": row["citizen_id"],
        "occupation": row["occupation"],
        "data": json.loads(row["data_json"]),
    }


def find_record(stt):
    connection = database_connection()
    try:
        row = connection.execute(
            "SELECT * FROM records WHERE stt = ? ORDER BY row_number LIMIT 1", (display_value(stt),)
        ).fetchone()
        return row_to_record(row) if row is not None else None
    finally:
        connection.close()


def save_settings(path):
    target = settings_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(json.dumps({"workbook": str(path)}, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as error:
        logging.warning("Could not save settings: %s", error)


def load_saved_workbook():
    target = settings_path()
    try:
        saved = json.loads(target.read_text(encoding="utf-8")).get("workbook")
    except FileNotFoundError:
        return False
    except (OSError, ValueError) as error:
        logging.warning("Could not read settings: %s", error)
        return False
    if not saved:
        return False
    try:
        set_workbook(Path(saved), persist=False)
    except Exception as error:
        logging.exception("Could not load saved workbook: %s", error)
        state["error"] = str(error)
        return False
    return True


def set_workbook(path, persist=True):
    path = validate_workbook_path(path)
    signature = file_signature(path)
    cached = cached_workbook_info(path, signature)
    if cached is None:
        sheet, headers, record_count, signature, loaded_at = sync_workbook_to_database(path)
    else:
        sheet, headers, record_count, loaded_at = cached
    with state_lock:
        state.update(
            workbook=path,
            sheet=sheet,
            headers=headers,
            record_count=record_count,
            signature=signature,
            loaded_at=loaded_at,
            error=None,
            last_output_dir=None,
        )
    if persist:
        save_settings(path)


def refresh_state(force=False):
    with state_lock:
        path = state["workbook"]
        previous = state["signature"]
    if path is None:
        return False
    try:
        current = file_signature(path)
    except OSError as error:
        state["error"] = f"Không thể đọc tệp Excel: {error}"
        return False
    if not force and previous == current:
        return False
    try:
        sheet, headers, record_count, current, loaded_at = sync_workbook_to_database(path)
    except Exception as error:
        logging.exception("Could not refresh workbook: %s", error)
        state["error"] = f"Không thể làm mới tệp Excel: {error}"
        return False
    with state_lock:
        state.update(
            sheet=sheet,
            headers=headers,
            record_count=record_count,
            signature=current,
            loaded_at=loaded_at,
            error=None,
        )
    return True


def field_definitions():
    headers = state["headers"]
    definitions = []
    seen = set()
    for section, section_headers in FORM_GROUPS:
        for header in section_headers:
            column = headers.get(header)
            if not column or header in seen:
                continue
            seen.add(header)
            definitions.append({
                "section": section,
                "header": header,
                "label": " ".join(header.split()),
                "column": column,
                "name": f"column_{column}",
                "input_type": "textarea" if header in LONG_TEXT_HEADERS else ("number" if header in NUMERIC_HEADERS else "text"),
                "required": header == "TÊN THƯỜNG DÙNG",
            })
    return definitions


def form_values(record=None):
    if record is None:
        return {definition["name"]: "" for definition in field_definitions()}
    return {
        definition["name"]: display_value(record["data"].get(definition["header"]))
        for definition in field_definitions()
    }


def coerce_cell_value(header, raw):
    raw = raw.strip()
    if not raw:
        return None
    if header not in NUMERIC_HEADERS:
        return raw
    try:
        number = float(raw)
    except ValueError:
        return raw
    return int(number) if number.is_integer() else number


def validate_form_values(values):
    errors = []
    name_column = state["headers"].get("TÊN THƯỜNG DÙNG")
    name = values.get(name_column, "").strip() if name_column else ""
    if not name:
        errors.append("TÊN THƯỜNG DÙNG là thông tin bắt buộc.")
    for header in NUMERIC_HEADERS:
        column = state["headers"].get(header)
        raw = values.get(column, "").strip() if column else ""
        if raw:
            try:
                number = float(raw)
                if not number.is_integer() or number < 0:
                    raise ValueError
            except ValueError:
                errors.append(f"{header} phải là số nguyên không âm.")
    citizen_column = state["headers"].get("Căn cước")
    citizen_id = values.get(citizen_column, "").strip() if citizen_column else ""
    if citizen_id:
        original_stt = values.get("__original_stt", "")
        connection = database_connection()
        try:
            duplicate = connection.execute(
                "SELECT 1 FROM records WHERE citizen_id_key = ? AND stt <> ? LIMIT 1",
                (fold(citizen_id), original_stt),
            ).fetchone()
        finally:
            connection.close()
        if duplicate is not None:
            errors.append("Số CCCD đã được dùng cho một hồ sơ khác.")
    return errors


def copy_row_style(sheet, source_row, target_row):
    if not source_row:
        return
    for column in range(1, sheet.max_column + 1):
        source = sheet.cell(source_row, column)
        target = sheet.cell(target_row, column)
        if source.has_style:
            target._style = copy.copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy.copy(source.alignment)
        if source.protection:
            target.protection = copy.copy(source.protection)
    if source_row in sheet.row_dimensions:
        sheet.row_dimensions[target_row].height = sheet.row_dimensions[source_row].height


def unique_backup_path(path):
    base_name = f"{path.stem}.backup-{datetime.now():%Y%m%d-%H%M%S}"
    candidate = path.with_name(base_name + path.suffix)
    number = 2
    while candidate.exists():
        candidate = path.with_name(f"{base_name}-{number}{path.suffix}")
        number += 1
    return candidate


def save_person(path, sheet_name, updates, expected_signature, original_stt=None):
    """Atomically update one record and preserve a recoverable backup."""
    path = validate_workbook_path(path)
    if path.suffix.lower() != ".xlsx":
        raise WorkbookError("Ứng dụng chỉ lưu biểu mẫu vào tệp .xlsx. Hãy mở tệp .xlsm bằng Excel để chỉnh sửa.")
    with state_lock:
        current_signature = file_signature(path)
    if signature_token(current_signature) != expected_signature:
        raise StaleWorkbookError("Tệp Excel đã thay đổi bên ngoài ứng dụng. Hãy làm mới và mở lại hồ sơ.")
    if workbook_appears_open(path):
        raise WorkbookBusyError("Tệp Excel đang được mở. Hãy lưu, đóng Excel rồi thử lại.")

    workbook = load_workbook(path, read_only=False, data_only=False)
    temp_name = None
    try:
        if sheet_name not in workbook.sheetnames:
            raise WorkbookError(f"Không tìm thấy trang tính: {sheet_name}")
        sheet = workbook[sheet_name]
        headers = {normalized_header(cell.value): cell.column for cell in sheet[1] if normalized_header(cell.value)}
        stt_column = headers.get("STT")
        name_column = headers.get("TÊN THƯỜNG DÙNG")
        if not stt_column or not name_column:
            raise WorkbookError("Hàng 1 của tệp Excel phải có cột STT và TÊN THƯỜNG DÙNG.")

        target_row = None
        if original_stt is not None:
            for row_number in range(3, sheet.max_row + 1):
                if display_value(sheet.cell(row_number, stt_column).value) == display_value(original_stt):
                    target_row = row_number
                    break
            if target_row is None:
                raise StaleWorkbookError("Hồ sơ không còn trong tệp Excel. Hãy làm mới và thử lại.")
        else:
            target_row = 3
            highest_stt = 0
            for row_number in range(3, sheet.max_row + 1):
                existing_stt = display_value(sheet.cell(row_number, stt_column).value)
                existing_name = display_value(sheet.cell(row_number, name_column).value)
                if not existing_stt or not existing_name:
                    continue
                target_row = row_number + 1
                try:
                    highest_stt = max(highest_stt, int(float(existing_stt)))
                except (TypeError, ValueError):
                    continue
            copy_row_style(sheet, target_row - 1, target_row)
            updates[stt_column] = highest_stt + 1

        for column, value in updates.items():
            sheet.cell(target_row, column).value = value

        if not display_value(sheet.cell(target_row, name_column).value):
            raise WorkbookError("TÊN THƯỜNG DÙNG là thông tin bắt buộc.")
        saved_stt = display_value(sheet.cell(target_row, stt_column).value)

        backup_name = unique_backup_path(path)
        with tempfile.NamedTemporaryFile(prefix=f".{path.stem}-", suffix=path.suffix, dir=path.parent, delete=False) as temporary:
            temp_name = temporary.name
        workbook.save(temp_name)
        workbook.close()
        check = load_workbook(temp_name, read_only=True, data_only=False)
        check.close()
        shutil.copy2(path, backup_name)
        os.replace(temp_name, path)
        temp_name = None
        return target_row, saved_stt, backup_name
    except PermissionError as error:
        raise WorkbookBusyError("Excel đang mở tệp này hoặc thư mục không cho phép ghi tệp.") from error
    finally:
        try:
            workbook.close()
        except OSError as error:
            logging.warning("Could not close workbook handle: %s", error)
        if temp_name:
            try:
                Path(temp_name).unlink(missing_ok=True)
            except OSError:
                logging.warning("Could not remove temporary workbook %s", temp_name)


def escaped_like(value):
    return value.replace("\\", "\\\\").replace("%", "\\%").replace("_", "\\_")


def query_records(query, filters, limit, offset):
    conditions = []
    parameters = []
    query_key = fold(query)
    if query_key:
        conditions.append("search_text LIKE ? ESCAPE '\\'")
        parameters.append(f"%{escaped_like(query_key)}%")
    for key, (_, key_column) in FILTER_DB_COLUMNS.items():
        value = fold(filters.get(key, ""))
        if value:
            conditions.append(f"{key_column} = ?")
            parameters.append(value)
    address = fold(filters.get("address", ""))
    if address:
        conditions.append("address_key LIKE ? ESCAPE '\\'")
        parameters.append(f"%{escaped_like(address)}%")
    where = f" WHERE {' AND '.join(conditions)}" if conditions else ""

    connection = database_connection()
    try:
        count = int(connection.execute(f"SELECT COUNT(*) FROM records{where}", parameters).fetchone()[0])
        rows = connection.execute(
            f"SELECT * FROM records{where} ORDER BY row_number LIMIT ? OFFSET ?",
            [*parameters, limit, offset],
        ).fetchall()
        return [row_to_record(row) for row in rows], count
    finally:
        connection.close()


def query_filter_options(key):
    if key not in FILTER_DB_COLUMNS:
        raise ValueError(f"Unknown filter: {key}")
    value_column, key_column = FILTER_DB_COLUMNS[key]
    connection = database_connection()
    try:
        rows = connection.execute(
            f"""
            SELECT {value_column} AS value, {key_column} AS normalized
            FROM records
            WHERE {key_column} <> ''
            ORDER BY row_number
            """
        )
        values = {}
        for row in rows:
            values.setdefault(row["normalized"], row["value"])
        return sorted(values.values(), key=fold)
    finally:
        connection.close()


def csrf_token():
    token = session.get("csrf_token")
    if not token:
        token = secrets.token_urlsafe(32)
        session["csrf_token"] = token
    return token


def valid_csrf(value):
    return bool(value) and hmac.compare_digest(value, session.get("csrf_token", ""))


def native_workbook_dialog():
    root = None
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        selected = filedialog.askopenfilename(
            title="Chọn tệp Excel",
            filetypes=[("Tệp Excel", "*.xlsx *.xlsm"), ("Tất cả tệp", "*.*")],
        )
        return Path(selected) if selected else None
    except Exception as error:
        logging.exception("Native file dialog failed: %s", error)
        return None
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception as error:
                logging.warning("Could not close native file dialog: %s", error)


PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 1240px; margin: 0 auto; padding: 28px; } h1 { margin: 0; color: #163d68; } h2 { color: #163d68; margin: 0 0 10px; font-size: 19px; }
    p { line-height: 1.45; } .muted, .status { color: #637089; } .status { font-size: 14px; }
    .card, form.panel { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: 0 2px 10px #1720330c; }
    .app-header, .summary, .actions, .pagination { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; justify-content: space-between; }
    .actions, .pagination { justify-content: flex-start; } .quick-search { margin-top: 18px; }
    .search-row, .filter-grid { display: grid; gap: 12px; } .search-row { grid-template-columns: minmax(260px, 1fr) auto auto; align-items: end; } .filter-grid { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); margin-top: 14px; }
    label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; } input, select, textarea, button { box-sizing: border-box; font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 10px 11px; }
    input, select, textarea { background: #fff; width: 100%; } button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-block; padding: 10px 13px; border-radius: 7px; white-space: nowrap; }
    button.secondary, .button.secondary { color: #1e5b91; background: #e9f1f8; } .notice { padding: 12px 14px; border-radius: 8px; background: #e9f4eb; color: #265b31; } .notice.error { background: #fdecec; color: #8f2f2f; }
    details.advanced { border-top: 1px solid #e4e9f0; margin-top: 18px; padding-top: 14px; } summary { cursor: pointer; color: #1e5b91; font-weight: 700; }
    .table-wrap { overflow-x: auto; margin-top: 14px; } table { width: 100%; min-width: 800px; border-collapse: collapse; background: #fff; } th, td { text-align: left; padding: 11px; border-bottom: 1px solid #e4e9f0; vertical-align: top; }
    th { color: #344a66; background: #eef3f8; } td.actions-cell { white-space: nowrap; } .inline-form { display: inline; }
    .pagination { margin-top: 16px; } @media (max-width: 700px) { body { padding: 14px; } .search-row { grid-template-columns: 1fr; } .search-row button, .search-row .button { text-align: center; } }
  </style>
</head>
<body>
  <header class="app-header"><div><h1>Quản lý hồ sơ thanh niên</h1><p class="muted">Tra cứu, cập nhật hồ sơ và tạo tệp Word từ Excel</p></div></header>
  {% with messages = get_flashed_messages(with_categories=true) %}{% for category, message in messages %}<p class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</p>{% endfor %}{% endwith %}
  {% if error %}<p class="notice error">{{ error }}</p>{% endif %}
  {% if not loaded %}
    <section class="card"><h2>Chọn tệp Excel</h2><p>Chọn tệp Excel gốc để ứng dụng đọc trực tiếp. Thay đổi đã lưu trong Excel sẽ xuất hiện sau khi làm mới.</p><div class="actions"><a class="button" href="{{ url_for('choose_workbook') }}">Chọn tệp Excel</a></div>
      <details class="advanced"><summary>Hoặc tải lên bản sao tệp Excel</summary><form class="panel" action="{{ url_for('load_uploaded_workbook') }}" method="post" enctype="multipart/form-data"><label>Tệp Excel<input type="file" name="workbook" accept=".xlsx,.xlsm" required></label><p><button type="submit">Dùng bản sao đã tải lên</button></p></form></details>
    </section>
  {% else %}
    <section class="card summary"><div><strong>{{ workbook_name }}</strong><div class="status">Trang tính: {{ sheet_name }} · {{ record_count }} hồ sơ · Đồng bộ SQLite lúc {{ loaded_at }}</div></div><div class="actions"><a class="button secondary" href="{{ url_for('refresh') }}">Làm mới</a><a class="button secondary" href="{{ url_for('open_excel') }}">Mở bằng Excel</a><a class="button" href="{{ url_for('new_person') }}">Thêm hồ sơ</a><a class="button secondary" href="{{ url_for('choose_workbook') }}">Đổi tệp Excel</a></div></section>
    {% if read_only %}<p class="notice">Tệp .xlsm chỉ có thể xem và tạo Word trong ứng dụng. Hãy dùng Excel để lưu thay đổi.</p>{% endif %}
    <form class="card quick-search" action="{{ url_for('index') }}" method="get"><h2>Tìm kiếm hồ sơ</h2><div class="search-row"><label>Họ tên, STT hoặc CCCD<input name="q" value="{{ query }}" placeholder="Ví dụ: Nguyễn, 12 hoặc số CCCD" autofocus></label><button type="submit">Tìm kiếm</button><a class="button secondary" href="{{ url_for('index') }}">Xóa tìm kiếm</a></div>
      <details class="advanced" {% if filters_active %}open{% endif %}><summary>Bộ lọc nâng cao{% if filters_active %} đang được áp dụng{% endif %}</summary><div class="filter-grid"><label>Năm sinh<select name="birth_year"><option value="">Tất cả</option>{% for value in options.birth_year %}<option value="{{ value }}" {% if filters.birth_year == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Nghề nghiệp<select name="occupation"><option value="">Tất cả</option>{% for value in options.occupation %}<option value="{{ value }}" {% if filters.occupation == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Trình độ văn hóa<select name="education"><option value="">Tất cả</option>{% for value in options.education %}<option value="{{ value }}" {% if filters.education == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Dân tộc<select name="ethnicity"><option value="">Tất cả</option>{% for value in options.ethnicity %}<option value="{{ value }}" {% if filters.ethnicity == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Tôn giáo<select name="religion"><option value="">Tất cả</option>{% for value in options.religion %}<option value="{{ value }}" {% if filters.religion == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Địa chỉ<input name="address" value="{{ filters.address }}" placeholder="Thường trú hoặc nơi ở hiện nay"></label></div><p class="actions"><button type="submit">Áp dụng bộ lọc</button><a class="button secondary" href="{{ url_for('index') }}">Xóa bộ lọc</a></p></details>
    </form>
    <div class="status">Hiển thị {{ shown_start }}–{{ shown_end }} trong tổng số {{ filtered_count }} hồ sơ phù hợp</div>
    <div class="table-wrap"><table><thead><tr><th>STT</th><th>Họ và tên</th><th>Năm sinh</th><th>CCCD</th><th>Nghề nghiệp</th><th>Thao tác</th></tr></thead><tbody>{% for record in records %}<tr><td>{{ record.stt }}</td><td>{{ record.name }}</td><td>{{ record.birth_year }}</td><td>{{ record.citizen_id }}</td><td>{{ record.occupation }}</td><td class="actions-cell"><a class="button secondary" href="{{ url_for('edit_person', stt=record.stt) }}">Sửa</a> <form class="inline-form" action="{{ url_for('generate', stt=record.stt) }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"><button type="submit">Tạo Word</button></form></td></tr>{% else %}<tr><td colspan="6">Không tìm thấy hồ sơ phù hợp.</td></tr>{% endfor %}</tbody></table></div>
    {% if pages > 1 %}<nav class="pagination" aria-label="Phân trang">{% for page_number in range(1, pages + 1) %}<a class="button {{ 'secondary' if page_number != page else '' }}" href="{{ page_urls[page_number] }}">{{ page_number }}</a>{% endfor %}</nav>{% endif %}
    <script>const currentSignature = {{ signature|tojson }}; setInterval(async () => { try { const response = await fetch({{ url_for('api_status')|tojson }}, {cache: 'no-store'}); const status = await response.json(); if (status.signature && currentSignature && status.signature !== currentSignature) window.location.reload(); } catch (error) { console.warn('Không thể kiểm tra thay đổi tệp Excel', error); } }, 5000);</script>
  {% endif %}
</body>
</html>
"""


FORM_PAGE = r"""
<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{{ title }} · Quản lý hồ sơ thanh niên</title>
<style>
  :root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; } body { max-width: 1080px; margin: 0 auto; padding: 28px; } h1 { color: #163d68; margin-bottom: 4px; }
  form { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; } .section-card { border: 1px solid #dbe3ee; border-radius: 9px; margin: 14px 0; overflow: hidden; } summary { cursor: pointer; display: flex; justify-content: space-between; gap: 12px; padding: 14px 16px; color: #163d68; font-weight: 700; background: #f7f9fc; } .section-card[open] summary { border-bottom: 1px solid #dbe3ee; background: #eef4fa; } .field-count { color: #637089; font-size: 14px; font-weight: 500; } .section-body { padding: 16px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 14px; } label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; } input, textarea, button, .button { font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 10px 11px; box-sizing: border-box; } input, textarea { width: 100%; } textarea { min-height: 84px; resize: vertical; } .wide { grid-column: 1 / -1; } .actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 22px; } button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; } .button.secondary { color: #1e5b91; background: #e9f1f8; } .notice { padding: 12px 14px; border-radius: 8px; background: #fdecec; color: #8f2f2f; } .muted { color: #637089; } .required { color: #a52a2a; } @media (max-width: 700px) { body { padding: 14px; } .grid { grid-template-columns: 1fr; } }
</style></head><body>
<h1>{{ title }}</h1><p class="muted">Tệp Excel: {{ workbook_name }} · Trang tính: {{ sheet_name }}</p><p class="muted"><span class="required">*</span> Trường bắt buộc</p>
{% if message %}<p class="notice">{{ message }}</p>{% endif %}{% if errors %}<div class="notice"><strong>Chưa thể lưu hồ sơ:</strong><ul>{% for error in errors %}<li>{{ error }}</li>{% endfor %}</ul></div>{% endif %}
<form action="{{ url_for('save_person_route') }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"><input type="hidden" name="original_stt" value="{{ original_stt }}"><input type="hidden" name="signature" value="{{ signature }}">
  {% for section in sections %}<details class="section-card" {% if section.name in open_sections %}open{% endif %}><summary><span>{{ section.name }}</span><span class="field-count">{{ section.fields|length }} trường</span></summary><div class="section-body"><div class="grid">{% for field in section.fields %}<label class="{{ 'wide' if field.input_type == 'textarea' else '' }}">{{ field.label }}{% if field.required %} <span class="required">*</span>{% endif %}{% if field.input_type == 'textarea' %}<textarea name="{{ field.name }}" {% if read_only %}readonly{% endif %}>{{ values.get(field.name, '') }}</textarea>{% else %}<input type="{{ field.input_type }}" name="{{ field.name }}" value="{{ values.get(field.name, '') }}" {% if field.required %}required{% endif %} {% if read_only %}readonly{% endif %}>{% endif %}</label>{% endfor %}</div></div></details>{% endfor %}
  <div class="actions">{% if not read_only %}<button type="submit">Lưu vào Excel</button>{% endif %}<a class="button secondary" href="{{ url_for('index') }}">Quay lại danh sách</a></div>
</form></body></html>
"""


def render_form(title, record=None, values=None, errors=None, message=None):
    definitions = field_definitions()
    errors = errors or []
    open_sections = {"Thông tin cá nhân"}
    for error in errors:
        folded_error = fold(error)
        for definition in definitions:
            if fold(definition["header"]) in folded_error:
                open_sections.add(definition["section"])
    grouped = []
    for section_name, _ in FORM_GROUPS:
        grouped.append({"name": section_name, "fields": [field for field in definitions if field["section"] == section_name]})
    return render_template_string(
        FORM_PAGE,
        title=title,
        workbook_name=state["workbook"].name if state["workbook"] else "",
        sheet_name=state["sheet"] or "",
        sections=grouped,
        values=values if values is not None else form_values(record),
        errors=errors,
        open_sections=open_sections,
        message=message,
        original_stt=record["stt"] if record else "",
        signature=signature_token(state["signature"]),
        csrf=csrf_token(),
        read_only=not state["workbook"] or state["workbook"].suffix.lower() == ".xlsm",
    )


def page_url(page_number, params):
    query = dict(params)
    query["page"] = page_number
    return url_for("index", **query)


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
    if validation_errors:
        return render_form("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=validation_errors)

    headers_by_column = {definition["column"]: definition["header"] for definition in definitions}
    updates = {column: coerce_cell_value(headers_by_column[column], raw) for column, raw in values_by_column.items()}
    try:
        row, stt, backup = save_person(
            state["workbook"], state["sheet"], updates, request.form.get("signature", ""), original_stt=original_stt,
        )
        refresh_state(force=True)
        flash(f"Đã lưu hồ sơ STT {stt}. Bản sao lưu: {backup.name}", "success")
        return redirect(url_for("index"))
    except Exception as error:
        logging.exception("Could not save workbook: %s", error)
        return render_form("Cập nhật hồ sơ" if original_stt else "Thêm hồ sơ", record=record, values=values, errors=[str(error)])


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
        output = generate_document(state["workbook"], state["sheet"], record["row"], TEMPLATE, output_dir)
    except Exception as error:
        logging.exception("Could not generate document: %s", error)
        flash(f"Không thể tạo tệp Word: {error}", "error")
        return redirect(url_for("index"))
    state["last_output_dir"] = output_dir
    flash(f"Đã tạo tệp Word: {output.name}.", "success")
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


if __name__ == "__main__":
    if not TEMPLATE.is_file():
        raise SystemExit("Mau_Ho_So_Thanh_Nien.docx phải nằm cùng thư mục với ứng dụng.")
    load_saved_workbook()
    threading.Timer(0.7, lambda: webbrowser.open("http://127.0.0.1:8765")).start()
    app.run(host="127.0.0.1", port=8765, debug=False)

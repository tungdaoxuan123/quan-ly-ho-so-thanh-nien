"""Form field metadata derived from the loaded workbook's headers, and validation."""

from quan_ly_ho_so.config import FORM_GROUPS, LONG_TEXT_HEADERS, NUMERIC_HEADERS
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import display_value, fold
from quan_ly_ho_so.workbook.cache import database_connection


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

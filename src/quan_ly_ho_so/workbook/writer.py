"""Atomic single-record writes back into the source .xlsx workbook."""

import copy
import logging
import os
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from quan_ly_ho_so.errors import StaleWorkbookError, WorkbookBusyError, WorkbookError
from quan_ly_ho_so.state import state_lock
from quan_ly_ho_so.utils.text import display_value, normalized_header
from quan_ly_ho_so.workbook.cache import (
    file_signature,
    signature_token,
    validate_workbook_path,
    workbook_appears_open,
)


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


def checked_writable_path(path, expected_signature):
    """Validate `path` for a write: an .xlsx, unchanged since `expected_signature`, and not open in Excel."""
    path = validate_workbook_path(path)
    if path.suffix.lower() != ".xlsx":
        raise WorkbookError("Ứng dụng chỉ lưu biểu mẫu vào tệp .xlsx. Hãy mở tệp .xlsm bằng Excel để chỉnh sửa.")
    with state_lock:
        current_signature = file_signature(path)
    if signature_token(current_signature) != expected_signature:
        raise StaleWorkbookError("Tệp Excel đã thay đổi bên ngoài ứng dụng. Hãy làm mới và mở lại hồ sơ.")
    if workbook_appears_open(path):
        raise WorkbookBusyError("Tệp Excel đang được mở. Hãy lưu, đóng Excel rồi thử lại.")
    return path


def replace_workbook(path, workbook):
    """Save `workbook` over `path` atomically, keeping the old file as a backup, and return the backup's path."""
    backup_name = unique_backup_path(path)
    with tempfile.NamedTemporaryFile(prefix=f".{path.stem}-", suffix=path.suffix, dir=path.parent, delete=False) as temporary:
        temp_name = temporary.name
    try:
        workbook.save(temp_name)
        workbook.close()
        check = load_workbook(temp_name, read_only=True, data_only=False)
        check.close()
        shutil.copy2(path, backup_name)
        os.replace(temp_name, path)
    except BaseException:
        try:
            Path(temp_name).unlink(missing_ok=True)
        except OSError:
            logging.warning("Could not remove temporary workbook %s", temp_name)
        raise
    return backup_name


def last_person_row(sheet, stt_column, name_column):
    """Return (row of the last person, highest numeric STT); row is 2 when the sheet has no people yet."""
    last_row, highest_stt = 2, 0
    for row_number in range(3, sheet.max_row + 1):
        existing_stt = display_value(sheet.cell(row_number, stt_column).value)
        existing_name = display_value(sheet.cell(row_number, name_column).value)
        if not existing_stt or not existing_name:
            continue
        last_row = row_number
        try:
            highest_stt = max(highest_stt, int(float(existing_stt)))
        except (TypeError, ValueError):
            continue
    return last_row, highest_stt


def save_person(path, sheet_name, updates, expected_signature, original_stt=None):
    """Atomically update one record and preserve a recoverable backup."""
    path = checked_writable_path(path, expected_signature)
    workbook = load_workbook(path, read_only=False, data_only=False)
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
            last_row, highest_stt = last_person_row(sheet, stt_column, name_column)
            target_row = last_row + 1
            copy_row_style(sheet, last_row, target_row)
            updates[stt_column] = highest_stt + 1

        for column, value in updates.items():
            sheet.cell(target_row, column).value = value

        if not display_value(sheet.cell(target_row, name_column).value):
            raise WorkbookError("TÊN THƯỜNG DÙNG là thông tin bắt buộc.")
        saved_stt = display_value(sheet.cell(target_row, stt_column).value)

        backup_name = replace_workbook(path, workbook)
        return target_row, saved_stt, backup_name
    except PermissionError as error:
        raise WorkbookBusyError("Excel đang mở tệp này hoặc thư mục không cho phép ghi tệp.") from error
    finally:
        try:
            workbook.close()
        except OSError as error:
            logging.warning("Could not close workbook handle: %s", error)


def append_people(path, sheet_name, people, expected_signature):
    """Append new records in one atomic write and return ([their STTs], backup path).

    `people` is a list of {normalized header: value}; headers the sheet lacks are ignored.
    """
    path = checked_writable_path(path, expected_signature)
    workbook = load_workbook(path, read_only=False, data_only=False)
    try:
        if sheet_name not in workbook.sheetnames:
            raise WorkbookError(f"Không tìm thấy trang tính: {sheet_name}")
        sheet = workbook[sheet_name]
        headers = {normalized_header(cell.value): cell.column for cell in sheet[1] if normalized_header(cell.value)}
        stt_column = headers.get("STT")
        name_column = headers.get("TÊN THƯỜNG DÙNG")
        if not stt_column or not name_column:
            raise WorkbookError("Hàng 1 của tệp Excel phải có cột STT và TÊN THƯỜNG DÙNG.")

        last_row, highest_stt = last_person_row(sheet, stt_column, name_column)
        assigned = []
        for offset, person in enumerate(people, start=1):
            row = last_row + offset
            copy_row_style(sheet, last_row, row)
            sheet.cell(row, stt_column).value = highest_stt + offset
            for header, value in person.items():
                if header in headers and header != "STT":
                    sheet.cell(row, headers[header]).value = value
            if not display_value(sheet.cell(row, name_column).value):
                raise WorkbookError("TÊN THƯỜNG DÙNG là thông tin bắt buộc.")
            assigned.append(str(highest_stt + offset))
        return assigned, replace_workbook(path, workbook)
    except PermissionError as error:
        raise WorkbookBusyError("Excel đang mở tệp này hoặc thư mục không cho phép ghi tệp.") from error
    finally:
        try:
            workbook.close()
        except OSError as error:
            logging.warning("Could not close workbook handle: %s", error)

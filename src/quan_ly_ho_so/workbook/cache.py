"""Workbook path validation and the SQLite read/search cache."""

import json
import logging
import sqlite3
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from quan_ly_ho_so.config import CACHE_SEARCH_SCOPE, FILTER_DB_COLUMNS, SUPPORTED_SUFFIXES, UNSET_DIEN_FILTER
from quan_ly_ho_so.errors import StaleWorkbookError, WorkbookError
from quan_ly_ho_so.forms.enums import NvqsStatus
from quan_ly_ho_so.state import current_database_path, state
from quan_ly_ho_so.utils.text import display_value, fold, normalized_header


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


def _stt_keyed_tables(connection):
    """Names of the app-owned tables still keyed by the old, shifting STT."""
    stale = []
    for table in ("record_types", "record_files"):
        columns = {row["name"] for row in connection.execute(f"PRAGMA table_info({table})")}
        if columns and "citizen_id" not in columns:
            stale.append(table)
    return stale


def _rekey_to_citizen_id(connection, tables):
    """Re-point rows from STT onto the owner's CCCD, using whatever the cache can still match.

    A row whose STT no longer resolves cannot be attributed to anyone, so it is dropped rather
    than left pointing at whoever happens to hold that number now.
    """
    for table in tables:
        columns = [row["name"] for row in connection.execute(f"PRAGMA table_info({table}_by_stt)")]
        carried = [name for name in columns if name not in ("stt", "id")]
        selected = ", ".join(f"old.{name}" for name in carried)
        moved = connection.execute(
            f"""
            INSERT INTO {table} (citizen_id, {", ".join(carried)})
            SELECT records.citizen_id, {selected}
            FROM {table}_by_stt AS old
            JOIN records ON records.stt = old.stt
            WHERE records.citizen_id <> ''
            """
        ).rowcount
        dropped = connection.execute(f"SELECT COUNT(*) FROM {table}_by_stt").fetchone()[0] - moved
        connection.execute(f"DROP TABLE {table}_by_stt")
        logging.info("Re-keyed %s to CCCD: %d row(s) moved, %d dropped", table, moved, dropped)


def _stop_reusing_file_ids(connection):
    """Rebuild record_files with AUTOINCREMENT if it predates that fix, keeping existing ids."""
    created = connection.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'record_files'"
    ).fetchone()
    if created is None or "AUTOINCREMENT" in created["sql"]:
        return
    columns = [row["name"] for row in connection.execute("PRAGMA table_info(record_files)")]
    names = ", ".join(columns)
    connection.executescript(
        f"""
        ALTER TABLE record_files RENAME TO record_files_reused_ids;
        CREATE TABLE record_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workbook TEXT NOT NULL,
            citizen_id TEXT NOT NULL,
            source_name TEXT NOT NULL,
            page INTEGER NOT NULL,
            image BLOB NOT NULL,
            thumbnail BLOB NOT NULL,
            created_at TEXT NOT NULL
        );
        INSERT INTO record_files ({names}) SELECT {names} FROM record_files_reused_ids;
        DROP TABLE record_files_reused_ids;
        CREATE INDEX IF NOT EXISTS record_files_owner_idx ON record_files(workbook, citizen_id);
        """
    )
    logging.info("Rebuilt record_files so deleted image ids are never handed out again")


def database_connection():
    path = current_database_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(path, timeout=10)
    connection.row_factory = sqlite3.Row
    stale_tables = _stt_keyed_tables(connection)
    for table in stale_tables:
        connection.execute(f"ALTER TABLE {table} RENAME TO {table}_by_stt")
    connection.executescript(
        """
        CREATE TABLE IF NOT EXISTS cache_metadata (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        -- Manager-chosen NVQS status. Unlike `records`, this is NOT rebuilt from the workbook,
        -- so it has to survive the DELETE/re-insert that every sync performs. Keyed by CCCD
        -- rather than STT, because STT shifts whenever rows are inserted or renumbered in Excel;
        -- the workbook is part of the key since the same person may appear in separate files.
        CREATE TABLE IF NOT EXISTS record_types (
            workbook TEXT NOT NULL,
            citizen_id TEXT NOT NULL,
            type TEXT NOT NULL,
            PRIMARY KEY (workbook, citizen_id)
        );
        -- Related paperwork, already rendered to page images. Also app-owned data that the
        -- workbook knows nothing about, so it must outlive the rebuild of `records` below.
        -- AUTOINCREMENT matters here: a plain INTEGER PRIMARY KEY hands a deleted row's id to the
        -- next upload, so an image URL would quietly start serving a different document.
        CREATE TABLE IF NOT EXISTS record_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            workbook TEXT NOT NULL,
            citizen_id TEXT NOT NULL,
            source_name TEXT NOT NULL,
            page INTEGER NOT NULL,
            image BLOB NOT NULL,
            thumbnail BLOB NOT NULL,
            created_at TEXT NOT NULL
        );
        CREATE INDEX IF NOT EXISTS record_files_owner_idx ON record_files(workbook, citizen_id);
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
            data_json TEXT NOT NULL,
            type TEXT
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
    # Cache files created before the "type" column existed need an in-place upgrade;
    # CREATE TABLE IF NOT EXISTS above leaves an already-existing table untouched.
    existing_columns = {row["name"] for row in connection.execute("PRAGMA table_info(records)")}
    if "type" not in existing_columns:
        connection.execute("ALTER TABLE records ADD COLUMN type TEXT")
    if stale_tables:
        _rekey_to_citizen_id(connection, stale_tables)
    _stop_reusing_file_ids(connection)
    connection.commit()
    return connection


def metadata_values(connection):
    return {row["key"]: row["value"] for row in connection.execute("SELECT key, value FROM cache_metadata")}


def stored_record_types(connection, workbook):
    rows = connection.execute("SELECT citizen_id, type FROM record_types WHERE workbook = ?", (str(workbook),))
    return {row["citizen_id"]: row["type"] for row in rows}


def set_record_type(workbook, citizen_id, type_code):
    """Persist the NVQS status picked in the form, keeping it out of the rebuilt `records` table."""
    connection = database_connection()
    try:
        with connection:
            if type_code:
                connection.execute(
                    "INSERT OR REPLACE INTO record_types (workbook, citizen_id, type) VALUES (?, ?, ?)",
                    (str(workbook), display_value(citizen_id), type_code),
                )
            else:
                connection.execute(
                    "DELETE FROM record_types WHERE workbook = ? AND citizen_id = ?",
                    (str(workbook), display_value(citizen_id)),
                )
    finally:
        connection.close()


def move_record_owner(workbook, old_citizen_id, new_citizen_id):
    """Carry the status and documents across when a record's CCCD is corrected."""
    old_citizen_id, new_citizen_id = display_value(old_citizen_id), display_value(new_citizen_id)
    if not old_citizen_id or not new_citizen_id or old_citizen_id == new_citizen_id:
        return
    connection = database_connection()
    try:
        with connection:
            connection.execute(
                "UPDATE record_files SET citizen_id = ? WHERE workbook = ? AND citizen_id = ?",
                (new_citizen_id, str(workbook), old_citizen_id),
            )
            # The status is one row per owner, so clear any stub already sitting on the new CCCD.
            connection.execute(
                "DELETE FROM record_types WHERE workbook = ? AND citizen_id = ?",
                (str(workbook), new_citizen_id),
            )
            connection.execute(
                "UPDATE record_types SET citizen_id = ? WHERE workbook = ? AND citizen_id = ?",
                (new_citizen_id, str(workbook), old_citizen_id),
            )
    finally:
        connection.close()


def add_record_files(workbook, citizen_id, source_name, pages):
    """Store the rendered pages of one uploaded document against a record."""
    created_at = datetime.now().isoformat()
    connection = database_connection()
    try:
        with connection:
            connection.executemany(
                """
                INSERT INTO record_files (workbook, citizen_id, source_name, page, image, thumbnail, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                [
                    (str(workbook), display_value(citizen_id), source_name, number, image, thumbnail, created_at)
                    for number, (image, thumbnail) in enumerate(pages, start=1)
                ],
            )
    finally:
        connection.close()


def record_files(workbook, citizen_id):
    """List the stored pages for a record without pulling the image data along."""
    if not display_value(citizen_id):
        return []
    connection = database_connection()
    try:
        rows = connection.execute(
            """
            SELECT id, source_name, page, created_at
            FROM record_files WHERE workbook = ? AND citizen_id = ?
            ORDER BY created_at, id
            """,
            (str(workbook), display_value(citizen_id)),
        ).fetchall()
        return [dict(row) for row in rows]
    finally:
        connection.close()


def record_file_image(workbook, citizen_id, file_id, thumbnail=False):
    column = "thumbnail" if thumbnail else "image"
    connection = database_connection()
    try:
        row = connection.execute(
            f"SELECT {column} AS data FROM record_files WHERE id = ? AND workbook = ? AND citizen_id = ?",
            (file_id, str(workbook), display_value(citizen_id)),
        ).fetchone()
        return row["data"] if row is not None else None
    finally:
        connection.close()


def delete_record_file(workbook, citizen_id, file_id):
    connection = database_connection()
    try:
        with connection:
            cursor = connection.execute(
                "DELETE FROM record_files WHERE id = ? AND workbook = ? AND citizen_id = ?",
                (file_id, str(workbook), display_value(citizen_id)),
            )
        return cursor.rowcount > 0
    finally:
        connection.close()


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
                religion_key, address_key, citizen_id_key, data_json, type
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        count = 0
        batch = []
        loaded_at = datetime.now()
        stored_types = stored_record_types(connection, path)
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
                    # What the manager picked in the form wins: it is never written back to the
                    # workbook, so a stale "Diện" column must not overwrite it. That optional
                    # column is only a starting value for records nobody has set in the app yet.
                    stored_types.get(citizen_id) or NvqsStatus.type_for(data.get("Diện", "")),
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
        "type": row["type"],
        "dien": NvqsStatus.value_for(row["type"]),
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
    dien = filters.get("dien", "")
    if dien == UNSET_DIEN_FILTER:
        conditions.append("(type IS NULL OR type = '')")
    elif dien:
        conditions.append("type = ?")
        parameters.append(dien)
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

"""Shared in-memory workbook state and the on-disk settings/cache paths."""

import logging
import os
import sys
import tempfile
import threading
from pathlib import Path

from quan_ly_ho_so.config import BASE_DIR
from quan_ly_ho_so.errors import WorkbookError

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
            # Verify directory is actually writable
            test_file = directory / ".write_test"
            test_file.write_text("ok", encoding="utf-8")
            test_file.unlink(missing_ok=True)
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

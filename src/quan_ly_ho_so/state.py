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
    """Return a writable path for settings and the SQLite database.

    They live beside the application, so the manager's data travels with the folder. If that
    folder is read-only (for example under Program Files), fall back to the per-user
    application-data folder, then to the temp folder, rather than preventing the workbook
    from loading altogether.
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

    preferred = BASE_DIR
    fallbacks = (root / "Quan_Ly_Ho_So_Thanh_Nien", Path(tempfile.gettempdir()) / "Quan_Ly_Ho_So_Thanh_Nien")
    for directory in (preferred, *fallbacks):
        try:
            directory.mkdir(parents=True, exist_ok=True)
            # mkdir succeeds on a read-only folder that already exists, so prove it is writable.
            with tempfile.TemporaryFile(dir=directory):
                pass
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

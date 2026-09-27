"""High-level workbook lifecycle: persisted settings, load/refresh, file picker."""

import json
import logging
import subprocess
import sys
from pathlib import Path

from quan_ly_ho_so.state import settings_path, state, state_lock
from quan_ly_ho_so.workbook.cache import (
    cached_workbook_info,
    file_signature,
    sync_workbook_to_database,
    validate_workbook_path,
)


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


def native_workbook_dialog():
    root = None
    try:
        if sys.platform == "darwin":
            script = '''
                try
                    set selectedFile to choose file with prompt "Chọn tệp Excel" of type {"org.openxmlformats.spreadsheetml.sheet", "org.openxmlformats.spreadsheetml.sheet.macroenabled"}
                    return POSIX path of selectedFile
                on error number -128
                    return ""
                end try
            '''
            result = subprocess.run(
                ["osascript", "-e", script],
                check=True,
                capture_output=True,
                text=True,
            )
            selected = result.stdout.strip()
            return Path(selected) if selected else None

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

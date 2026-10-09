"""Related paperwork kept as ordinary files, one folder per CCCD, beside the application."""

import os
import shutil
from datetime import datetime
from pathlib import Path

from quan_ly_ho_so.config import DOCUMENT_DIR
from quan_ly_ho_so.utils.text import display_value
from quan_ly_ho_so.word.export import safe_filename, unique_output_path

PARTIAL_SUFFIX = ".part"


def owner_dir(citizen_id):
    return DOCUMENT_DIR / safe_filename(display_value(citizen_id))


def save_document(citizen_id, filename, data):
    """Write one document into the owner's folder, never replacing a file of the same name."""
    if not display_value(citizen_id):
        raise ValueError("A document needs the owner's CCCD.")
    folder = owner_dir(citizen_id)
    folder.mkdir(parents=True, exist_ok=True)
    name = Path(filename)
    target = unique_output_path(folder, name.stem, name.suffix)
    # Write beside the target, then rename: a failed write must not leave a half file that lists as a document.
    partial = target.with_name(target.name + PARTIAL_SUFFIX)
    try:
        partial.write_bytes(data)
        os.replace(partial, target)
    except OSError:
        partial.unlink(missing_ok=True)
        raise
    return target.name


def list_documents(citizen_id):
    """The owner's documents, oldest first: [{"name", "size", "modified"}]."""
    if not display_value(citizen_id):
        return []
    folder = owner_dir(citizen_id)
    if not folder.is_dir():
        return []
    files = [path for path in folder.iterdir() if path.is_file() and path.suffix != PARTIAL_SUFFIX]
    files.sort(key=lambda path: (path.stat().st_mtime, path.name))
    return [
        {"name": path.name, "size": path.stat().st_size, "modified": datetime.fromtimestamp(path.stat().st_mtime)}
        for path in files
    ]


def document_path(citizen_id, name):
    """Path of one of the owner's documents, or None. Only the file name counts, so no `../`."""
    if not display_value(citizen_id):
        return None
    path = owner_dir(citizen_id) / Path(name).name
    return path if path.is_file() and path.suffix != PARTIAL_SUFFIX else None


def delete_document(citizen_id, name):
    path = document_path(citizen_id, name)
    if path is None:
        return False
    path.unlink()
    return True


def move_documents(old_citizen_id, new_citizen_id):
    """Carry the documents across when a record's CCCD is corrected."""
    if not display_value(old_citizen_id) or not display_value(new_citizen_id) or display_value(old_citizen_id) == display_value(new_citizen_id):
        return
    old_folder = owner_dir(old_citizen_id)
    if not old_folder.is_dir():
        return
    new_folder = owner_dir(new_citizen_id)
    new_folder.mkdir(parents=True, exist_ok=True)
    for path in old_folder.iterdir():
        target = unique_output_path(new_folder, path.stem, path.suffix)
        shutil.move(str(path), str(target))
    old_folder.rmdir()

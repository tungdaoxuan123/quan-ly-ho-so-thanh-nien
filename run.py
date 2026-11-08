#!/usr/bin/env python3
"""Entry point: puts src/ on the import path, then starts the app."""

import sys
from pathlib import Path

# Redirected Windows consoles can default to an encoding without Vietnamese.
for stream in (sys.stdout, sys.stderr):
    if stream is not None and hasattr(stream, "reconfigure"):
        stream.reconfigure(encoding="utf-8", errors="backslashreplace")

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

if __name__ == "__main__":
    if "--choose-workbook" in sys.argv[1:]:
        # This GUI owns the main thread of a fresh process, never a Flask worker.
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        try:
            selected = filedialog.askopenfilename(
                parent=root, title="Chọn tệp Excel",
                filetypes=[("Tệp Excel", "*.xlsx *.xlsm")],
            )
            print(selected, end="")
        finally:
            root.destroy()
    else:
        from quan_ly_ho_so.app import main
        main()

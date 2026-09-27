#!/usr/bin/env python3
"""Entry point: puts src/ on the import path, then starts the app."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from quan_ly_ho_so.app import main

if __name__ == "__main__":
    main()

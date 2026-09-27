"""Test package bootstrap: makes the src-layout quan_ly_ho_so package importable
without installing it, so `python -m unittest discover -s tests` works as-is."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

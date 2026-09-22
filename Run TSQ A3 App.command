#!/bin/zsh
set -e

TASK_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="$TASK_DIR/.venv/bin/python3"

if [[ ! -x "$PYTHON_BIN" ]]; then
  osascript -e 'display dialog "Run the virtual-environment setup in README.md before opening this app." buttons {"OK"} default button "OK" with icon caution'
  exit 1
fi

"$PYTHON_BIN" "$TASK_DIR/tsq_a3_app.py"

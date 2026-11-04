#!/bin/zsh
# Opens the app to other computers on the same local network (LAN).
TASK_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="$TASK_DIR/.venv/bin/python3"

if [[ ! -x "$PYTHON_BIN" ]]; then
  echo "Chua co .venv. Lam theo huong dan cai dat trong README.md truoc."
  exit 1
fi

# Asks for a new admin password in this window when the admin still uses the default one.
"$PYTHON_BIN" "$TASK_DIR/run.py" --setup || exit 1

export QLHS_HOST=0.0.0.0
echo
echo "May khac cung wifi vao bang dia chi hien o dong 'May khac trong mang noi bo vao bang' ben duoi."
echo "Neu wifi chan cac may noi chuyen voi nhau thi dung 'Chay Online.command'."
echo
"$PYTHON_BIN" "$TASK_DIR/run.py"

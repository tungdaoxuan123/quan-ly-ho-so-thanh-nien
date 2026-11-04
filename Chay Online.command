#!/bin/zsh
# One click: set the admin password if needed, start the app, and publish it with a Cloudflare tunnel.
TASK_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="$TASK_DIR/.venv/bin/python3"
PORT="${FLASK_RUN_PORT:-3000}"

if [[ ! -x "$PYTHON_BIN" ]]; then
  echo "Chua co .venv. Lam theo huong dan cai dat trong README.md truoc."
  exit 1
fi
if ! command -v cloudflared >/dev/null 2>&1; then
  echo "Chua cai cloudflared. Cai bang lenh:  brew install cloudflared"
  echo "Sau do bam lai file nay."
  exit 1
fi

# Asks for a new admin password in this window when the admin still uses the default one.
"$PYTHON_BIN" "$TASK_DIR/run.py" --setup || exit 1

# Only the tunnel on this computer may reach the app: it listens on 127.0.0.1, not on the LAN.
export QLHS_ONLINE=1 QLHS_HOST=127.0.0.1 OPEN_BROWSER=0 FLASK_RUN_PORT="$PORT"
"$PYTHON_BIN" "$TASK_DIR/run.py" &
APP_PID=$!
trap 'kill $APP_PID 2>/dev/null' EXIT INT TERM

for _ in {1..30}; do
  curl -s -o /dev/null "http://127.0.0.1:$PORT/login" && break
  kill -0 $APP_PID 2>/dev/null || { echo "Ung dung khong khoi dong duoc."; exit 1; }
  sleep 1
done

echo
echo "Dang mo duong ham. Tim dong 'https://....trycloudflare.com' ben duoi va gui cho moi nguoi."
echo "Khach tra cuu vao them /tra-cuu o cuoi dia chi. Dong cua so nay de tat truy cap tu ben ngoai."
echo
cloudflared tunnel --url "http://127.0.0.1:$PORT"

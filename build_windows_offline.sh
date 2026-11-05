#!/bin/bash
# Builds dist/QuanLyHoSo-Windows.zip on a Mac or Linux machine with internet: the official Windows embeddable
# Python, every library already downloaded for Windows, and the app. Copy the zip to a Windows 10/11 PC
# (no internet or installation needed there), extract it and double-click "Chay Quan Ly Ho So (LAN).bat".
set -euo pipefail
cd "$(dirname "$0")"

PY_VERSION=3.13.17
PY_TAG=313
NAME=QuanLyHoSo-Windows
OUT="dist/$NAME"
CACHE=dist/.cache
PYTHON_ZIP="$CACHE/python-$PY_VERSION-embed-amd64.zip"
HOST_PYTHON=".venv/bin/python3"
[[ -x "$HOST_PYTHON" ]] || HOST_PYTHON=python3

rm -rf "$OUT" "dist/$NAME.zip"
mkdir -p "$OUT/python" "$CACHE"
if [[ ! -f "$PYTHON_ZIP" ]]; then
  curl -fL --retry 3 -o "$PYTHON_ZIP.part" "https://www.python.org/ftp/python/$PY_VERSION/python-$PY_VERSION-embed-amd64.zip"
  mv "$PYTHON_ZIP.part" "$PYTHON_ZIP"
fi
unzip -q "$PYTHON_ZIP" -d "$OUT/python"

# The embeddable Python only searches the paths in its ._pth file and skips site-packages (and the .pth files
# pywin32 relies on) until "import site" is listed.
printf 'python%s.zip\r\n.\r\nLib\\site-packages\r\nimport site\r\n' "$PY_TAG" > "$OUT/python/python$PY_TAG._pth"

# pip evaluates environment markers against this machine, not Windows, so the Windows-only packages
# (pywin32 for Word -> PDF, colorama for click) are named explicitly.
"$HOST_PYTHON" -m pip install --quiet --disable-pip-version-check --no-compile \
  --target "$OUT/python/Lib/site-packages" \
  --platform win_amd64 --python-version "$PY_VERSION" --implementation cp --abi "cp$PY_TAG" --only-binary=:all: \
  -r requirements.txt pywin32 colorama

# Only application files, tracked or new but never git-ignored: no Excel data, databases, uploads or HTTPS keys.
git ls-files -z --cached --others --exclude-standard run.py src template 'Mau_*' README.md .env.example \
  "Chay Quan Ly Ho So (LAN).bat" "Mo tuong lua Windows (Run as administrator).bat" \
  "Dat IP co dinh.bat" "Tra lai IP tu dong.bat" ip_co_dinh.ps1 \
  | rsync -a --from0 --files-from=- ./ "$OUT/"
# cmd.exe misreads some LF-only batch files; give the Windows scripts Windows line endings.
for bat in "$OUT"/*.bat "$OUT"/*.ps1; do
  perl -pi -e 's/\r?\n/\r\n/' "$bat"
done

(cd dist && zip -qr "$NAME.zip" "$NAME")
echo "Built dist/$NAME.zip ($(du -h "dist/$NAME.zip" | cut -f1))"

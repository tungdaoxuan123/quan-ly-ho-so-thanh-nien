"""Exercise the ZIP through real HTTP, in an isolated directory (Windows only)."""
import http.cookiejar
import io
import json
import os
from pathlib import Path
import re
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from tests.test_quan_ly_ho_so import create_test_workbook


def main():
    archive = Path(sys.argv[1] if len(sys.argv) > 1 else "QuanLyHoSo-Windows-Offline.zip").resolve()
    checks = []
    with tempfile.TemporaryDirectory(prefix="qlhs-offline-") as temp:
        root = Path(temp)
        with zipfile.ZipFile(archive) as bundle:
            assert bundle.testzip() is None
            assert not any(Path(n).suffix in (".db", ".sqlite3", ".pem") or "/uploads/" in n for n in bundle.namelist())
            bundle.extractall(root)
        checks.append("ZIP integrity; no database, uploads or private keys")
        folder = root / "QuanLyHoSoThanhNien"
        for name in ("QuanLyHoSoThanhNien.exe", "template", "Mau_Ho_So_Thanh_Nien.docx", "Mau_Danh_Sach_Thanh_Nien.xlsx"):
            assert (folder / name).exists(), name
        fixture = root / "records.xlsx"
        create_test_workbook(fixture)
        with socket.socket() as sock:
            sock.bind(("127.0.0.1", 0))
            port = sock.getsockname()[1]
        env = {k: v for k, v in os.environ.items() if not k.startswith(("QLHS_", "FLASK_", "PYTHON"))}
        env.update(OPEN_BROWSER="0", FLASK_RUN_PORT=str(port), QLHS_HOST="127.0.0.1")
        base = f"http://127.0.0.1:{port}"
        client = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(http.cookiejar.CookieJar()), urllib.request.ProxyHandler({}))

        def request(path, data=None, headers=None):
            if isinstance(data, dict):
                data = urllib.parse.urlencode(data).encode()
            return client.open(urllib.request.Request(base + path, data=data, headers=headers or {}), timeout=10)

        def page(path):
            with request(path) as response:
                assert response.status == 200
                return response.read().decode("utf-8")

        def csrf(html):
            return re.search(r'name="csrf_token" value="([^"]+)"', html).group(1)

        with (root / "server.log").open("wb") as log:
            process = subprocess.Popen([str(folder / "QuanLyHoSoThanhNien.exe")], cwd=root, env=env, stdout=log, stderr=log)
            try:
                for _ in range(100):
                    assert process.poll() is None, (root / "server.log").read_text(errors="replace")
                    try:
                        html = page("/login")
                        break
                    except (OSError, urllib.error.URLError):
                        time.sleep(0.2)
                else:
                    raise AssertionError("Server failed to start")
                checks.append("Standalone EXE starts from unrelated working directory")
                assert "tra-cuu" in page("/")
                assert "login" in request("/admin").url
                checks.append("Guest lookup and protected admin redirect")
                with request("/login", dict(username="admin", password="admin123", csrf_token=csrf(html))) as response:
                    response.read()
                admin = page("/admin")
                assert "admin" in admin
                token = csrf(admin)
                boundary = "OfflineTestBoundary"
                body = (f'--{boundary}\r\nContent-Disposition: form-data; name="csrf_token"\r\n\r\n{token}\r\n'
                        f'--{boundary}\r\nContent-Disposition: form-data; name="workbook"; filename="records.xlsx"\r\nContent-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n').encode() + fixture.read_bytes() + f"\r\n--{boundary}--\r\n".encode()
                with request("/load", body, {"Content-Type": f"multipart/form-data; boundary={boundary}"}) as response:
                    assert "NGUYỄN VĂN AN" in response.read().decode()
                checks.append("Admin login and Excel multipart upload")
                for path in ("/", "/?q=nguyen", "/?birth_year=2005", "/person/new", "/person/1/edit", "/person/1/preview", "/profile", "/admin", "/admin/approvals", "/audit-log", "/api/status"):
                    page(path)
                checks.append("Search, filters, create/edit/preview pages, profile, approvals, audit and status")
                with request("/generate/1", {"csrf_token": csrf(page("/person/1/edit"))}) as response:
                    output = response.read()
                    with zipfile.ZipFile(io.BytesIO(output)) as word:
                        assert "word/document.xml" in word.namelist()
                checks.append("Real Word export from bundled template")
                before = (folder / "uploads" / "records.xlsx").read_bytes()
                with request("/person/save", {"csrf_token": "invalid"}) as response:
                    assert "Biểu mẫu đã hết hạn" in response.read().decode()
                assert (folder / "uploads" / "records.xlsx").read_bytes() == before
                checks.append("Invalid CSRF rejected")
                assert "login" in request("/logout").url
                checks.append("Logout")
            finally:
                process.terminate()
                process.wait(timeout=10)
        assert (folder / "uploads" / "records.xlsx").exists()
        print(json.dumps({"checks": checks, "result": "PASS"}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

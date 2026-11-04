import os
import re
import tempfile
import unittest
from unittest import mock
from pathlib import Path
from openpyxl import Workbook, load_workbook

import quan_ly_ho_so.app as app
from quan_ly_ho_so.auth import create_user, set_auth_database_path
from quan_ly_ho_so.audit import get_audit_logs
from quan_ly_ho_so.workbook.cache import find_record, find_record_by_citizen_id


def create_sample_workbook(path):
    headers = ["STT"]
    for _, section_headers in app.FORM_GROUPS:
        for header in section_headers:
            if header not in headers:
                headers.append(header)
    records = [
        {
            "STT": 1,
            "TÊN THƯỜNG DÙNG": "NGUYỄN VĂN AN",
            "Năm sinh": 2005,
            "Căn cước": "001205000001",
            "Nghề nghiệp": "Sinh viên",
        },
        {
            "STT": 2,
            "TÊN THƯỜNG DÙNG": "TRẦN THỊ BÌNH",
            "Năm sinh": 2004,
            "Căn cước": "001204000002",
            "Nghề nghiệp": "Nông dân",
        },
    ]
    wb = Workbook()
    sheet = wb.active
    sheet.title = "Hồ sơ"
    sheet.append(headers)
    sheet.append([])
    for record in records:
        sheet.append([record.get(header) for header in headers])
    wb.save(path)
    wb.close()


class NewFeaturesTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workbook_path = Path(self.temp_dir.name) / "records.xlsx"
        self.previous_database = app.state["database"]
        app.state["database"] = Path(self.temp_dir.name) / "cache.sqlite3"
        set_auth_database_path(Path(self.temp_dir.name) / "users.db")
        create_sample_workbook(self.workbook_path)
        app.set_workbook(self.workbook_path, persist=False)

    def tearDown(self):
        set_auth_database_path(None)
        app.state.update(
            workbook=None,
            sheet=None,
            headers={},
            database=self.previous_database,
            record_count=0,
            signature=None,
            loaded_at=None,
            error=None,
            last_output_dir=None,
        )
        self.temp_dir.cleanup()

    def _login(self, client, username="admin", password="admin123"):
        login_page = client.get("/login")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', login_page.get_data(as_text=True)).group(1)
        client.post("/login", data={"username": username, "password": password, "csrf_token": csrf})

    def test_soft_delete_keeps_excel_untouched_hides_record_and_is_audited(self):
        client = app.app.test_client()
        self._login(client)
        excel_before = self.workbook_path.read_bytes()

        home = client.get("/")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', home.get_data(as_text=True)).group(1)

        resp = client.post("/bulk/delete", data={"csrf_token": csrf, "stt": ["1"]}, follow_redirects=True)
        self.assertEqual(resp.status_code, 200)
        self.assertIn("Đã xóa 1 hồ sơ", resp.get_data(as_text=True))

        # The record is hidden, but the workbook itself is not touched (no backup either).
        self.assertEqual(app.state["record_count"], 1)
        self.assertEqual(find_record("1")["is_deleted"], True)
        self.assertEqual(self.workbook_path.read_bytes(), excel_before)
        self.assertEqual(list(self.workbook_path.parent.glob("*.backup-*")), [])

        logs, _ = get_audit_logs()
        deletion = next(entry for entry in logs if entry["action"] == "delete")
        self.assertEqual((deletion["username"], deletion["stt"]), ("admin", "1"))

        # ...and can be brought back.
        client.post("/bulk/restore", data={"csrf_token": csrf, "stt": ["1"]}, follow_redirects=True)
        self.assertEqual(app.state["record_count"], 2)
        self.assertFalse(find_record("1")["is_deleted"])

    def test_crud_audit_logging(self):
        client = app.app.test_client()
        self._login(client)

        # 1. Update an existing record
        edit_page = client.get("/person/2/edit")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', edit_page.get_data(as_text=True)).group(1)
        sig = re.search(r'name="signature" value="([^"]*)"', edit_page.get_data(as_text=True)).group(1)

        name_col = app.state["headers"]["TÊN THƯỜNG DÙNG"]
        cccd_col = app.state["headers"]["Căn cước"]
        client.post(
            "/person/save",
            data={
                "csrf_token": csrf,
                "original_stt": "2",
                "signature": sig,
                f"column_{name_col}": "TRẦN THỊ BÌNH (ĐÃ SỬA)",
                f"column_{cccd_col}": "001204000002",
            },
            follow_redirects=True,
        )

        logs, count = get_audit_logs()
        self.assertGreaterEqual(count, 1)
        update_log = next((l for l in logs if l["action"] == "update"), None)
        self.assertIsNotNone(update_log)
        self.assertEqual(update_log["username"], "admin")
        self.assertEqual(update_log["stt"], "2")
        self.assertIn("changes", update_log["details"])
        change_fields = [c["field"] for c in update_log["details"]["changes"]]
        self.assertIn("TÊN THƯỜNG DÙNG", change_fields)

        # 2. Check admin audit log web view
        audit_resp = client.get("/admin/audit-log")
        self.assertEqual(audit_resp.status_code, 200)
        self.assertIn("Nhật ký hoạt động", audit_resp.get_data(as_text=True))
        self.assertIn("Cập nhật", audit_resp.get_data(as_text=True))
        self.assertIn("TRẦN THỊ BÌNH (ĐÃ SỬA)", audit_resp.get_data(as_text=True))

        # 3. Check home page action dropdown button and row-level audit log link
        home_resp = client.get("/")
        home_html = home_resp.get_data(as_text=True)
        self.assertIn('class="row-menu"', home_html)
        self.assertIn('/admin/audit-log?stt=2', home_html)

        # 4. Check filtered audit log by STT
        stt_resp = client.get("/admin/audit-log?stt=2")
        self.assertEqual(stt_resp.status_code, 200)
        stt_html = stt_resp.get_data(as_text=True)
        self.assertIn("Lịch sử thao tác cho riêng hồ sơ STT 2", stt_html)
        self.assertIn("TRẦN THỊ BÌNH (ĐÃ SỬA)", stt_html)

        # Filtering for nonexistent STT returns 0 records
        empty_stt_resp = client.get("/admin/audit-log?stt=999999")
        self.assertIn("Chưa có lượt thao tác nào cho hồ sơ STT 999999", empty_stt_resp.get_data(as_text=True))

    def test_guest_lookup_page_unauthenticated(self):
        client = app.app.test_client()

        # 1. Access guest lookup without authentication (GET /tra-cuu)
        resp = client.get("/tra-cuu")
        self.assertEqual(resp.status_code, 200)
        html = resp.get_data(as_text=True)
        self.assertIn("Tra cứu hồ sơ thanh niên", html)
        self.assertIn("Số căn cước công dân (CCCD)", html)

        csrf = re.search(r'name="csrf_token" value="([^"]+)"', html).group(1)

        # 2. Lookup existing CCCD -> Readonly mode with hover tooltip
        resp_existing = client.post("/tra-cuu", data={"cccd": "001205000001", "csrf_token": csrf})
        self.assertEqual(resp_existing.status_code, 200)
        existing_html = resp_existing.get_data(as_text=True)
        self.assertIn("Hồ sơ đã tồn tại", existing_html)
        self.assertIn("Kiểm tra thông tin và liên hệ admin nếu cần sửa đổi", existing_html)
        self.assertIn("NGUYỄN VĂN AN", existing_html)
        self.assertIn("readonly", existing_html)

        # 3. Lookup non-existing CCCD -> asked first, no form yet
        resp_ask = client.post("/tra-cuu", data={"cccd": "001209999999", "csrf_token": csrf})
        self.assertEqual(resp_ask.status_code, 200)
        ask_html = resp_ask.get_data(as_text=True)
        self.assertIn("Chưa có hồ sơ", ask_html)
        self.assertIn("Có, thêm thông tin của tôi", ask_html)
        self.assertNotIn("Gửi hồ sơ", ask_html)
        self.assertNotIn('action="/tra-cuu/save"', ask_html)

        # ...and only after agreeing does the form appear
        resp_new = client.post("/tra-cuu/them", data={"cccd": "001209999999", "csrf_token": csrf})
        self.assertEqual(resp_new.status_code, 200)
        new_html = resp_new.get_data(as_text=True)
        self.assertIn("Không tìm thấy hồ sơ với CCCD", new_html)
        self.assertIn("001209999999", new_html)
        self.assertIn("Gửi hồ sơ", new_html)

        self.assertIn("Hồ sơ đã tồn tại", client.post("/tra-cuu/them", data={"cccd": "001205000001", "csrf_token": csrf}).get_data(as_text=True))
        self.assertEqual(client.post("/tra-cuu/them", data={"cccd": "001209999999", "csrf_token": "x"}).status_code, 302)

        # 4. Guest submit new record
        sig = re.search(r'name="signature" value="([^"]*)"', new_html).group(1)
        name_col = app.state["headers"]["TÊN THƯỜNG DÙNG"]
        cccd_col = app.state["headers"]["Căn cước"]

        save_resp = client.post(
            "/tra-cuu/save",
            data={
                "csrf_token": csrf,
                "cccd": "001209999999",
                "signature": sig,
                f"column_{name_col}": "HOÀNG GIA BẢO",
                f"column_{cccd_col}": "001209999999",
            },
            follow_redirects=True,
        )
        self.assertEqual(save_resp.status_code, 200)
        save_html = save_resp.get_data(as_text=True)
        self.assertIn("Đã tạo hồ sơ thành công", save_html)
        self.assertIn("HOÀNG GIA BẢO", save_html)
        self.assertIn("Kiểm tra thông tin và liên hệ admin nếu cần sửa đổi", save_html)

        # Verify record exists in cache
        record = find_record_by_citizen_id("001209999999")
        self.assertIsNotNone(record)
        self.assertEqual(record["name"], "HOÀNG GIA BẢO")


    def _login_attempt(self, client, password, **kwargs):
        page = client.get("/login", **kwargs)
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', page.get_data(as_text=True)).group(1)
        return client.post("/login", data={"username": "admin", "password": password, "csrf_token": csrf}, **kwargs)

    def test_repeated_wrong_passwords_lock_login_even_for_the_right_password(self):
        app.login_failures._hits.clear()
        client = app.app.test_client()
        for _ in range(5):
            self.assertEqual(self._login_attempt(client, "sai-mat-khau").status_code, 200)
        locked = self._login_attempt(client, "admin123")
        self.assertEqual(locked.status_code, 429)
        self.assertIn("quá nhiều lần", locked.get_data(as_text=True))
        app.login_failures._hits.clear()

    def test_online_mode_opens_the_network_but_throttles_per_forwarded_address(self):
        app.login_failures._hits.clear()
        app.guest_lookups._hits.clear()
        client = app.app.test_client()
        public = {"environ_overrides": {"REMOTE_ADDR": "8.8.8.8"}}
        self.assertEqual(client.get("/login", **public).status_code, 403)
        with mock.patch.dict(os.environ, {"QLHS_ONLINE": "1"}):
            self.assertEqual(client.get("/login", **public).status_code, 200)

            # Through the tunnel every request is from 127.0.0.1; the forwarded address tells callers apart.
            tunnel = lambda ip: {"environ_overrides": {"REMOTE_ADDR": "127.0.0.1"}, "headers": {"CF-Connecting-IP": ip}}
            for _ in range(5):
                self._login_attempt(client, "sai", **tunnel("203.0.113.9"))
            self.assertEqual(self._login_attempt(client, "admin123", **tunnel("203.0.113.9")).status_code, 429)
            self.assertEqual(self._login_attempt(client, "admin123", **tunnel("203.0.113.10")).status_code, 302)

            # The forwarded header is ignored when the request did not come from this computer.
            spoof = {"environ_overrides": {"REMOTE_ADDR": "8.8.8.8"}, "headers": {"CF-Connecting-IP": "10.0.0.1"}}
            with app.app.test_request_context("/", **{"environ_overrides": spoof["environ_overrides"], "headers": spoof["headers"]}):
                self.assertEqual(app.client_ip(), "8.8.8.8")

            page = client.get("/tra-cuu").get_data(as_text=True)
            csrf = re.search(r'name="csrf_token" value="([^"]+)"', page).group(1)
            statuses = [
                client.post("/tra-cuu", data={"cccd": "001209999999", "csrf_token": csrf}, **tunnel("198.51.100.7")).status_code
                for _ in range(31)
            ]
            self.assertEqual(statuses[:30], [200] * 30)
            self.assertEqual(statuses[30], 302)
        app.login_failures._hits.clear()
        app.guest_lookups._hits.clear()

    def test_ensure_admin_password_refuses_without_terminal_and_sets_new_password(self):
        from werkzeug.security import check_password_hash
        from quan_ly_ho_so.auth import get_user_by_username

        with mock.patch("sys.stdin.isatty", return_value=False):
            with self.assertRaises(SystemExit):
                app.ensure_admin_password()
        answers = iter(["admin123", "short", "MatKhauMoi-2026", "khac-nhau-123", "MatKhauMoi-2026", "MatKhauMoi-2026"])
        with mock.patch("sys.stdin.isatty", return_value=True), mock.patch("getpass.getpass", side_effect=lambda _: next(answers)):
            app.ensure_admin_password()
        self.assertTrue(check_password_hash(get_user_by_username("admin")["password_hash"], "MatKhauMoi-2026"))
        app.ensure_admin_password()  # no longer default: returns without asking

    def test_flat_ca_list_reads_khu_pho_from_address_rejects_broken_cccd_and_appends_notes(self):
        import io
        from datetime import datetime
        from quan_ly_ho_so.workbook.ca_import import NOTE_HEADER, OTHER_INFO, changes_for, read_ca_file

        ca = Workbook()
        sheet = ca.active
        sheet.append(["STT", "HỌ VÀ TÊN", "Ngày tháng năm sinh", "CCCD", "Chủ hộ", "Địa chỉ thường trú", "Địa chỉ tạm trú",
                      "Tên cha", "Tên mẹ", "Đang thường trú hay tạm trú", "Ghi chú"])
        sheet.append([1, "NGUYỄN ĐỨC HUY", datetime(2008, 11, 13), 15208004933, "ĐOÀN THỊ KIM QUYÊN",
                      "2.06 Lô A, khu phố 29, Phường Hạnh Thông", "12 Phố B", "NGUYỄN ĐỨC NAM", "ĐOÀN THỊ KIM QUYÊN",
                      "Đang thường trú", "HSSV"])
        # Two CCCDs in one cell, a typed date, no khu phố in the address.
        sheet.append([2, "TRẦN VĂN HAI", "05/02/2006", "79209022887\n 79209011340", None, "1 Phố C, Phường Hạnh Thông"])
        # A number Excel already rounded: the digits are gone.
        sheet.append([3, "LÊ VĂN BA", datetime(2001, 1, 1), 2.52207e16, None, "Khu phố 83, Phường Hạnh Thông"])
        content = io.BytesIO()
        ca.save(content)

        huy, hai, ba = read_ca_file(content.getvalue())
        self.assertEqual((huy["citizen_id"], huy["quarter"], huy["notes"]), ("015208004933", "Khu phố 29", []))
        self.assertEqual(huy["values"]["Tên khai sinh"], "Nguyễn Đức Huy")
        self.assertEqual(huy["values"]["TÊN THƯỜNG DÙNG"], "NGUYỄN ĐỨC HUY")
        self.assertEqual([huy["values"][h] for h in ("Ngày sinh", "Tháng sinh", "Năm sinh")], [13, 11, 2008])
        self.assertEqual((huy["values"]["Thường trú"], huy["values"]["Nơi ở hiện nay"]), ("2.06 Lô A, khu phố 29, Phường Hạnh Thông", "12 Phố B"))
        self.assertEqual((huy["values"]["Tên cha"], huy["values"]["Tên mẹ"]), ("Nguyễn Đức Nam", "Đoàn Thị Kim Quyên"))
        self.assertEqual((huy["values"][NOTE_HEADER], huy["values"][OTHER_INFO]), ("HSSV", "Chủ hộ: Đoàn Thị Kim Quyên"))
        self.assertNotIn("type", huy)  # Ghi chú is only a note, never a Diện

        self.assertEqual((hai["citizen_id"], hai["quarter"]), ("", ""))
        self.assertEqual([hai["values"][h] for h in ("Ngày sinh", "Tháng sinh", "Năm sinh")], [5, 2, 2006])
        self.assertTrue(hai["notes"][0].startswith("CCCD không hợp lệ"))
        self.assertIn("không thấy khu phố trong địa chỉ thường trú, chưa gán khu phố", hai["notes"])
        self.assertEqual((ba["citizen_id"], ba["quarter"]), ("", ""))
        self.assertEqual(len(ba["notes"]), 2)  # broken CCCD, and Khu phố 83 does not exist

        # On an update, notes go below what is held (health notes, phone numbers stay), and only once.
        record = {"data": {NOTE_HEADER: "Cao 1m70", OTHER_INFO: "SĐT 0901", "Tên khai sinh": "Nguyễn Đức Huy"}, "quarter": "Khu phố 29"}
        changes = {change["header"]: change["new"] for change in changes_for(huy, record)}
        self.assertEqual(changes[NOTE_HEADER], "Cao 1m70\nHSSV")
        self.assertEqual(changes[OTHER_INFO], "SĐT 0901\nChủ hộ: Đoàn Thị Kim Quyên")
        self.assertNotIn("Tên khai sinh", changes)
        self.assertNotIn(None, changes)  # same khu phố: no change
        record["data"].update(changes)
        self.assertFalse({NOTE_HEADER, OTHER_INFO} & {change["header"] for change in changes_for(huy, record)})

    def test_ca_import_only_add_leaves_existing_people_untouched(self):
        import io

        ca = Workbook()
        sheet = ca.active
        sheet.append(["STT", "HỌ VÀ TÊN", "Ngày tháng năm sinh", "CCCD", "Địa chỉ thường trú"])
        sheet.append([1, "NGUYỄN VĂN AN ĐỔI TÊN", "01/01/2005", "001205000001", "1 Phố A, khu phố 3"])  # already held
        sheet.append([2, "LÊ VĂN MỚI", "02/02/2006", "079206000999", "2 Phố B, khu phố 4"])
        content = io.BytesIO()
        ca.save(content)

        client = app.app.test_client()
        self._login(client)
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', client.get("/").get_data(as_text=True)).group(1)
        client.post("/import/ca", data={"csrf_token": csrf, "ca_files": (io.BytesIO(content.getvalue()), "ca.xlsx")},
                    content_type="multipart/form-data")
        preview = client.get("/").get_data(as_text=True)
        self.assertIn('name="only_add" value="1">Chỉ thêm 1 mới, không cập nhật</button>', preview)

        excel_before = find_record_by_citizen_id("001205000001")["data"]
        client.post("/import/ca/confirm", data={"csrf_token": csrf, "only_add": "1"})
        page = client.get("/").get_data(as_text=True)
        self.assertIn("1 thêm mới", page)
        self.assertIn("0 cập nhật", page)
        self.assertIn("Có thay đổi trong file CA nhưng không cập nhật — đã chọn chỉ thêm mới (1)", page)
        self.assertEqual(find_record_by_citizen_id("001205000001")["data"], excel_before)
        self.assertEqual(find_record_by_citizen_id("001205000001")["quarter"], "")
        self.assertEqual(find_record_by_citizen_id("079206000999")["quarter"], "Khu phố 4")
        self.assertEqual(app.state["record_count"], 3)

    def test_lan_certificate_covers_address_and_reissues_only_when_address_changes(self):
        import ipaddress
        from cryptography import x509
        from quan_ly_ho_so.tls import lan_certificate

        directory = Path(self.temp_dir.name) / "HTTPS"
        cert_path, key_path = lan_certificate(directory, "192.168.1.69", "may-chu.local")
        ca = x509.load_pem_x509_certificate((directory / "ca.pem").read_bytes())
        cert = x509.load_pem_x509_certificate(cert_path.read_bytes())
        cert.verify_directly_issued_by(ca)
        ips = cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value.get_values_for_type(x509.IPAddress)
        self.assertEqual(ips, [ipaddress.ip_address("127.0.0.1"), ipaddress.ip_address("192.168.1.69")])
        names = cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value.get_values_for_type(x509.DNSName)
        self.assertEqual(names, ["localhost", "may-chu.local"])
        self.assertEqual(os.stat(key_path).st_mode & 0o777, 0o600)

        # Same address and name: nothing rewritten. New address or name: new server certificate, same CA,
        # so phones need no reinstall.
        first_cert, ca_bytes = cert_path.read_bytes(), (directory / "ca.pem").read_bytes()
        lan_certificate(directory, "192.168.1.69", "MAY-CHU.local")
        self.assertEqual(cert_path.read_bytes(), first_cert)
        lan_certificate(directory, "192.168.1.69", "may-khac.local")
        self.assertNotEqual(cert_path.read_bytes(), first_cert)
        second_cert = cert_path.read_bytes()
        lan_certificate(directory, "10.0.0.5", "may-khac.local")
        self.assertNotEqual(cert_path.read_bytes(), second_cert)
        self.assertEqual((directory / "ca.pem").read_bytes(), ca_bytes)
        x509.load_pem_x509_certificate(cert_path.read_bytes()).verify_directly_issued_by(ca)

        # The placeholder lan_address() returns on failure still yields a localhost certificate.
        lan_certificate(Path(self.temp_dir.name) / "fallback", "<địa chỉ IP của máy này>", "may-chu.local")

    def test_https_server_keeps_serving_while_a_client_connects_and_stays_silent(self):
        import socket
        import ssl
        import threading
        from werkzeug.serving import make_server
        from quan_ly_ho_so.tls import lan_certificate, lan_ssl_context

        directory = Path(self.temp_dir.name) / "HTTPS"
        server = make_server("127.0.0.1", 0, app.app, threaded=True, ssl_context=lan_ssl_context(*lan_certificate(directory, "127.0.0.1", "may-chu.local")))
        threading.Thread(target=server.serve_forever, daemon=True).start()
        port = server.socket.getsockname()[1]
        # Like a browser preconnect: open the TCP connection, never start the TLS handshake.
        silent = socket.create_connection(("127.0.0.1", port))
        try:
            client_context = ssl.create_default_context(cafile=str(directory / "ca.pem"))
            with socket.create_connection(("127.0.0.1", port), timeout=5) as raw:
                with client_context.wrap_socket(raw, server_hostname="127.0.0.1") as tls:
                    tls.sendall(b"GET /login HTTP/1.1\r\nHost: 127.0.0.1\r\nConnection: close\r\n\r\n")
                    self.assertTrue(tls.recv(64).startswith(b"HTTP/1.1 200"))
        finally:
            silent.close()
            server.shutdown()
            server.server_close()

    def test_ca_certificate_download_needs_no_login(self):
        client = app.app.test_client()
        https_dir = Path(self.temp_dir.name) / "HTTPS"
        with mock.patch.object(app, "HTTPS_DIR", https_dir):
            self.assertEqual(client.get("/chung-chi").status_code, 404)
            app.lan_certificate(https_dir, "192.168.1.69", "may-chu.local")
            resp = client.get("/chung-chi")
            self.assertEqual(resp.status_code, 200)
            self.assertEqual(resp.mimetype, "application/x-x509-ca-cert")
            self.assertEqual(resp.data, (https_dir / "ca.pem").read_bytes())
            resp.close()


if __name__ == "__main__":
    unittest.main()

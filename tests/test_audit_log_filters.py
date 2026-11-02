import re
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook

import quan_ly_ho_so.app as app
from quan_ly_ho_so.auth import set_auth_database_path, create_user
from quan_ly_ho_so.audit import get_audit_logs, get_audit_performers, log_action


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


class AuditLogFiltersTests(unittest.TestCase):
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

    def test_routes_and_access(self):
        client = app.app.test_client()

        # Unauthenticated: redirects to login
        resp = client.get("/audit-log")
        self.assertEqual(resp.status_code, 302)
        self.assertIn("/login", resp.headers["Location"])

        resp2 = client.get("/admin/audit-log")
        self.assertEqual(resp2.status_code, 302)
        self.assertIn("/login", resp2.headers["Location"])

        # Authenticated as admin: both endpoints work
        self._login(client)
        resp_audit = client.get("/audit-log")
        self.assertEqual(resp_audit.status_code, 200)
        html = resp_audit.get_data(as_text=True)
        self.assertIn("Nhật ký hoạt động", html)
        self.assertIn("Thời gian từ:", html)
        self.assertIn("Người thực hiện:", html)
        self.assertIn("Hành động:", html)
        self.assertIn("Tìm kiếm theo Người thực hiện, STT, Họ tên", html)

        resp_admin_audit = client.get("/admin/audit-log")
        self.assertEqual(resp_admin_audit.status_code, 200)

    def test_filters_by_user_action_and_date(self):
        client = app.app.test_client()
        self._login(client)

        today_str = datetime.now().strftime("%Y-%m-%d")
        yesterday_str = (datetime.now() - timedelta(days=2)).strftime("%Y-%m-%d")

        # Manually seed diverse audit entries
        log_action("admin", "create", stt="1", record_name="NGUYỄN VĂN AN", details={"fields": {"name": "AN"}})
        log_action("staff_user", "update", stt="2", record_name="TRẦN THỊ BÌNH", details={"changes": []})
        log_action("staff_user", "delete", stt="3", record_name="LÊ VĂN CƯỜNG", details={"reason": "soft_delete"})

        # Override timestamps directly to test date filtering
        from quan_ly_ho_so.audit import _audit_connection
        conn = _audit_connection()
        conn.execute("UPDATE audit_log SET timestamp = ? WHERE stt = '1'", (f"{yesterday_str}T10:00:00",))
        conn.execute("UPDATE audit_log SET timestamp = ? WHERE stt = '2'", (f"{today_str}T10:00:00",))
        conn.execute("UPDATE audit_log SET timestamp = ? WHERE stt = '3'", (f"{today_str}T12:00:00",))
        conn.commit()
        conn.close()

        # 1. Filter by username
        entries, total = get_audit_logs(username="staff_user")
        self.assertEqual(total, 2)
        self.assertTrue(all(e["username"] == "staff_user" for e in entries))

        resp_user = client.get("/audit-log?username=staff_user")
        self.assertEqual(resp_user.status_code, 200)
        user_html = resp_user.get_data(as_text=True)
        self.assertIn("TRẦN THỊ BÌNH", user_html)
        self.assertIn("LÊ VĂN CƯỜNG", user_html)
        self.assertNotIn("NGUYỄN VĂN AN", user_html)

        # 2. Filter by action
        entries, total = get_audit_logs(action="delete")
        self.assertEqual(total, 1)
        self.assertEqual(entries[0]["stt"], "3")

        resp_action = client.get("/audit-log?action=delete")
        self.assertEqual(resp_action.status_code, 200)
        action_html = resp_action.get_data(as_text=True)
        self.assertIn("LÊ VĂN CƯỜNG", action_html)
        self.assertNotIn("TRẦN THỊ BÌNH", action_html)

        # Filter by action alias (e.g. xoa_mem)
        entries, total = get_audit_logs(action="xoa_mem")
        self.assertEqual(total, 1)

        # 3. Filter by date
        entries, total = get_audit_logs(date_from=today_str, date_to=today_str)
        self.assertEqual(total, 2)
        self.assertEqual({e["stt"] for e in entries}, {"2", "3"})

        entries_old, total_old = get_audit_logs(date_to=yesterday_str)
        self.assertEqual(total_old, 1)
        self.assertEqual(entries_old[0]["stt"], "1")

        resp_date = client.get(f"/audit-log?date_from={yesterday_str}&date_to={yesterday_str}")
        self.assertEqual(resp_date.status_code, 200)
        date_html = resp_date.get_data(as_text=True)
        self.assertIn("NGUYỄN VĂN AN", date_html)
        self.assertNotIn("TRẦN THỊ BÌNH", date_html)

    def test_search_for_user_stt_and_name(self):
        client = app.app.test_client()
        self._login(client)

        log_action("admin", "create", stt="101", record_name="NGUYỄN VĂN AN")
        log_action("moderator", "update", stt="202", record_name="TRẦN THỊ BÌNH")
        log_action("admin", "update", stt="303", record_name="PHẠM QUANG ĐỨC")

        # Search by STT in q
        entries, total = get_audit_logs(query="202")
        self.assertEqual(total, 1)
        self.assertEqual(entries[0]["record_name"], "TRẦN THỊ BÌNH")

        resp_q_stt = client.get("/audit-log?q=202")
        self.assertIn("TRẦN THỊ BÌNH", resp_q_stt.get_data(as_text=True))
        self.assertNotIn("NGUYỄN VĂN AN", resp_q_stt.get_data(as_text=True))

        # Search by Họ tên in q (accent insensitive: 'duc' matches 'ĐỨC')
        entries, total = get_audit_logs(query="duc")
        self.assertEqual(total, 1)
        self.assertEqual(entries[0]["stt"], "303")

        resp_q_name = client.get("/audit-log?q=duc")
        self.assertIn("PHẠM QUANG ĐỨC", resp_q_name.get_data(as_text=True))

        # Search by performer in q
        entries, total = get_audit_logs(query="moderator")
        self.assertEqual(total, 1)
        self.assertEqual(entries[0]["username"], "moderator")

        resp_q_user = client.get("/audit-log?q=moderator")
        self.assertIn("TRẦN THỊ BÌNH", resp_q_user.get_data(as_text=True))

    def test_pagination_50_rows_limit(self):
        client = app.app.test_client()
        self._login(client)

        # Seed 75 records
        for i in range(1, 76):
            log_action("admin", "update", stt=str(i), record_name=f"Hồ sơ số {i}")

        # Page 1: returns 50 records
        entries_p1, total = get_audit_logs(limit=50, offset=0)
        self.assertEqual(total, 75)
        self.assertEqual(len(entries_p1), 50)
        # Should be newest first (id descending)
        self.assertEqual(entries_p1[0]["stt"], "75")
        self.assertEqual(entries_p1[-1]["stt"], "26")

        # Page 2: returns remaining 25 records
        entries_p2, total2 = get_audit_logs(limit=50, offset=50)
        self.assertEqual(total2, 75)
        self.assertEqual(len(entries_p2), 25)
        self.assertEqual(entries_p2[0]["stt"], "25")
        self.assertEqual(entries_p2[-1]["stt"], "1")

        # Web view check
        resp_p1 = client.get("/audit-log?page=1")
        self.assertEqual(resp_p1.status_code, 200)
        p1_html = resp_p1.get_data(as_text=True)
        self.assertIn("Tổng cộng 75 lượt thao tác", p1_html)
        self.assertIn("Trang 1 / 2", p1_html)
        self.assertIn("Tối đa 50 dòng/trang", p1_html)
        self.assertIn("Sau →", p1_html)

        resp_p2 = client.get("/audit-log?page=2")
        self.assertEqual(resp_p2.status_code, 200)
        p2_html = resp_p2.get_data(as_text=True)
        self.assertIn("Trang 2 / 2", p2_html)
        self.assertIn("← Trước", p2_html)

        # Pagination preserving search parameters
        resp_filtered_page = client.get("/audit-log?page=1&action=update&q=Hồ+sơ")
        fp_html = resp_filtered_page.get_data(as_text=True)
        # Ensure pagination link retains query parameter
        self.assertIn("page=2", fp_html)
        self.assertIn("action=update", fp_html)


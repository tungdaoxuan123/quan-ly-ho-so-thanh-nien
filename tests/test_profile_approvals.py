import re
import sqlite3
import tempfile
import unittest
from pathlib import Path
from openpyxl import Workbook, load_workbook

import quan_ly_ho_so.app as app
from quan_ly_ho_so.approvals import (
    IDENTITY_CRITICAL_HEADERS,
    approve_request,
    count_pending_requests,
    get_latest_request_by_cccd,
    get_pending_request_by_cccd,
    get_request_by_id,
    has_pending_request,
    is_identity_critical,
    list_requests,
    reject_request,
)
from quan_ly_ho_so.audit import _audit_database_path, get_audit_logs
from quan_ly_ho_so.auth import create_user, set_auth_database_path
from quan_ly_ho_so.workbook.cache import find_record_by_citizen_id


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
            "Tên khai sinh": "NGUYỄN VĂN AN",
            "Ngày sinh": 15,
            "Tháng sinh": 8,
            "Năm sinh": 2005,
            "Căn cước": "001205000001",
            "Nghề nghiệp": "Học sinh",
            "Thường trú": "Số 10 Đường Hoa, KP 1",
            "Dân tộc": "Kinh",
            "Tôn giáo": "Không",
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


class ProfileApprovalsWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workbook_path = Path(self.temp_dir.name) / "records.xlsx"
        self.previous_database = app.state["database"]
        app.state["database"] = Path(self.temp_dir.name) / "cache.sqlite3"
        set_auth_database_path(Path(self.temp_dir.name) / "users.db")
        create_sample_workbook(self.workbook_path)
        app.set_workbook(self.workbook_path, persist=False)
        self.client = app.app.test_client()

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

    def get_csrf(self, html_text):
        match = re.search(r'name="csrf_token" value="([^"]+)"', html_text)
        return match.group(1) if match else ""

    def submit_update(self, **changes):
        response = self.client.get("/tra-cuu")
        data = {
            "csrf_token": self.get_csrf(response.get_data(as_text=True)),
            "cccd": "001205000001",
        }
        for header, value in changes.items():
            data[f"column_{app.state['headers'][header]}"] = value
        self.client.post("/tra-cuu/cap-nhat", data=data)
        request = get_pending_request_by_cccd(data["cccd"])
        self.assertIsNotNone(request)
        return request

    def test_browser_newlines_do_not_add_unchanged_notes_to_request(self):
        notes_header = "TẤT CẢ THÔNG TIN KHÁC\nLIÊN QUAN ĐẾN THANH NIÊN"
        notes = "Cha: 0700000000\nMe: 0300000000\nTn: 0380000000"
        workbook = load_workbook(self.workbook_path)
        workbook.active.cell(3, app.state["headers"][notes_header]).value = notes
        workbook.save(self.workbook_path)
        workbook.close()
        app.refresh_state(force=True)
        pending = self.submit_update(**{
            "Nơi đăng ký khai sinh (2 cấp)": "Phường Hạnh Thông, TP.Hồ Chí Minh 123",
            notes_header: notes.replace("\n", "\r\n"),
        })
        self.assertEqual(len(pending["changes"]), 1)
        self.assertEqual(pending["changes"][0]["field"], "Nơi đăng ký khai sinh (2 cấp)")
        approve_request(pending["id"], "admin")
        workbook = load_workbook(self.workbook_path)
        self.assertEqual(workbook.active.cell(3, app.state["headers"][notes_header]).value, notes)
        workbook.close()

    def test_approval_preserves_unrelated_newer_edits_and_identity_cell_types(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        # Excel changed after submission; approval must refresh without needing a page visit.
        workbook = load_workbook(self.workbook_path)
        workbook.active.cell(3, app.state["headers"]["Thường trú"]).value = "Địa chỉ quản trị vừa sửa"
        workbook.save(self.workbook_path)
        workbook.close()

        approve_request(pending["id"], reviewer="admin")

        record = find_record_by_citizen_id("001205000001")
        self.assertEqual(record["data"]["Nghề nghiệp"], "Bác sĩ")
        self.assertEqual(record["data"]["Thường trú"], "Địa chỉ quản trị vừa sửa")
        workbook = load_workbook(self.workbook_path)
        try:
            self.assertEqual(workbook.active.cell(3, app.state["headers"]["Năm sinh"]).value, 2005)
        finally:
            workbook.close()

    def test_approval_rejects_conflicting_edits_without_writing_any_fields(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ", "Thường trú": "Địa chỉ đề xuất"})
        workbook = load_workbook(self.workbook_path)
        workbook.active.cell(3, app.state["headers"]["Nghề nghiệp"]).value = "Quản trị đã sửa"
        workbook.save(self.workbook_path)
        workbook.close()
        before = self.workbook_path.read_bytes()

        with self.assertRaisesRegex(ValueError, "Nghề nghiệp"):
            approve_request(pending["id"], reviewer="admin")

        self.assertEqual(self.workbook_path.read_bytes(), before)
        self.assertEqual(get_request_by_id(pending["id"])["status"], "pending")
        self.assertEqual(get_audit_logs(action="approve_update")[1], 0)

    def test_approval_does_not_fall_back_to_reused_stt(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        workbook = load_workbook(self.workbook_path)
        workbook.active.cell(3, app.state["headers"]["Căn cước"]).value = "001205999999"
        workbook.active.cell(3, app.state["headers"]["TÊN THƯỜNG DÙNG"]).value = "NGƯỜI KHÁC"
        workbook.save(self.workbook_path)
        workbook.close()
        before = self.workbook_path.read_bytes()

        with self.assertRaises(ValueError):
            approve_request(pending["id"], reviewer="admin")

        self.assertEqual(self.workbook_path.read_bytes(), before)
        self.assertEqual(get_request_by_id(pending["id"])["status"], "pending")

    def test_approval_follows_cccd_after_stt_is_renumbered(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        workbook = load_workbook(self.workbook_path)
        workbook.active.cell(3, app.state["headers"]["STT"]).value = 42
        workbook.save(self.workbook_path)
        workbook.close()

        approve_request(pending["id"], reviewer="admin")

        record = find_record_by_citizen_id("001205000001")
        self.assertEqual(record["stt"], "42")
        self.assertEqual(record["data"]["Nghề nghiệp"], "Bác sĩ")

    def test_request_cannot_be_approved_against_another_workbook(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        other_path = Path(self.temp_dir.name) / "other.xlsx"
        create_sample_workbook(other_path)  # Same CCCD, different source workbook.
        app.set_workbook(other_path, persist=False)
        before = other_path.read_bytes()

        with self.assertRaises(ValueError):
            approve_request(pending["id"], reviewer="admin")

        self.assertEqual(other_path.read_bytes(), before)
        self.assertFalse(has_pending_request("001205000001"))
        self.assertIsNone(get_latest_request_by_cccd("001205000001"))
        self.assertEqual(get_request_by_id(pending["id"])["status"], "pending")
        app.set_workbook(self.workbook_path, persist=False)
        self.assertTrue(has_pending_request("001205000001"))

    def test_approval_rejects_a_different_source_sheet(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        workbook = load_workbook(self.workbook_path)
        workbook.active.title = "Trang khác"
        workbook.save(self.workbook_path)
        workbook.close()
        before = self.workbook_path.read_bytes()

        with self.assertRaises(ValueError):
            approve_request(pending["id"], reviewer="admin")

        self.assertEqual(self.workbook_path.read_bytes(), before)
        self.assertEqual(get_request_by_id(pending["id"])["status"], "pending")

    def test_approval_matches_headers_after_columns_are_reordered(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        old_address = find_record_by_citizen_id("001205000001")["data"]["Thường trú"]
        occupation, address = (app.state["headers"][header] for header in ("Nghề nghiệp", "Thường trú"))
        workbook = load_workbook(self.workbook_path)
        sheet = workbook.active
        for row in range(1, sheet.max_row + 1):
            first, second = sheet.cell(row, occupation), sheet.cell(row, address)
            first.value, second.value = second.value, first.value
        workbook.save(self.workbook_path)
        workbook.close()

        approve_request(pending["id"], reviewer="admin")

        record = find_record_by_citizen_id("001205000001")
        self.assertEqual(record["data"]["Nghề nghiệp"], "Bác sĩ")
        self.assertEqual(record["data"]["Thường trú"], old_address)

    def test_approval_rejects_a_missing_changed_header(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ", "Thường trú": "Địa chỉ đề xuất"})
        workbook = load_workbook(self.workbook_path)
        workbook.active.delete_cols(app.state["headers"]["Nghề nghiệp"])
        workbook.save(self.workbook_path)
        workbook.close()
        before = self.workbook_path.read_bytes()

        with self.assertRaisesRegex(ValueError, "Nghề nghiệp"):
            approve_request(pending["id"], reviewer="admin")

        self.assertEqual(self.workbook_path.read_bytes(), before)
        self.assertEqual(get_request_by_id(pending["id"])["status"], "pending")

    def test_legacy_requests_without_a_source_remain_rejectable_but_cannot_be_applied(self):
        pending = self.submit_update(**{"Nghề nghiệp": "Bác sĩ"})
        # Simulate the schema used before requests recorded their source.
        with sqlite3.connect(_audit_database_path()) as connection:
            columns = {row[1] for row in connection.execute("PRAGMA table_info(profile_edit_requests)")}
            for column in ("workbook", "sheet"):
                if column in columns:
                    connection.execute(f"ALTER TABLE profile_edit_requests DROP COLUMN {column}")
        before = self.workbook_path.read_bytes()

        with self.assertRaises(ValueError):
            approve_request(pending["id"], reviewer="admin")

        self.assertEqual(self.workbook_path.read_bytes(), before)
        rejected = reject_request(pending["id"], reviewer="admin", rejection_reason="Vui lòng gửi lại")
        self.assertEqual(rejected["status"], "rejected")

    def test_identity_critical_fields_definition(self):
        """Verify the definition of identity-critical fields."""
        expected = {
            "TÊN THƯỜNG DÙNG",
            "Tên khai sinh",
            "Ngày sinh",
            "Tháng sinh",
            "Năm sinh",
            "Căn cước",
        }
        self.assertEqual(IDENTITY_CRITICAL_HEADERS, expected)
        for h in expected:
            self.assertTrue(is_identity_critical(h))
        self.assertFalse(is_identity_critical("Thường trú"))
        self.assertFalse(is_identity_critical("Nghề nghiệp"))

    def test_lookup_and_render_editable_profile(self):
        """On /tra-cuu, viewing profile shows identity-critical fields as readonly and general fields as editable."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        post_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        html = post_res.get_data(as_text=True)

        self.assertIn("NGUYỄN VĂN AN", html)
        self.assertIn("001205000001", html)
        # Identity-critical fields are marked with badge and locked
        self.assertIn("Trường định danh", html)
        self.assertIn("locked-field", html)
        self.assertIn("Gửi yêu cầu cập nhật", html)

    def test_submit_profile_edit_generates_pending_request_and_does_not_modify_db(self):
        """Submitting an edit generates a pending request and does NOT immediately modify the database."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))

        # Lookup
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        lookup_html = lookup_res.get_data(as_text=True)
        form_csrf = self.get_csrf(lookup_html)

        # Get column names
        occ_col = app.state["headers"]["Nghề nghiệp"]
        addr_col = app.state["headers"]["Thường trú"]

        # Submit general profile edits
        edit_data = {
            "csrf_token": form_csrf,
            "cccd": "001205000001",
            f"column_{occ_col}": "Kỹ sư phần mềm",
            f"column_{addr_col}": "Số 99 Đường Mặt Trời, KP 5",
        }
        submit_res = self.client.post("/tra-cuu/cap-nhat", data=edit_data)
        submit_html = submit_res.get_data(as_text=True)

        self.assertIn("Yêu cầu cập nhật thông tin đã được gửi", submit_html)

        # Check DB was NOT modified directly
        rec = find_record_by_citizen_id("001205000001")
        self.assertEqual(rec["data"]["Nghề nghiệp"], "Học sinh")
        self.assertEqual(rec["data"]["Thường trú"], "Số 10 Đường Hoa, KP 1")

        # Check pending request was created
        self.assertTrue(has_pending_request("001205000001"))
        self.assertEqual(count_pending_requests(), 1)

        pending = get_pending_request_by_cccd("001205000001")
        self.assertIsNotNone(pending)
        self.assertEqual(pending["status"], "pending")
        self.assertEqual(pending["cccd"], "001205000001")

        # Diff View contains only modified fields
        changed_fields = {c["field"]: (c["from"], c["to"]) for c in pending["changes"]}
        self.assertEqual(len(changed_fields), 2)
        self.assertEqual(changed_fields["Nghề nghiệp"], ("Học sinh", "Kỹ sư phần mềm"))
        self.assertEqual(changed_fields["Thường trú"], ("Số 10 Đường Hoa, KP 1", "Số 99 Đường Mặt Trời, KP 5"))

    def test_strictly_prohibit_user_changes_to_identity_critical_fields(self):
        """Even if submitted form payload tampers with identity fields, they must be strictly preserved."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        form_csrf = self.get_csrf(lookup_res.get_data(as_text=True))

        name_col = app.state["headers"]["TÊN THƯỜNG DÙNG"]
        dob_col = app.state["headers"]["Năm sinh"]
        occ_col = app.state["headers"]["Nghề nghiệp"]

        # Attempt to tamper with name and birth year
        edit_data = {
            "csrf_token": form_csrf,
            "cccd": "001205000001",
            f"column_{name_col}": "HACKER TRẦN",
            f"column_{dob_col}": "1990",
            f"column_{occ_col}": "Chuyên viên IT",
        }
        self.client.post("/tra-cuu/cap-nhat", data=edit_data)

        pending = get_pending_request_by_cccd("001205000001")
        self.assertIsNotNone(pending)

        # Name and DOB changes MUST NOT be present in changes diff
        change_fields = [c["field"] for c in pending["changes"]]
        self.assertNotIn("TÊN THƯỜNG DÙNG", change_fields)
        self.assertNotIn("Năm sinh", change_fields)
        self.assertIn("Nghề nghiệp", change_fields)

        # The submitted data stored in the request must retain the original identity values
        self.assertEqual(pending["submitted_data"]["TÊN THƯỜNG DÙNG"], "NGUYỄN VĂN AN")
        self.assertEqual(str(pending["submitted_data"]["Năm sinh"]), "2005")

    def test_temporary_lockout_spam_prevention(self):
        """Once an update request is submitted, user is temporarily locked from submitting further updates."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        form_csrf = self.get_csrf(lookup_res.get_data(as_text=True))

        occ_col = app.state["headers"]["Nghề nghiệp"]
        edit_data = {
            "csrf_token": form_csrf,
            "cccd": "001205000001",
            f"column_{occ_col}": "Sinh viên Y",
        }
        # First submission succeeds
        self.client.post("/tra-cuu/cap-nhat", data=edit_data)
        self.assertTrue(has_pending_request("001205000001"))

        # Lookup page now displays lockout notice and pending diff
        lookup2 = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        html2 = lookup2.get_data(as_text=True)
        self.assertIn("Yêu cầu chỉnh sửa đang chờ phê duyệt", html2)
        self.assertIn("chức năng gửi chỉnh sửa tạm thời bị khóa", html2.casefold())
        self.assertIn("Sinh viên Y", html2)

        # Attempting second submission while locked is strictly prevented
        form_csrf2 = self.get_csrf(html2)
        second_submit = self.client.post(
            "/tra-cuu/cap-nhat",
            data={
                "csrf_token": form_csrf2,
                "cccd": "001205000001",
                f"column_{occ_col}": "Sinh viên Luật",
            },
        )
        second_html = second_submit.get_data(as_text=True)
        self.assertIn("Hồ sơ đang có một yêu cầu cập nhật chờ xét duyệt", second_html)

        # Only 1 pending request exists
        self.assertEqual(count_pending_requests(), 1)

    def test_approval_workflow_and_unlock_condition(self):
        """Approving request modifies the record, unlocks the user, and records audit log."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        form_csrf = self.get_csrf(lookup_res.get_data(as_text=True))

        occ_col = app.state["headers"]["Nghề nghiệp"]
        self.client.post(
            "/tra-cuu/cap-nhat",
            data={
                "csrf_token": form_csrf,
                "cccd": "001205000001",
                f"column_{occ_col}": "Bác sĩ",
            },
        )
        self.assertTrue(has_pending_request("001205000001"))
        pending = get_pending_request_by_cccd("001205000001")

        # Approve request
        approved = approve_request(pending["id"], reviewer="admin", admin_note="Đã xác thực giấy tờ")
        self.assertEqual(approved["status"], "approved")
        self.assertEqual(approved["reviewer"], "admin")
        self.assertEqual(approved["reviewer_note"], "Đã xác thực giấy tờ")

        # Record is updated in database
        rec = find_record_by_citizen_id("001205000001")
        self.assertEqual(rec["data"]["Nghề nghiệp"], "Bác sĩ")

        # User is automatically UNLOCKED!
        self.assertFalse(has_pending_request("001205000001"))
        self.assertEqual(count_pending_requests(), 0)

        # User can now submit another update request
        lookup_after = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        html_after = lookup_after.get_data(as_text=True)
        self.assertIn("Yêu cầu cập nhật gần nhất của bạn đã được phê duyệt", html_after)
        self.assertIn("Gửi yêu cầu cập nhật", html_after)

    def test_rejection_workflow_and_unlock_condition(self):
        """Rejecting request leaves record unchanged, unlocks the user, and records audit log."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        form_csrf = self.get_csrf(lookup_res.get_data(as_text=True))

        occ_col = app.state["headers"]["Nghề nghiệp"]
        self.client.post(
            "/tra-cuu/cap-nhat",
            data={
                "csrf_token": form_csrf,
                "cccd": "001205000001",
                f"column_{occ_col}": "Cầu thủ bóng đá",
            },
        )
        self.assertTrue(has_pending_request("001205000001"))
        pending = get_pending_request_by_cccd("001205000001")

        # Reject request
        rejected = reject_request(pending["id"], reviewer="admin", rejection_reason="Thông tin không chính xác")
        self.assertEqual(rejected["status"], "rejected")
        self.assertEqual(rejected["reviewer"], "admin")
        self.assertEqual(rejected["reviewer_note"], "Thông tin không chính xác")

        # Record is UNCHANGED in database
        rec = find_record_by_citizen_id("001205000001")
        self.assertEqual(rec["data"]["Nghề nghiệp"], "Học sinh")

        # User is automatically UNLOCKED!
        self.assertFalse(has_pending_request("001205000001"))
        self.assertEqual(count_pending_requests(), 0)

        # User sees rejection reason on lookup and can submit a new request
        lookup_after = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        html_after = lookup_after.get_data(as_text=True)
        self.assertIn("Yêu cầu cập nhật gần nhất đã bị từ chối", html_after)
        self.assertIn("Thông tin không chính xác", html_after)
        self.assertIn("Gửi yêu cầu cập nhật", html_after)

    def test_immutable_audit_logging_lifecycle(self):
        """Full lifecycle audit logging: submission, approval, rejection."""
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        form_csrf = self.get_csrf(lookup_res.get_data(as_text=True))

        occ_col = app.state["headers"]["Nghề nghiệp"]
        self.client.post(
            "/tra-cuu/cap-nhat",
            data={
                "csrf_token": form_csrf,
                "cccd": "001205000001",
                f"column_{occ_col}": "Nhà văn",
            },
        )
        pending = get_pending_request_by_cccd("001205000001")

        # Check submission audit log
        logs, total = get_audit_logs(action="request_update")
        self.assertEqual(total, 1)
        sub_log = logs[0]
        self.assertEqual(sub_log["action"], "request_update")
        self.assertIn("001205000001", sub_log["username"])
        self.assertEqual(sub_log["details"]["lifecycle"], "submitted")
        self.assertEqual(sub_log["details"]["changes"][0]["field"], "Nghề nghiệp")
        self.assertEqual(sub_log["details"]["changes"][0]["to"], "Nhà văn")

        # Approve and check approval audit log
        approve_request(pending["id"], reviewer="admin_reviewer", admin_note="Phê duyệt hợp lệ")

        app_logs, app_total = get_audit_logs(action="approve_update")
        self.assertEqual(app_total, 1)
        app_log = app_logs[0]
        self.assertEqual(app_log["action"], "approve_update")
        self.assertEqual(app_log["username"], "admin_reviewer")
        self.assertEqual(app_log["details"]["lifecycle"], "approved")
        self.assertEqual(app_log["details"]["admin_note"], "Phê duyệt hợp lệ")
        self.assertEqual(app_log["details"]["reviewer"], "admin_reviewer")

    def test_admin_review_ui_access_and_diff_view(self):
        """Admin review UI displays list and diff view, accessible by admin or user with update permission."""
        # Create an account with update permission
        create_user("moderator", "123456", can_read=True, can_update=True)

        # Create a pending request
        res = self.client.get("/tra-cuu")
        csrf = self.get_csrf(res.get_data(as_text=True))
        lookup_res = self.client.post("/tra-cuu", data={"csrf_token": csrf, "cccd": "001205000001"})
        form_csrf = self.get_csrf(lookup_res.get_data(as_text=True))
        occ_col = app.state["headers"]["Nghề nghiệp"]
        self.client.post(
            "/tra-cuu/cap-nhat",
            data={
                "csrf_token": form_csrf,
                "cccd": "001205000001",
                f"column_{occ_col}": "Kiến trúc sư",
            },
        )

        # Admin logs in
        login_res = self.client.get("/login")
        login_csrf = self.get_csrf(login_res.get_data(as_text=True))
        self.client.post("/login", data={"csrf_token": login_csrf, "username": "admin", "password": "admin123"})

        # Admin views review queue
        queue_res = self.client.get("/admin/requests")
        self.assertEqual(queue_res.status_code, 200)
        html = queue_res.get_data(as_text=True)

        # Verify Diff View is present
        self.assertIn("So sánh thay đổi (Diff View)", html)
        self.assertIn("Nghề nghiệp", html)
        self.assertIn("Học sinh", html)
        self.assertIn("Kiến trúc sư", html)
        self.assertIn("Phê duyệt", html)
        self.assertIn("Từ chối", html)

        # Admin performs review via POST endpoint
        req = get_pending_request_by_cccd("001205000001")
        queue_csrf = self.get_csrf(html)
        rev_res = self.client.post(
            f"/admin/requests/{req['id']}/review",
            data={
                "csrf_token": queue_csrf,
                "action": "approve",
                "note": "Hồ sơ đầy đủ điều kiện",
            },
            follow_redirects=True,
        )
        self.assertEqual(rev_res.status_code, 200)
        self.assertIn("Đã PHÊ DUYỆT yêu cầu", rev_res.get_data(as_text=True))

        # Check record in workbook updated
        rec = find_record_by_citizen_id("001205000001")
        self.assertEqual(rec["data"]["Nghề nghiệp"], "Kiến trúc sư")

import tests
import os
import re
import tempfile
import unittest
from pathlib import Path

from flask import Flask
from openpyxl import Workbook, load_workbook

import quan_ly_ho_so.app as app
from quan_ly_ho_so.config import BASE_DIR
from quan_ly_ho_so.auth import (
    DEFAULT_ADMIN_PASSWORD,
    DEFAULT_ADMIN_USERNAME,
    auth_database_connection,
    auth_database_path,
    auth_required,
    authenticate,
    create_user,
    current_user,
    delete_user,
    get_persistent_secret_key,
    get_user_by_username,
    list_users,
    set_auth_database_path,
    update_user,
)
from tests.test_quan_ly_ho_so import create_test_workbook


class AuthAndPermissionsTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.workbook_path = Path(self.temp_dir.name) / "records.xlsx"
        self.previous_database = app.state["database"]
        # Point cache and auth DB to temporary directory
        app.state["database"] = Path(self.temp_dir.name) / "cache.sqlite3"
        set_auth_database_path(Path(self.temp_dir.name) / "users.db")
        create_test_workbook(self.workbook_path)
        app.set_workbook(self.workbook_path, persist=False)
        # Ensure auth DB is initialized with default admin in the temp directory
        conn = auth_database_connection()
        conn.close()

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

    def _login(self, client, username, password):
        login_page = client.get("/login")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', login_page.get_data(as_text=True)).group(1)
        return client.post(
            "/login",
            data={"username": username, "password": password, "csrf_token": csrf},
            follow_redirects=True,
        )

    def test_default_single_admin_initialization(self):
        users = list_users()
        self.assertEqual(len(users), 1)
        admin = users[0]
        self.assertEqual(admin["username"], DEFAULT_ADMIN_USERNAME)
        self.assertEqual(admin["role"], "admin")
        self.assertEqual(admin["can_create"], 1)
        self.assertEqual(admin["can_read"], 1)
        self.assertEqual(admin["can_update"], 1)
        self.assertEqual(admin["can_delete"], 1)

        # Cannot delete the single admin
        with self.assertRaises(ValueError):
            delete_user(admin["id"])

        # Authenticate with default password
        user = authenticate(DEFAULT_ADMIN_USERNAME, DEFAULT_ADMIN_PASSWORD)
        self.assertIsNotNone(user)
        self.assertEqual(user["username"], "admin")

    def test_create_and_manage_users(self):
        # Admin creates regular user with C and R permissions only
        user1 = create_user("editor_user", "pass1234", can_create=True, can_read=True, can_update=False, can_delete=False)
        self.assertEqual(user1["username"], "editor_user")
        self.assertEqual(user1["role"], "user")
        self.assertEqual(user1["can_create"], 1)
        self.assertEqual(user1["can_read"], 1)
        self.assertEqual(user1["can_update"], 0)
        self.assertEqual(user1["can_delete"], 0)

        # Duplicate username is rejected
        with self.assertRaises(ValueError):
            create_user("editor_user", "pass9999")

        # Update user permissions to grant Update
        updated = update_user(user1["id"], can_update=True)
        self.assertEqual(updated["can_update"], 1)

        # Delete user
        delete_user(user1["id"])
        self.assertIsNone(get_user_by_username("editor_user"))

    def test_authentication_flow_and_session_caching(self):
        client = app.app.test_client()

        # Unauthenticated request redirects to /login
        resp = client.get("/")
        self.assertEqual(resp.status_code, 302)
        self.assertIn("/login", resp.headers.get("Location", ""))

        # Invalid login fails with error message
        login_page = client.get("/login")
        self.assertEqual(login_page.status_code, 200)
        self.assertIn("Đăng nhập vào hệ thống", login_page.get_data(as_text=True))
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', login_page.get_data(as_text=True)).group(1)

        fail_resp = client.post(
            "/login",
            data={"username": "admin", "password": "wrongpassword", "csrf_token": csrf},
            follow_redirects=True,
        )
        self.assertIn("Tên đăng nhập hoặc mật khẩu không chính xác.", fail_resp.get_data(as_text=True))

        # Valid login succeeds and stores session
        success_resp = client.post(
            "/login",
            data={"username": "admin", "password": "admin123", "csrf_token": csrf},
            follow_redirects=True,
        )
        self.assertEqual(success_resp.status_code, 200)
        self.assertIn("Đăng nhập thành công", success_resp.get_data(as_text=True))

        # Verify session has cached permissions and user info
        with client.session_transaction() as sess:
            self.assertEqual(sess.get("username"), "admin")
            self.assertEqual(sess.get("role"), "admin")
            self.assertTrue(sess.get("permissions")["create"])
            self.assertTrue(sess.get("permissions")["read"])
            self.assertTrue(sess.get("permissions")["update"])
            self.assertTrue(sess.get("permissions")["delete"])

        # Logout clears session
        logout_resp = client.get("/logout", follow_redirects=True)
        self.assertIn("Đã đăng xuất khỏi hệ thống", logout_resp.get_data(as_text=True))
        with client.session_transaction() as sess:
            self.assertIsNone(sess.get("user_id"))

    def test_admin_page_authorization(self):
        # Create a regular user
        create_user("normal_user", "pass1234", can_read=True, must_change_password=False)

        client = app.app.test_client()
        # Normal user cannot access /admin
        self._login(client, "normal_user", "pass1234")
        resp = client.get("/admin", follow_redirects=True)
        self.assertIn("Chỉ Quản trị viên mới có quyền", resp.get_data(as_text=True))

        client.get("/logout")

        # Admin can access /admin
        self._login(client, "admin", "admin123")
        admin_resp = client.get("/admin")
        self.assertEqual(admin_resp.status_code, 200)
        admin_html = admin_resp.get_data(as_text=True)
        self.assertIn("Quản trị người dùng & Phân quyền", admin_html)
        self.assertIn("Thêm người dùng mới", admin_html)
        self.assertIn("normal_user", admin_html)

        # Verify UI display changes
        self.assertIn(">Trạng thái</th>", admin_html)
        self.assertNotIn("Trạng thái MK", admin_html)
        self.assertNotIn("Thao tác</th>", admin_html)
        self.assertNotIn("Quyền dữ liệu (CRUD)", admin_html)
        self.assertIn("checkbox-group vertical", admin_html)
        self.assertIn('class="user-row"', admin_html)
        self.assertIn('openUserModal', admin_html)

        # Form for new user has no password input, shows details, and only shows Tạo, Xem, Sửa, Xóa
        create_form_snippet = admin_html.split('action="/admin/users/create"')[1].split('</form>')[0]
        self.assertNotIn('type="password"', create_form_snippet)
        self.assertNotIn('name="password"', create_form_snippet)
        self.assertIn(">Tạo</span>", create_form_snippet)
        self.assertIn(">Xem</span>", create_form_snippet)
        self.assertIn(">Sửa</span>", create_form_snippet)
        self.assertIn(">Xóa</span>", create_form_snippet)
        self.assertIn("Thêm mới hồ sơ lý lịch thanh niên", create_form_snippet)
        self.assertIn("Tra cứu, tìm kiếm, lọc danh sách hồ sơ", create_form_snippet)
        self.assertNotIn("<strong>C</strong>", create_form_snippet)
        self.assertNotIn("<strong>R</strong>", create_form_snippet)
        self.assertNotIn("<strong>U</strong>", create_form_snippet)
        self.assertNotIn("<strong>D</strong>", create_form_snippet)

        # Admin cannot set custom new password (no password input in overlay modals)
        self.assertNotIn('Mật khẩu chỉ định', admin_html)
        self.assertNotIn('name="password"', admin_html)

        # User overlays exist
        self.assertIn('class="modal-overlay"', admin_html)

    def test_admin_crud_users_via_ui_routes(self):
        client = app.app.test_client()
        self._login(client, "admin", "admin123")

        admin_page = client.get("/admin")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', admin_page.get_data(as_text=True)).group(1)

        # Create user via UI (password always defaults to 123456)
        create_resp = client.post(
            "/admin/users/create",
            data={
                "username": "staff1",
                "can_create": "1",
                "can_read": "1",
                "csrf_token": csrf,
            },
            follow_redirects=True,
        )
        self.assertIn("Đã tạo người dùng &#39;staff1&#39; thành công", create_resp.get_data(as_text=True))
        staff = get_user_by_username("staff1")
        self.assertIsNotNone(staff)
        self.assertEqual(staff["can_create"], 1)
        self.assertEqual(staff["can_read"], 1)
        self.assertEqual(staff["can_update"], 0)
        self.assertEqual(staff["can_delete"], 0)
        self.assertIsNotNone(authenticate("staff1", "123456"))

        # Create user via UI without password (new UI behavior, defaults to 123456)
        create_resp2 = client.post(
            "/admin/users/create",
            data={
                "username": "staff_default_pw",
                "can_read": "1",
                "csrf_token": csrf,
            },
            follow_redirects=True,
        )
        self.assertIn("Đã tạo người dùng &#39;staff_default_pw&#39; thành công", create_resp2.get_data(as_text=True))
        staff_default = authenticate("staff_default_pw", "123456")
        self.assertIsNotNone(staff_default)
        self.assertEqual(staff_default["username"], "staff_default_pw")

        # Update user permissions via UI (attempting to set custom password is ignored)
        update_resp = client.post(
            f"/admin/users/{staff['id']}/update",
            data={
                "password": "attempt_custom_pw",
                "can_create": "1",
                "can_read": "1",
                "can_update": "1",
                "can_delete": "1",
                "csrf_token": csrf,
            },
            follow_redirects=True,
        )
        self.assertIn("Đã cập nhật quyền người dùng", update_resp.get_data(as_text=True))
        staff_updated = get_user_by_username("staff1")
        self.assertEqual(staff_updated["can_update"], 1)
        self.assertEqual(staff_updated["can_delete"], 1)
        # Verify custom password was NOT accepted
        self.assertIsNone(authenticate("staff1", "attempt_custom_pw"))
        self.assertIsNotNone(authenticate("staff1", "123456"))

        # Delete user via UI
        del_resp = client.post(
            f"/admin/users/{staff['id']}/delete",
            data={"csrf_token": csrf},
            follow_redirects=True,
        )
        self.assertIn("Đã xóa người dùng thành công", del_resp.get_data(as_text=True))
        self.assertIsNone(get_user_by_username("staff1"))

    def test_permission_read_enforcement(self):
        # User without Read permission
        create_user("no_read_user", "pass1234", can_read=False, can_create=False, must_change_password=False)

        client = app.app.test_client()
        self._login(client, "no_read_user", "pass1234")

        # Home page shows notice that user cannot view records
        resp = client.get("/")
        self.assertEqual(resp.status_code, 200)
        html = resp.get_data(as_text=True)
        self.assertIn("Không có quyền xem dữ liệu", html)
        self.assertNotIn("PHAN TRẦN GIA HUY", html)

        # Word export routes blocked
        with client.session_transaction() as sess:
            csrf = sess.get("csrf_token")
        gen_resp = client.post("/generate/1", data={"csrf_token": csrf}, follow_redirects=True)
        self.assertIn("Bạn không có quyền xem dữ liệu hồ sơ", gen_resp.get_data(as_text=True))

    def test_permission_create_enforcement(self):
        # User with Read only (no Create)
        create_user("reader_only", "pass1234", can_read=True, can_create=False, must_change_password=False)

        client = app.app.test_client()
        self._login(client, "reader_only", "pass1234")

        # Accessing /person/new is denied
        resp = client.get("/person/new", follow_redirects=True)
        self.assertIn("Bạn không có quyền thêm mới dữ liệu hồ sơ", resp.get_data(as_text=True))

        # "Thêm hồ sơ" button is not in UI
        home_html = client.get("/").get_data(as_text=True)
        self.assertNotIn('<a class="button" href="/person/new">Thêm hồ sơ</a>', home_html)

    def test_permission_update_enforcement(self):
        # User with Read only (no Update)
        create_user("reader_only", "pass1234", can_read=True, can_update=False, must_change_password=False)

        client = app.app.test_client()
        self._login(client, "reader_only", "pass1234")

        # Accessing /person/1/edit is denied
        resp = client.get("/person/1/edit", follow_redirects=True)
        self.assertIn("Bạn không có quyền chỉnh sửa dữ liệu hồ sơ", resp.get_data(as_text=True))

        # "Sửa" link is not shown in records table for this user
        home_html = client.get("/").get_data(as_text=True)
        self.assertNotIn('href="/person/1/edit">Sửa</a>', home_html)

    def test_permission_delete_enforcement_and_deletion(self):
        # User without delete permission
        create_user("no_delete_user", "pass1234", can_read=True, can_delete=False, must_change_password=False)

        client = app.app.test_client()
        self._login(client, "no_delete_user", "pass1234")

        # Home page does not display delete button for this user
        home_html = client.get("/").get_data(as_text=True)
        self.assertNotIn('class="button danger">Xóa</button>', home_html)

        csrf = re.search(r'name="csrf_token" value="([^"]+)"', home_html).group(1)
        del_attempt = client.post("/person/1/delete", data={"csrf_token": csrf}, follow_redirects=True)
        self.assertIn("Bạn không có quyền xóa dữ liệu hồ sơ", del_attempt.get_data(as_text=True))

        client.get("/logout")

        # Now test with user who has delete permission
        create_user("deleter_user", "pass1234", can_read=True, can_delete=True, must_change_password=False)
        self._login(client, "deleter_user", "pass1234")

        home_resp = client.get("/")
        home_html = home_resp.get_data(as_text=True)
        self.assertIn('class="button danger">Xóa</button>', home_html)

        csrf = re.search(r'name="csrf_token" value="([^"]+)"', home_html).group(1)
        sig = re.search(r'name="signature" value="([^"]*)"', home_html).group(1)

        # Record count before deletion
        self.assertEqual(app.state["record_count"], 3)

        del_success = client.post(
            "/person/1/delete",
            data={"csrf_token": csrf, "signature": sig},
            follow_redirects=True,
        )
        self.assertIn("Đã xóa hồ sơ STT 1", del_success.get_data(as_text=True))
        # Record count after deletion
        self.assertEqual(app.state["record_count"], 2)

    def test_default_user_creation_and_first_login_password_change(self):
        client = app.app.test_client()
        self._login(client, "admin", "admin123")

        admin_page = client.get("/admin")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', admin_page.get_data(as_text=True)).group(1)

        # Admin creates user with default password (leaving password blank)
        create_resp = client.post(
            "/admin/users/create",
            data={
                "username": "newbie",
                "password": "",
                "can_read": "1",
                "csrf_token": csrf,
            },
            follow_redirects=True,
        )
        self.assertIn("Đã tạo người dùng &#39;newbie&#39; thành công", create_resp.get_data(as_text=True))
        client.get("/logout")

        # User exists and must_change_password is 1
        newbie = get_user_by_username("newbie")
        self.assertIsNotNone(newbie)
        self.assertEqual(newbie["must_change_password"], 1)

        # User logs in with default password "123456"
        login_resp = self._login(client, "newbie", "123456")
        # System immediately redirects to /force-change-password
        self.assertIn("Yêu cầu đổi mật khẩu mới", login_resp.get_data(as_text=True))

        # Attempting to access other routes redirects back to /force-change-password
        index_attempt = client.get("/", follow_redirects=True)
        self.assertIn("Yêu cầu đổi mật khẩu mới", index_attempt.get_data(as_text=True))

        profile_attempt = client.get("/profile", follow_redirects=True)
        self.assertIn("Yêu cầu đổi mật khẩu mới", profile_attempt.get_data(as_text=True))

        # Trying to submit new password matching default '123456' is rejected
        force_page = client.get("/force-change-password")
        csrf_force = re.search(r'name="csrf_token" value="([^"]+)"', force_page.get_data(as_text=True)).group(1)

        same_pw_resp = client.post(
            "/force-change-password",
            data={"new_password": "123456", "confirm_password": "123456", "csrf_token": csrf_force},
            follow_redirects=True,
        )
        self.assertIn("Mật khẩu mới không được trùng với mật khẩu mặc định", same_pw_resp.get_data(as_text=True))

        # Trying mismatched password is rejected
        mismatch_resp = client.post(
            "/force-change-password",
            data={"new_password": "passNew123", "confirm_password": "otherPass", "csrf_token": csrf_force},
            follow_redirects=True,
        )
        self.assertIn("Mật khẩu mới và xác nhận mật khẩu không khớp nhau", mismatch_resp.get_data(as_text=True))

        # Submitting valid new password succeeds and grants access
        success_resp = client.post(
            "/force-change-password",
            data={"new_password": "passNew123", "confirm_password": "passNew123", "csrf_token": csrf_force},
            follow_redirects=True,
        )
        self.assertIn("Cài đặt mật khẩu mới thành công", success_resp.get_data(as_text=True))

        # Now can access index directly without redirection
        resp_after = client.get("/")
        self.assertEqual(resp_after.status_code, 200)
        self.assertIn("Tra cứu, cập nhật hồ sơ", resp_after.get_data(as_text=True))

        # Check DB flag
        newbie_updated = get_user_by_username("newbie")
        self.assertEqual(newbie_updated["must_change_password"], 0)

        # Logout and log in with new password works
        client.get("/logout")
        login_new = self._login(client, "newbie", "passNew123")
        self.assertEqual(login_new.status_code, 200)
        self.assertIn("Đăng nhập thành công", login_new.get_data(as_text=True))

    def test_admin_reset_user_password_and_forced_change(self):
        # Create user with custom password and must_change_password=0
        user = create_user("forgetful_user", "oldsecret99", can_read=True, must_change_password=False)
        self.assertEqual(user["must_change_password"], 0)

        client = app.app.test_client()
        self._login(client, "admin", "admin123")

        admin_page = client.get("/admin")
        csrf = re.search(r'name="csrf_token" value="([^"]+)"', admin_page.get_data(as_text=True)).group(1)

        # Admin resets password for user
        reset_resp = client.post(f"/admin/users/{user['id']}/reset-password", data={"csrf_token": csrf}, follow_redirects=True)
        self.assertIn("Đã đặt lại mật khẩu cho người dùng", reset_resp.get_data(as_text=True))

        # DB has default password and must_change_password = 1
        user_reset = get_user_by_username("forgetful_user")
        self.assertEqual(user_reset["must_change_password"], 1)

        client.get("/logout")

        # User logs in with default password "123456"
        login_resp = self._login(client, "forgetful_user", "123456")
        self.assertIn("Yêu cầu đổi mật khẩu mới", login_resp.get_data(as_text=True))

        # User changes password
        force_page = client.get("/force-change-password")
        csrf_force = re.search(r'name="csrf_token" value="([^"]+)"', force_page.get_data(as_text=True)).group(1)
        client.post(
            "/force-change-password",
            data={"new_password": "newFreshPassword123", "confirm_password": "newFreshPassword123", "csrf_token": csrf_force},
            follow_redirects=True,
        )

        user_final = get_user_by_username("forgetful_user")
        self.assertEqual(user_final["must_change_password"], 0)

    def test_user_profile_and_voluntary_password_change(self):
        create_user(
            "profile_user",
            "mypassword1",
            can_create=True,
            can_read=True,
            can_update=False,
            can_delete=False,
            must_change_password=False,
        )

        client = app.app.test_client()
        self._login(client, "profile_user", "mypassword1")

        # Visit profile page
        profile_resp = client.get("/profile")
        self.assertEqual(profile_resp.status_code, 200)
        profile_html = profile_resp.get_data(as_text=True)
        self.assertIn("Quyền hạn của bản thân", profile_html)
        self.assertIn("profile_user", profile_html)
        self.assertIn("Quyền Xem", profile_html)
        self.assertIn("Quyền Tạo", profile_html)
        self.assertIn("Quyền Sửa", profile_html)
        self.assertIn("Quyền Xóa", profile_html)

        csrf_profile = re.search(r'name="csrf_token" value="([^"]+)"', profile_html).group(1)

        # Wrong current password fails
        bad_cur_resp = client.post(
            "/profile/change-password",
            data={
                "current_password": "wrongCurrentPassword",
                "new_password": "newPassword777",
                "confirm_password": "newPassword777",
                "csrf_token": csrf_profile,
            },
            follow_redirects=True,
        )
        self.assertIn("Mật khẩu hiện tại không chính xác", bad_cur_resp.get_data(as_text=True))

        # Setting to default password "123456" is rejected
        default_resp = client.post(
            "/profile/change-password",
            data={
                "current_password": "mypassword1",
                "new_password": "123456",
                "confirm_password": "123456",
                "csrf_token": csrf_profile,
            },
            follow_redirects=True,
        )
        self.assertIn("Mật khẩu mới không được trùng với mật khẩu mặc định", default_resp.get_data(as_text=True))

        # Valid password change succeeds
        ok_resp = client.post(
            "/profile/change-password",
            data={
                "current_password": "mypassword1",
                "new_password": "newPassword777",
                "confirm_password": "newPassword777",
                "csrf_token": csrf_profile,
            },
            follow_redirects=True,
        )
        self.assertIn("Đổi mật khẩu thành công", ok_resp.get_data(as_text=True))

        # Logout and log in with new password
        client.get("/logout")
        new_login = self._login(client, "profile_user", "newPassword777")
        self.assertEqual(new_login.status_code, 200)
        self.assertIn("Đăng nhập thành công", new_login.get_data(as_text=True))

    def test_auth_database_path_in_main_project_directory_and_not_excel_dependent(self):
        # Clear test override temporarily to verify production default behavior
        set_auth_database_path(None)
        try:
            # Without AUTH_DB_PATH, default is BASE_DIR / "users.db"
            old_env = os.environ.pop("AUTH_DB_PATH", None)
            default_path = auth_database_path()
            self.assertEqual(default_path, (BASE_DIR / "users.db").resolve())

            # Changing the workbook cache database in state does NOT affect auth database path
            app.state["database"] = Path("/some/custom/path/cache.sqlite3")
            self.assertEqual(auth_database_path(), (BASE_DIR / "users.db").resolve())

            # Setting AUTH_DB_PATH in env overrides the database path at project root
            os.environ["AUTH_DB_PATH"] = "custom_users.db"
            self.assertEqual(auth_database_path(), (BASE_DIR / "custom_users.db").resolve())
        finally:
            if old_env is not None:
                os.environ["AUTH_DB_PATH"] = old_env
            else:
                os.environ.pop("AUTH_DB_PATH", None)
            # Restore isolated test override
            set_auth_database_path(Path(self.temp_dir.name) / "users.db")

    def test_secret_key_and_env_constants(self):
        secret = get_persistent_secret_key()
        self.assertTrue(bool(secret))
        self.assertGreaterEqual(len(secret), 32)

    def test_auth_required_decorator_flexible_permissions_and_api(self):
        # Create an isolated Flask app specifically to test auth_required combinations
        dummy_app = Flask("dummy_decorator_test")
        dummy_app.secret_key = "test_decorator_secret"

        @dummy_app.get("/login")
        def login():
            return "LOGIN_PAGE"

        @dummy_app.get("/")
        def index():
            return "INDEX_PAGE"

        @dummy_app.get("/test/any-permission")
        @auth_required(permissions=["create", "read"], require_all=False)
        def any_perm_endpoint():
            return "SUCCESS_ANY"

        @dummy_app.get("/test/all-permissions")
        @auth_required(permissions=["read", "create"], require_all=True)
        def all_perm_endpoint():
            return "SUCCESS_ALL"

        @dummy_app.get("/api/test/json-protected")
        @auth_required(permission="create")
        def api_protected_endpoint():
            return {"status": "ok"}

        client = dummy_app.test_client()

        # 1. Unauthenticated JSON API call gets 401 JSON
        api_unauth = client.get("/api/test/json-protected")
        self.assertEqual(api_unauth.status_code, 401)
        self.assertEqual(api_unauth.get_json()["error"], "Unauthorized")

        # 2. Unauthenticated web call redirects to /login
        web_unauth = client.get("/test/any-permission")
        self.assertEqual(web_unauth.status_code, 302)
        self.assertIn("/login", web_unauth.headers.get("Location", ""))

        # 3. Authenticate session with 'read' only (no 'create')
        with client.session_transaction() as sess:
            sess["user_id"] = 99
            sess["username"] = "perm_tester"
            sess["role"] = "user"
            sess["permissions"] = {"create": False, "read": True, "update": False, "delete": False}

        # 4. User has 'read', so any-permission (create OR read) succeeds
        resp_any = client.get("/test/any-permission")
        self.assertEqual(resp_any.status_code, 200)
        self.assertEqual(resp_any.get_data(as_text=True), "SUCCESS_ANY")

        # 5. User lacks 'create', so all-permissions (read AND create) fails with redirect
        resp_all = client.get("/test/all-permissions", follow_redirects=True)
        self.assertIn("INDEX_PAGE", resp_all.get_data(as_text=True))

        # 6. Authenticated JSON API call lacking 'create' gets 403 JSON
        api_denied = client.get("/api/test/json-protected")
        self.assertEqual(api_denied.status_code, 403)
        self.assertEqual(api_denied.get_json()["error"], "Forbidden")


if __name__ == "__main__":
    unittest.main()


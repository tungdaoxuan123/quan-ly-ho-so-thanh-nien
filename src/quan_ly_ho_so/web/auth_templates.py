"""HTML templates for authentication, user management, and user profiles."""

LOGIN_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Đăng nhập · Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { min-height: 100vh; margin: 0; display: flex; align-items: center; justify-content: center; padding: 20px; box-sizing: border-box; }
    .login-card { width: 100%; max-width: 440px; background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 34px 30px; box-shadow: 0 4px 20px rgba(23,32,51,0.08); }
    .brand { text-align: center; margin-bottom: 24px; }
    .brand h1 { color: #163d68; font-size: 22px; margin: 0 0 6px; }
    .brand p { color: #637089; font-size: 14px; margin: 0; }
    label { display: flex; flex-direction: column; gap: 6px; font-weight: 600; font-size: 14px; margin-bottom: 16px; }
    input[type="text"], input[type="password"] { box-sizing: border-box; font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 11px 12px; width: 100%; background: #fff; }
    input:focus { outline: none; border-color: #1e5b91; box-shadow: 0 0 0 3px rgba(30,91,145,0.15); }
    .toggle-pw-row { display: flex; align-items: center; gap: 8px; margin-top: -6px; margin-bottom: 16px; font-size: 13px; color: #4b5875; cursor: pointer; }
    .toggle-pw-row input[type="checkbox"] { width: 15px; height: 15px; accent-color: #1e5b91; cursor: pointer; }
    button[type="submit"] { width: 100%; background: #1e5b91; border: 0; color: #fff; font: inherit; font-weight: 600; font-size: 15px; cursor: pointer; padding: 12px; border-radius: 7px; margin-top: 4px; transition: background 0.15s ease; }
    button[type="submit"]:hover { background: #174873; }
    .notice { padding: 12px 14px; border-radius: 8px; background: #e9f4eb; color: #265b31; margin-bottom: 18px; font-size: 14px; }
    .notice.error { background: #fdecec; color: #8f2f2f; }
    .help-hint { margin-top: 24px; padding-top: 18px; border-top: 1px solid #eef2f7; font-size: 13px; color: #637089; line-height: 1.55; text-align: center; }
    .help-hint strong { color: #163d68; }
  </style>
</head>
<body>
  <div class="login-card">
    <div class="brand">
      <h1>Quản lý hồ sơ thanh niên</h1>
      <p>Đăng nhập vào hệ thống</p>
    </div>
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% for category, message in messages %}
        <div class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</div>
      {% endfor %}
    {% endwith %}
    <form action="{{ url_for('login_post') }}" method="post">
      <input type="hidden" name="csrf_token" value="{{ csrf }}">
      <input type="hidden" name="next" value="{{ next_url or '' }}">
      <label>
        Tên đăng nhập
        <input type="text" name="username" value="{{ username or '' }}" placeholder="Nhập tên đăng nhập" autofocus required autocomplete="username">
      </label>
      <label>
        Mật khẩu
        <input type="password" id="loginPassword" name="password" placeholder="Nhập mật khẩu" required autocomplete="current-password">
      </label>
      <div class="toggle-pw-row" onclick="togglePasswordVisibility('loginPassword')">
        <input type="checkbox" id="showPwCheckbox">
        <label for="showPwCheckbox" style="margin: 0; font-weight: normal; cursor: pointer;">Hiện mật khẩu</label>
      </div>
      <button type="submit">Đăng nhập</button>
    </form>
    <div class="help-hint">
      Tài khoản Quản trị viên mặc định: <strong>admin</strong> / <strong>admin123</strong><br>
      Tài khoản người dùng mới tạo hoặc sau reset: mật khẩu mặc định là <strong>123456</strong> (hệ thống sẽ yêu cầu cài mật khẩu mới khi đăng nhập).
    </div>
  </div>
  <script>
    function togglePasswordVisibility(fieldId) {
      const field = document.getElementById(fieldId);
      const cb = document.getElementById('showPwCheckbox');
      if (field) {
        field.type = cb.checked ? 'text' : 'password';
      }
    }
  </script>
</body>
</html>
"""

FORCE_CHANGE_PASSWORD_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Thiết lập mật khẩu mới · Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { min-height: 100vh; margin: 0; display: flex; align-items: center; justify-content: center; padding: 20px; box-sizing: border-box; }
    .card { width: 100%; max-width: 460px; background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 34px 30px; box-shadow: 0 4px 20px rgba(23,32,51,0.08); }
    .header { text-align: center; margin-bottom: 22px; }
    .header .icon { font-size: 38px; margin-bottom: 8px; }
    .header h1 { color: #163d68; font-size: 22px; margin: 0 0 8px; }
    .header p { color: #505f79; font-size: 14px; margin: 0; line-height: 1.5; }
    .badge-username { display: inline-block; background: #e0edff; color: #163d68; padding: 2px 8px; border-radius: 6px; font-weight: 600; }
    .alert-box { background: #fff8e6; border: 1px solid #ffeeba; border-radius: 8px; padding: 12px 14px; margin-bottom: 20px; font-size: 13.5px; color: #856404; line-height: 1.5; }
    label { display: flex; flex-direction: column; gap: 6px; font-weight: 600; font-size: 14px; margin-bottom: 16px; }
    input[type="password"], input[type="text"] { box-sizing: border-box; font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 11px 12px; width: 100%; background: #fff; }
    input:focus { outline: none; border-color: #1e5b91; box-shadow: 0 0 0 3px rgba(30,91,145,0.15); }
    .toggle-pw-row { display: flex; align-items: center; gap: 8px; margin-top: -6px; margin-bottom: 16px; font-size: 13px; color: #4b5875; cursor: pointer; }
    .toggle-pw-row input[type="checkbox"] { width: 15px; height: 15px; accent-color: #1e5b91; cursor: pointer; }
    button[type="submit"] { width: 100%; background: #1e5b91; border: 0; color: #fff; font: inherit; font-weight: 600; font-size: 15px; cursor: pointer; padding: 12px; border-radius: 7px; margin-top: 6px; transition: background 0.15s ease; }
    button[type="submit"]:hover { background: #174873; }
    .actions-footer { margin-top: 18px; text-align: center; }
    .actions-footer a { color: #637089; font-size: 13.5px; text-decoration: none; }
    .actions-footer a:hover { color: #163d68; text-decoration: underline; }
    .notice { padding: 12px 14px; border-radius: 8px; background: #e9f4eb; color: #265b31; margin-bottom: 18px; font-size: 14px; }
    .notice.error { background: #fdecec; color: #8f2f2f; }
    .rules-list { margin: 8px 0 0; padding-left: 18px; font-size: 12.5px; color: #637089; }
  </style>
</head>
<body>
  <div class="card">
    <div class="header">
      <div class="icon">🔒</div>
      <h1>Yêu cầu đổi mật khẩu mới</h1>
      <p>Tài khoản: <span class="badge-username">{{ username }}</span></p>
    </div>

    <div class="alert-box">
      <strong>Lưu ý bảo mật:</strong> Đây là lần đăng nhập đầu tiên hoặc mật khẩu của bạn vừa được Quản trị viên đặt lại về mặc định (<strong>123456</strong>). Để bảo vệ an toàn thông tin, vui lòng cài đặt mật khẩu mới của bạn trước khi tiếp tục.
    </div>

    {% with messages = get_flashed_messages(with_categories=true) %}
      {% for category, message in messages %}
        <div class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</div>
      {% endfor %}
    {% endwith %}

    <form action="{{ url_for('force_change_password_post') }}" method="post">
      <input type="hidden" name="csrf_token" value="{{ csrf }}">
      <label>
        Mật khẩu mới
        <input type="password" id="newPw" name="new_password" placeholder="Nhập mật khẩu mới" required minlength="4" autofocus autocomplete="new-password">
      </label>
      <label>
        Xác nhận mật khẩu mới
        <input type="password" id="confirmPw" name="confirm_password" placeholder="Nhập lại mật khẩu mới" required minlength="4" autocomplete="new-password">
      </label>
      <div class="toggle-pw-row">
        <input type="checkbox" id="showPwToggle" onchange="togglePasswords()">
        <label for="showPwToggle" style="margin: 0; font-weight: normal; cursor: pointer;">Hiện mật khẩu</label>
      </div>
      <ul class="rules-list">
        <li>Mật khẩu mới phải có ít nhất 4 ký tự.</li>
        <li>Không được đặt trùng với mật khẩu mặc định (123456).</li>
      </ul>
      <button type="submit">Lưu mật khẩu mới & Bắt đầu sử dụng</button>
    </form>

    <div class="actions-footer">
      <a href="{{ url_for('logout') }}">← Đăng xuất khỏi tài khoản</a>
    </div>
  </div>

  <script>
    function togglePasswords() {
      const isChecked = document.getElementById('showPwToggle').checked;
      const type = isChecked ? 'text' : 'password';
      document.getElementById('newPw').type = type;
      document.getElementById('confirmPw').type = type;
    }
  </script>
</body>
</html>
"""

PROFILE_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Tài khoản & Quyền hạn · Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 1200px; margin: 0 auto; padding: 28px; }
    h1 { margin: 0; color: #163d68; }
    h2 { color: #163d68; margin: 0 0 10px; font-size: 19px; }
    h3 { color: #163d68; margin: 0 0 8px; font-size: 16px; }
    p { line-height: 1.5; }
    .muted { color: #637089; font-size: 14px; }
    .card { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 22px; margin: 16px 0; box-shadow: 0 2px 10px rgba(23,32,51,0.06); }
    .app-header, .actions { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; justify-content: space-between; }
    .actions { justify-content: flex-start; }
    button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-block; padding: 10px 14px; border-radius: 7px; font: inherit; font-size: 14px; white-space: nowrap; font-weight: 500; }
    button.secondary, .button.secondary { color: #1e5b91; background: #e9f1f8; }
    .notice { padding: 12px 14px; border-radius: 8px; background: #e9f4eb; color: #265b31; margin-bottom: 16px; }
    .notice.error { background: #fdecec; color: #8f2f2f; }

    .profile-layout { display: grid; grid-template-columns: 1.2fr 0.8fr; gap: 20px; align-items: start; }
    @media (max-width: 900px) { .profile-layout { grid-template-columns: 1fr; } }

    .info-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
    .info-table td { padding: 9px 12px; border-bottom: 1px solid #eef2f7; font-size: 14px; }
    .info-table td.label-col { width: 160px; font-weight: 600; color: #4b5875; }

    .role-badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
    .role-badge.admin { background: #e0edff; color: #163d68; border: 1px solid #b8d5ff; }
    .role-badge.user { background: #eaf5ea; color: #23632f; border: 1px solid #c2e3c5; }

    .perm-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 12px; }
    @media (max-width: 600px) { .perm-grid { grid-template-columns: 1fr; } }
    .perm-item { border: 1px solid #e2e8f0; border-radius: 9px; padding: 14px; background: #f8fafc; display: flex; flex-direction: column; justify-content: space-between; gap: 8px; }
    .perm-item.has-perm { border-color: #c3e6cb; background: #f7fdf9; }
    .perm-header { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
    .perm-name { font-weight: 700; color: #163d68; font-size: 14.5px; }
    .perm-desc { font-size: 13px; color: #505f79; line-height: 1.4; }
    .status-tag { display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 700; }
    .status-tag.yes { background: #e8f5e9; color: #1b5e20; border: 1px solid #a5d6a7; }
    .status-tag.no { background: #f1f3f5; color: #868e96; border: 1px solid #dee2e6; }

    .admin-callout { background: #eaf2fb; border-left: 4px solid #1e5b91; padding: 12px 16px; border-radius: 0 8px 8px 0; margin-top: 14px; font-size: 13.5px; color: #163d68; line-height: 1.5; }
    .user-callout { background: #f4f6fa; border-left: 4px solid #94a3b8; padding: 12px 16px; border-radius: 0 8px 8px 0; margin-top: 14px; font-size: 13px; color: #475569; line-height: 1.5; }

    label { display: flex; flex-direction: column; gap: 6px; font-weight: 600; font-size: 14px; margin-bottom: 14px; }
    input[type="password"], input[type="text"] { box-sizing: border-box; font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 10px 12px; width: 100%; background: #fff; }
    input:focus { outline: none; border-color: #1e5b91; box-shadow: 0 0 0 3px rgba(30,91,145,0.15); }
    .toggle-pw-row { display: flex; align-items: center; gap: 8px; margin-top: -4px; margin-bottom: 16px; font-size: 13px; color: #4b5875; cursor: pointer; }
    .toggle-pw-row input[type="checkbox"] { width: 15px; height: 15px; accent-color: #1e5b91; cursor: pointer; }
  </style>
</head>
<body>
  <header class="app-header">
    <div>
      <h1>Tài khoản & Quyền hạn của bản thân</h1>
      <p class="muted">Xem thông tin tài khoản cá nhân, phạm vi quyền hạn và đổi mật khẩu</p>
    </div>
    <div class="actions">
      <a class="button secondary" href="{{ url_for('index') }}">← Về trang tra cứu hồ sơ</a>
      {% if current_user.is_admin %}
        <a class="button secondary" href="{{ url_for('admin_dashboard') }}">Quản trị người dùng</a>
      {% endif %}
      <a class="button secondary" href="{{ url_for('logout') }}">Đăng xuất</a>
    </div>
  </header>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <div class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</div>
    {% endfor %}
  {% endwith %}

  <div class="profile-layout">
    <!-- Left Column: User details & Permissions -->
    <div>
      <section class="card">
        <h2>Thông tin tài khoản</h2>
        <table class="info-table">
          <tr>
            <td class="label-col">Tên đăng nhập:</td>
            <td><strong>{{ user.username }}</strong></td>
          </tr>
          <tr>
            <td class="label-col">Vai trò hệ thống:</td>
            <td>
              {% if user.role == 'admin' %}
                <span class="role-badge admin">Quản trị viên (Admin)</span>
              {% else %}
                <span class="role-badge user">Người dùng (User)</span>
              {% endif %}
            </td>
          </tr>
          <tr>
            <td class="label-col">Trạng thái:</td>
            <td><span class="status-tag yes">✓ Đang hoạt động</span></td>
          </tr>
          <tr>
            <td class="label-col">Ngày tạo tài khoản:</td>
            <td class="muted">{{ user.created_at[:19].replace('T', ' ') if user.created_at else 'Chưa xác định' }}</td>
          </tr>
          <tr>
            <td class="label-col">Cập nhật gần nhất:</td>
            <td class="muted">{{ user.updated_at[:19].replace('T', ' ') if user.updated_at else 'Chưa xác định' }}</td>
          </tr>
        </table>

        <h2>Quyền hạn của bản thân</h2>
        <p class="muted" style="margin-bottom: 12px;">Dưới đây là các quyền thao tác trên hệ thống quản lý hồ sơ thanh niên được phân cho bạn:</p>

        <div class="perm-grid">
          <!-- Read -->
          <div class="perm-item {{ 'has-perm' if current_user.can_read else '' }}">
            <div class="perm-header">
              <span class="perm-name">👁️ Quyền Xem</span>
              {% if current_user.can_read %}
                <span class="status-tag yes">Được phép</span>
              {% else %}
                <span class="status-tag no">Không có quyền</span>
              {% endif %}
            </div>
            <div class="perm-desc">
              Tra cứu danh sách hồ sơ, bộ lọc nâng cao, xem chi tiết mẫu lý lịch, tạo và tải tệp Word cá nhân và tệp Word lô.
            </div>
          </div>

          <!-- Create -->
          <div class="perm-item {{ 'has-perm' if current_user.can_create else '' }}">
            <div class="perm-header">
              <span class="perm-name">➕ Quyền Tạo</span>
              {% if current_user.can_create %}
                <span class="status-tag yes">Được phép</span>
              {% else %}
                <span class="status-tag no">Không có quyền</span>
              {% endif %}
            </div>
            <div class="perm-desc">
              Thêm mới hồ sơ lý lịch thanh niên vào cơ sở dữ liệu và lưu tự động vào bảng tính Excel.
            </div>
          </div>

          <!-- Update -->
          <div class="perm-item {{ 'has-perm' if current_user.can_update else '' }}">
            <div class="perm-header">
              <span class="perm-name">✏️ Quyền Sửa</span>
              {% if current_user.can_update %}
                <span class="status-tag yes">Được phép</span>
              {% else %}
                <span class="status-tag no">Không có quyền</span>
              {% endif %}
            </div>
            <div class="perm-desc">
              Chỉnh sửa thông tin hồ sơ đã có trong hệ thống và lưu các thay đổi vào bảng tính Excel.
            </div>
          </div>

          <!-- Delete -->
          <div class="perm-item {{ 'has-perm' if current_user.can_delete else '' }}">
            <div class="perm-header">
              <span class="perm-name">🗑️ Quyền Xóa</span>
              {% if current_user.can_delete %}
                <span class="status-tag yes">Được phép</span>
              {% else %}
                <span class="status-tag no">Không có quyền</span>
              {% endif %}
            </div>
            <div class="perm-desc">
              Xóa hồ sơ thanh niên khỏi danh sách và Excel (hệ thống tự động sao lưu dự phòng trước khi xóa).
            </div>
          </div>
        </div>

        {% if current_user.is_admin %}
          <div class="admin-callout">
            <strong>★ Đặc quyền Quản trị viên:</strong> Bạn là Quản trị viên duy nhất của hệ thống, có toàn quyền quản trị tài khoản người dùng, phân quyền và toàn bộ quyền thao tác dữ liệu (Tạo, Xem, Sửa, Xóa).
          </div>
        {% else %}
          <div class="user-callout">
            <strong>ℹ️ Ghi chú:</strong> Quyền thao tác dữ liệu được thiết lập bởi Quản trị viên hệ thống. Nếu bạn cần bổ sung quyền để phục vụ công tác chuyên môn, vui lòng liên hệ Quản trị viên.
          </div>
        {% endif %}
      </section>
    </div>

    <!-- Right Column: Change Password -->
    <div>
      <section class="card">
        <h2>Đổi mật khẩu tài khoản</h2>
        <p class="muted">Thay đổi mật khẩu đăng nhập của bạn để bảo mật tài khoản cá nhân.</p>

        <form action="{{ url_for('profile_change_password') }}" method="post">
          <input type="hidden" name="csrf_token" value="{{ csrf }}">
          <label>
            Mật khẩu hiện tại
            <input type="password" id="curPw" name="current_password" required placeholder="Nhập mật khẩu hiện tại" autocomplete="current-password">
          </label>
          <label>
            Mật khẩu mới
            <input type="password" id="newPw" name="new_password" required minlength="4" placeholder="Nhập mật khẩu mới (tối thiểu 4 ký tự)" autocomplete="new-password">
          </label>
          <label>
            Xác nhận mật khẩu mới
            <input type="password" id="confirmPw" name="confirm_password" required minlength="4" placeholder="Nhập lại mật khẩu mới" autocomplete="new-password">
          </label>
          <div class="toggle-pw-row">
            <input type="checkbox" id="showPwToggle" onchange="togglePasswords()">
            <label for="showPwToggle" style="margin: 0; font-weight: normal; cursor: pointer;">Hiện mật khẩu</label>
          </div>
          <p style="margin: 0 0 14px; font-size: 13px; color: #637089;">
            * Mật khẩu mới phải có ít nhất 4 ký tự và không được đặt là 123456.
          </p>
          <button type="submit" style="width: 100%;">Cập nhật mật khẩu</button>
        </form>
      </section>
    </div>
  </div>

  <script>
    function togglePasswords() {
      const isChecked = document.getElementById('showPwToggle').checked;
      const type = isChecked ? 'text' : 'password';
      document.getElementById('curPw').type = type;
      document.getElementById('newPw').type = type;
      document.getElementById('confirmPw').type = type;
    }
  </script>
</body>
</html>
"""

ADMIN_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Quản trị người dùng & Phân quyền · Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 1240px; margin: 0 auto; padding: 28px; }
    h1 { margin: 0; color: #163d68; }
    h2 { color: #163d68; margin: 0 0 10px; font-size: 19px; }
    p { line-height: 1.45; }
    .muted { color: #637089; font-size: 14px; }
    .card, form.panel { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: 0 2px 10px #1720330c; }
    .app-header, .summary, .actions { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; justify-content: space-between; }
    .actions { justify-content: flex-start; }
    button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-block; padding: 9px 13px; border-radius: 7px; font: inherit; font-size: 14px; white-space: nowrap; }
    button.secondary, .button.secondary { color: #1e5b91; background: #e9f1f8; }
    button.warning, .button.warning { background: #fff3cd; color: #856404; border: 1px solid #ffeeba; }
    button.warning:hover, .button.warning:hover { background: #ffe8a1; }
    button.danger, .button.danger { background: #fdecec; color: #8f2f2f; border: 1px solid #f5c6cb; }
    button.danger:hover, .button.danger:hover { background: #f8d7da; }
    .notice { padding: 12px 14px; border-radius: 8px; background: #e9f4eb; color: #265b31; margin-bottom: 14px; }
    .notice.error { background: #fdecec; color: #8f2f2f; }
    
    .grid-form { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 14px; align-items: end; }
    label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; font-size: 14px; }
    input[type="text"], input[type="password"] { box-sizing: border-box; font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 9px 11px; width: 100%; background: #fff; }
    
    .permissions-box { margin-top: 14px; padding: 14px; background: #f9fbfd; border: 1px solid #e2e8f0; border-radius: 8px; }
    .permissions-title { font-weight: 700; color: #163d68; margin-bottom: 8px; font-size: 14px; }
    .checkbox-group { display: flex; flex-wrap: wrap; gap: 18px; }
    .checkbox-group.vertical { display: flex; flex-direction: column; gap: 10px; }
    .check-label { display: flex; flex-direction: row; align-items: center; gap: 6px; font-weight: 500; font-size: 14px; cursor: pointer; }
    .check-label input[type="checkbox"] { width: 17px; height: 17px; cursor: pointer; accent-color: #1e5b91; }

    .perm-option { display: flex; align-items: flex-start; gap: 12px; padding: 10px 12px; border: 1px solid #e2e8f0; border-radius: 8px; background: #fff; cursor: pointer; transition: background 0.15s ease, border-color 0.15s ease; }
    .perm-option:hover { background: #f8fafc; border-color: #cbd5e1; }
    .perm-option input[type="checkbox"] { width: 18px; height: 18px; margin-top: 2px; cursor: pointer; accent-color: #1e5b91; flex-shrink: 0; }
    .perm-info { display: flex; flex-direction: column; gap: 2px; }
    .perm-title { font-weight: 700; color: #163d68; font-size: 14.5px; }
    .perm-desc { font-size: 13px; color: #505f79; line-height: 1.4; font-weight: normal; }

    .table-wrap { overflow-x: auto; margin-top: 14px; }
    table { width: 100%; min-width: 800px; border-collapse: collapse; background: #fff; }
    th, td { text-align: left; padding: 11px; border-bottom: 1px solid #e4e9f0; vertical-align: middle; }
    th { color: #344a66; background: #eef3f8; font-size: 14px; }
    td.actions-cell { white-space: nowrap; }
    .inline-form { display: inline; }

    tr.user-row { cursor: pointer; transition: background-color 0.15s ease; }
    tr.user-row:hover { background-color: #edf4fc; }
    tr.user-row:focus { outline: 2px solid #1e5b91; background-color: #edf4fc; }

    .role-badge { display: inline-block; padding: 2px 9px; border-radius: 12px; font-size: 12px; font-weight: 600; }
    .role-badge.admin { background: #e0edff; color: #163d68; border: 1px solid #b8d5ff; }
    .role-badge.user { background: #eaf5ea; color: #23632f; border: 1px solid #c2e3c5; }
    .role-badge.warning { background: #fff3cd; color: #856404; border: 1px solid #ffeeba; }

    .perm-badge { display: inline-block; padding: 3px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-right: 4px; }
    .perm-badge.active { background: #e8f5e9; color: #1b5e20; border: 1px solid #a5d6a7; }
    .perm-badge.inactive { background: #f1f3f5; color: #868e96; border: 1px solid #dee2e6; opacity: 0.65; }

    /* Modal Overlay Styles */
    .modal-overlay {
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(15, 23, 42, 0.48);
      z-index: 9999;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 16px;
      box-sizing: border-box;
      backdrop-filter: blur(2px);
    }
    .modal-card {
      background: #fff;
      border-radius: 12px;
      width: 100%;
      max-width: 520px;
      max-height: 90vh;
      overflow-y: auto;
      box-shadow: 0 16px 40px rgba(0, 0, 0, 0.2);
      border: 1px solid #dbe3ee;
    }
    .modal-header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 16px 20px;
      border-bottom: 1px solid #eef2f7;
      background: #f8fafc;
      border-top-left-radius: 12px;
      border-top-right-radius: 12px;
    }
    .modal-title {
      margin: 0;
      font-size: 17px;
      color: #163d68;
    }
    .modal-title .highlight {
      color: #1e5b91;
      font-weight: 700;
    }
    .modal-close-btn {
      background: transparent;
      border: 0;
      font-size: 24px;
      line-height: 1;
      color: #64748b;
      cursor: pointer;
      padding: 2px 8px;
      border-radius: 6px;
    }
    .modal-close-btn:hover {
      background: #e2e8f0;
      color: #0f172a;
    }
    .modal-body {
      padding: 20px;
    }
    .modal-info-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 12px;
      background: #f8fafc;
      padding: 14px;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
    }
    .modal-info-item {
      font-size: 13.5px;
    }
    .modal-info-label {
      display: block;
      font-size: 12px;
      color: #64748b;
      margin-bottom: 2px;
      font-weight: 600;
    }
    details.admin-pw { border-top: 1px solid #eef2f7; margin-top: 14px; padding-top: 12px; }
    details.admin-pw summary { cursor: pointer; color: #1e5b91; font-weight: 600; }
    .form-hint { font-size: 12.5px; color: #637089; margin-top: 4px; }
  </style>
</head>
<body>
  <header class="app-header">
    <div>
      <h1>Quản trị người dùng & Phân quyền</h1>
      <p class="muted">Quản lý danh sách tài khoản, đặt lại mật khẩu và phân quyền thao tác dữ liệu</p>
    </div>
    <div class="actions">
      <a class="button secondary" href="{{ url_for('admin_audit_log') }}">Nhật ký hoạt động</a>
      <a class="button secondary" href="{{ url_for('index') }}">← Về trang tra cứu hồ sơ</a>
      <a class="button secondary" href="{{ url_for('user_profile') }}">Tài khoản & Quyền hạn</a>
      <a class="button secondary" href="{{ url_for('logout') }}">Đăng xuất</a>
    </div>
  </header>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <div class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</div>
    {% endfor %}
  {% endwith %}

  <!-- Admin overview card -->
  <section class="card">
    <div class="summary">
      <div>
        <h2>Tài khoản Quản trị viên (Admin)</h2>
        <p class="muted" style="margin: 0;">
          Tên đăng nhập: <strong>{{ admin_user.username }}</strong> · 
          <span class="role-badge admin">Quản trị viên duy nhất</span>
        </p>
      </div>
      <div>
        <a class="button secondary" href="{{ url_for('user_profile') }}">Xem quyền của bản thân & Đổi mật khẩu cá nhân →</a>
      </div>
    </div>
    <p style="font-size: 14px; color: #4b5875; margin-top: 10px;">
      Hệ thống chỉ duy nhất 1 tài khoản Quản trị viên có toàn quyền thao tác dữ liệu và toàn quyền quản trị người dùng khác.
    </p>
  </section>

  <!-- Create user card -->
  <section class="card">
    <h2>Thêm người dùng mới</h2>
    <p class="muted">Tạo tài khoản người dùng mới và thiết lập các quyền thao tác dữ liệu. Mật khẩu mặc định là <strong>123456</strong>.</p>
    <form action="{{ url_for('admin_create_user') }}" method="post">
      <input type="hidden" name="csrf_token" value="{{ csrf }}">
      <div style="max-width: 420px;">
        <label>
          Tên đăng nhập
          <input type="text" name="username" placeholder="Ví dụ: nhanvien1" required pattern="[a-zA-Z0-9_\-\.]{3,30}" title="3-30 ký tự (chữ, số, _, -, .)">
          <span class="form-hint">Mật khẩu khởi tạo mặc định là <strong>123456</strong>. Người dùng sẽ phải tự cài mật khẩu mới khi đăng nhập lần đầu.</span>
        </label>
      </div>

      <div class="permissions-box">
        <div class="permissions-title">Phân quyền thao tác dữ liệu:</div>
        <div class="checkbox-group vertical">
          <label class="check-label perm-option">
            <input type="checkbox" name="can_create" value="1">
            <span class="perm-info">
              <span class="perm-title">Tạo</span>
              <span class="perm-desc">Thêm mới hồ sơ lý lịch thanh niên vào hệ thống và lưu tự động vào tệp Excel.</span>
            </span>
          </label>
          <label class="check-label perm-option">
            <input type="checkbox" name="can_read" value="1" checked>
            <span class="perm-info">
              <span class="perm-title">Xem</span>
              <span class="perm-desc">Tra cứu, tìm kiếm, lọc danh sách hồ sơ; xem trước lý lịch và xuất/tải tệp Word.</span>
            </span>
          </label>
          <label class="check-label perm-option">
            <input type="checkbox" name="can_update" value="1">
            <span class="perm-info">
              <span class="perm-title">Sửa</span>
              <span class="perm-desc">Chỉnh sửa thông tin hồ sơ thanh niên đã có trong hệ thống và lưu thay đổi vào Excel.</span>
            </span>
          </label>
          <label class="check-label perm-option">
            <input type="checkbox" name="can_delete" value="1">
            <span class="perm-info">
              <span class="perm-title">Xóa</span>
              <span class="perm-desc">Xóa hồ sơ khỏi danh sách và Excel (hệ thống tự động sao lưu dự phòng trước khi xóa).</span>
            </span>
          </label>
        </div>
      </div>

      <p class="actions" style="margin-top: 14px; margin-bottom: 0;">
        <button type="submit">Thêm người dùng</button>
      </p>
    </form>
  </section>

  <!-- Users list card -->
  <section class="card">
    <h2>Danh sách người dùng và phân quyền</h2>
    <p class="muted" style="margin-top: -4px; margin-bottom: 12px; font-size: 13.5px;">💡 Nhấp vào dòng người dùng trong bảng để xem thông tin chi tiết và thao tác quản lý.</p>
    <div class="table-wrap">
      <table>
        <thead>
          <tr>
            <th style="width: 50px;">ID</th>
            <th>Tên đăng nhập</th>
            <th style="width: 120px;">Vai trò</th>
            <th>Quyền dữ liệu</th>
            <th style="width: 130px;">Trạng thái</th>
            <th>Ngày tạo</th>
          </tr>
        </thead>
        <tbody>
          {% for u in users %}
          <tr class="user-row" onclick="openUserModal('user-modal-{{ u.id }}')" tabindex="0" role="button" aria-haspopup="dialog" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault(); openUserModal('user-modal-{{ u.id }}');}" title="Nhấp để xem chi tiết và quản lý người dùng {{ u.username }}">
            <td>{{ u.id }}</td>
            <td><strong>{{ u.username }}</strong></td>
            <td>
              {% if u.role == 'admin' %}
                <span class="role-badge admin">Quản trị viên</span>
              {% else %}
                <span class="role-badge user">Người dùng</span>
              {% endif %}
            </td>
            <td>
              {% if u.role == 'admin' %}
                <span class="perm-badge active" title="Quyền Tạo">Tạo</span>
                <span class="perm-badge active" title="Quyền Xem">Xem</span>
                <span class="perm-badge active" title="Quyền Sửa">Sửa</span>
                <span class="perm-badge active" title="Quyền Xóa">Xóa</span>
                <span class="muted" style="margin-left: 6px; font-size: 13px;">(Toàn quyền)</span>
              {% else %}
                <span class="perm-badge {{ 'active' if u.can_create else 'inactive' }}" title="Quyền Tạo">Tạo</span>
                <span class="perm-badge {{ 'active' if u.can_read else 'inactive' }}" title="Quyền Xem">Xem</span>
                <span class="perm-badge {{ 'active' if u.can_update else 'inactive' }}" title="Quyền Sửa">Sửa</span>
                <span class="perm-badge {{ 'active' if u.can_delete else 'inactive' }}" title="Quyền Xóa">Xóa</span>
              {% endif %}
            </td>
            <td>
              {% if u.role == 'admin' %}
                <span class="perm-badge active">Hoạt động</span>
              {% elif u.must_change_password %}
                <span class="role-badge warning" title="Người dùng phải đổi mật khẩu khi đăng nhập lần tới">Cần đổi MK</span>
              {% else %}
                <span class="role-badge user" title="Mật khẩu người dùng đã cài đặt">Đã cài MK</span>
              {% endif %}
            </td>
            <td class="muted" style="font-size: 13px;">{{ u.created_at[:19].replace('T', ' ') if u.created_at else '' }}</td>
          </tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
  </section>

  <!-- Modals / Overlays for each user -->
  {% for u in users %}
  <div id="user-modal-{{ u.id }}" class="modal-overlay" style="display: none;" onclick="handleBackdropClick(event, 'user-modal-{{ u.id }}')" role="dialog" aria-modal="true" aria-labelledby="modal-title-{{ u.id }}">
    <div class="modal-card">
      <div class="modal-header">
        <h3 id="modal-title-{{ u.id }}" class="modal-title">Thông tin người dùng: <span class="highlight">{{ u.username }}</span></h3>
        <button type="button" class="modal-close-btn" onclick="closeUserModal('user-modal-{{ u.id }}')" aria-label="Đóng" title="Đóng">&times;</button>
      </div>
      <div class="modal-body">
        <div class="modal-info-grid">
          <div class="modal-info-item">
            <span class="modal-info-label">Tên đăng nhập</span>
            <strong>{{ u.username }}</strong>
          </div>
          <div class="modal-info-item">
            <span class="modal-info-label">Vai trò</span>
            {% if u.role == 'admin' %}
              <span class="role-badge admin">Quản trị viên</span>
            {% else %}
              <span class="role-badge user">Người dùng</span>
            {% endif %}
          </div>
          <div class="modal-info-item">
            <span class="modal-info-label">Trạng thái</span>
            {% if u.role == 'admin' %}
              <span class="perm-badge active">Hoạt động</span>
            {% elif u.must_change_password %}
              <span class="role-badge warning" title="Người dùng phải đổi mật khẩu khi đăng nhập lần tới">Cần đổi MK</span>
            {% else %}
              <span class="role-badge user" title="Mật khẩu người dùng đã cài đặt">Đã cài MK</span>
            {% endif %}
          </div>
          <div class="modal-info-item">
            <span class="modal-info-label">Ngày tạo</span>
            <span class="muted">{{ u.created_at[:19].replace('T', ' ') if u.created_at else '—' }}</span>
          </div>
        </div>

        {% if u.role == 'admin' %}
          <div style="margin-top: 16px; padding: 14px; background: #eaf2fb; border-radius: 8px; border-left: 4px solid #1e5b91; font-size: 13.5px; color: #163d68; line-height: 1.5;">
            <strong>★ Quản trị viên duy nhất:</strong> Tài khoản này có toàn quyền trong hệ thống và không thể xóa hay giới hạn quyền.
          </div>
          <div style="margin-top: 20px; display: flex; justify-content: flex-end; gap: 10px;">
            <a class="button secondary" href="{{ url_for('user_profile') }}">Đổi mật khẩu cá nhân →</a>
            <button type="button" class="button secondary" onclick="closeUserModal('user-modal-{{ u.id }}')">Đóng</button>
          </div>
        {% else %}
          <!-- Edit Permissions Form -->
          <form action="{{ url_for('admin_update_user', user_id=u.id) }}" method="post" style="margin-top: 14px;">
            <input type="hidden" name="csrf_token" value="{{ csrf }}">
            
            <div class="permissions-title" style="margin-bottom: 8px;">Phân quyền thao tác dữ liệu:</div>
            <div class="checkbox-group vertical" style="margin-bottom: 16px;">
              <label class="check-label perm-option">
                <input type="checkbox" name="can_create" value="1" {% if u.can_create %}checked{% endif %}>
                <span class="perm-info">
                  <span class="perm-title">Tạo</span>
                  <span class="perm-desc">Thêm mới hồ sơ lý lịch thanh niên vào hệ thống.</span>
                </span>
              </label>
              <label class="check-label perm-option">
                <input type="checkbox" name="can_read" value="1" {% if u.can_read %}checked{% endif %}>
                <span class="perm-info">
                  <span class="perm-title">Xem</span>
                  <span class="perm-desc">Tra cứu danh sách, xem trước lý lịch và xuất tệp Word.</span>
                </span>
              </label>
              <label class="check-label perm-option">
                <input type="checkbox" name="can_update" value="1" {% if u.can_update %}checked{% endif %}>
                <span class="perm-info">
                  <span class="perm-title">Sửa</span>
                  <span class="perm-desc">Chỉnh sửa thông tin hồ sơ và cập nhật vào Excel.</span>
                </span>
              </label>
              <label class="check-label perm-option">
                <input type="checkbox" name="can_delete" value="1" {% if u.can_delete %}checked{% endif %}>
                <span class="perm-info">
                  <span class="perm-title">Xóa</span>
                  <span class="perm-desc">Xóa hồ sơ khỏi danh sách và Excel.</span>
                </span>
              </label>
            </div>

            <div style="display: flex; gap: 10px; justify-content: flex-start; margin-top: 14px;">
              <button type="submit">Lưu phân quyền</button>
              <button type="button" class="button secondary" onclick="closeUserModal('user-modal-{{ u.id }}')">Đóng</button>
            </div>
          </form>

          <!-- Password Reset Section -->
          <div style="margin-top: 20px; padding: 14px; background: #fbfcfe; border: 1px solid #e2e8f0; border-radius: 8px;">
            <div class="permissions-title" style="margin-bottom: 6px; font-size: 13.5px;">Mật khẩu tài khoản:</div>
            <p class="muted" style="font-size: 12.5px; margin: 0 0 10px; line-height: 1.45;">
              Quản trị viên không thể tự thiết lập mật khẩu tùy ý cho người dùng. Khi cần cấp lại, chỉ có thể đặt lại về mặc định <strong>123456</strong> (người dùng sẽ bắt buộc tự cài mật khẩu mới khi đăng nhập).
            </p>
            <form action="{{ url_for('admin_reset_user_password', user_id=u.id) }}" method="post" onsubmit="return confirm('Đặt lại mật khẩu cho người dùng \'{{ u.username }}\' về mặc định (\'123456\')?\n\nNgười dùng sẽ bắt buộc phải tự cài đặt mật khẩu mới khi đăng nhập tiếp theo.');">
              <input type="hidden" name="csrf_token" value="{{ csrf }}">
              <button type="submit" class="button secondary" style="width: 100%; font-size: 13px;" title="Đặt lại mật khẩu về 123456">Đặt lại mật khẩu về mặc định (123456)</button>
            </form>
          </div>

          <div style="margin-top: 16px; display: flex; justify-content: flex-end;">
            <!-- Delete User Button -->
            <form class="inline-form" action="{{ url_for('admin_delete_user', user_id=u.id) }}" method="post" onsubmit="return confirm('Bạn có chắc chắn muốn xóa người dùng \'{{ u.username }}\'? Thao tác này không thể hoàn tác.');">
              <input type="hidden" name="csrf_token" value="{{ csrf }}">
              <button type="submit" class="button danger" style="font-size: 13px;">Xóa người dùng</button>
            </form>
          </div>
        {% endif %}
      </div>
    </div>
  </div>
  {% endfor %}

  <script>
    function openUserModal(modalId) {
      var modal = document.getElementById(modalId);
      if (modal) {
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
      }
    }

    function closeUserModal(modalId) {
      var modal = document.getElementById(modalId);
      if (modal) {
        modal.style.display = 'none';
        document.body.style.overflow = '';
      }
    }

    function handleBackdropClick(event, modalId) {
      if (event.target && event.target.id === modalId) {
        closeUserModal(modalId);
      }
    }

    document.addEventListener('keydown', function(event) {
      if (event.key === 'Escape') {
        var openModals = document.querySelectorAll('.modal-overlay');
        openModals.forEach(function(m) {
          if (m.style.display === 'flex') {
            m.style.display = 'none';
            document.body.style.overflow = '';
          }
        });
      }
    });
  </script>
</body>
</html>
"""

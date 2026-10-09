"""HTML page templates and the renderers that feed them from workbook state."""

from flask import render_template_string, url_for

from quan_ly_ho_so.config import FORM_GROUPS, PREVIEW_FIELD_MAP, PREVIEW_SIBLING_HEADERS
from quan_ly_ho_so.forms.fields import field_definitions, form_values
from quan_ly_ho_so.security import csrf_token
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import fold
from quan_ly_ho_so.web.auth_templates import (
    ADMIN_PAGE,
    FORCE_CHANGE_PASSWORD_PAGE,
    LOGIN_PAGE,
    PROFILE_PAGE,
)
from quan_ly_ho_so.workbook.cache import signature_token

PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 1240px; margin: 0 auto; padding: 28px; } h1 { margin: 0; color: #163d68; } h2 { color: #163d68; margin: 0 0 10px; font-size: 19px; }
    p { line-height: 1.45; } .muted, .status { color: #637089; } .status { font-size: 14px; }
    .card, form.panel { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: 0 2px 10px #1720330c; }
    .app-header, .summary, .actions, .pagination { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; justify-content: space-between; }
    .actions, .pagination { justify-content: flex-start; } .quick-search { margin-top: 18px; }
    .search-row, .filter-grid { display: grid; gap: 12px; } .search-row { grid-template-columns: minmax(260px, 1fr) auto auto; align-items: end; } .filter-grid { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); margin-top: 14px; }
    label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; } input, select, textarea, button { box-sizing: border-box; font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 10px 11px; }
    input, select, textarea { background: #fff; width: 100%; } button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-block; padding: 10px 13px; border-radius: 7px; white-space: nowrap; }
    button.secondary, .button.secondary { color: #1e5b91; background: #e9f1f8; }
    button.danger, .button.danger { background: #fdecec; color: #8f2f2f; border: 1px solid #f5c6cb; }
    button.danger:hover, .button.danger:hover { background: #f8d7da; }
    .notice { padding: 12px 14px; border-radius: 8px; background: #e9f4eb; color: #265b31; } .notice.error { background: #fdecec; color: #8f2f2f; }
    details.advanced { border-top: 1px solid #e4e9f0; margin-top: 18px; padding-top: 14px; } summary { cursor: pointer; color: #1e5b91; font-weight: 700; }
    .table-wrap { overflow-x: auto; margin-top: 14px; padding-bottom: 70px; margin-bottom: -70px; } table { width: 100%; min-width: 800px; border-collapse: collapse; background: #fff; } th, td { text-align: left; padding: 11px; border-bottom: 1px solid #e4e9f0; vertical-align: top; }
    th { color: #344a66; background: #eef3f8; } td.actions-cell { white-space: nowrap; position: relative; } .inline-form { display: inline; }
    .action-dropdown { position: relative; display: inline-block; }
    .dropdown-toggle { padding: 6px 12px; font-size: 13px; font-weight: 500; display: inline-flex; align-items: center; gap: 5px; }
    .dropdown-toggle::after { content: "▾"; font-size: 11px; }
    .dropdown-menu { display: none; position: absolute; right: 0; top: calc(100% + 4px); background: #fff; border: 1px solid #dbe3ee; border-radius: 8px; box-shadow: 0 4px 16px rgba(23,32,51,0.14); min-width: 170px; z-index: 100; padding: 5px 0; }
    .action-dropdown.open .dropdown-menu { display: block; }
    .dropdown-item { display: flex; align-items: center; width: 100%; padding: 8px 14px; box-sizing: border-box; text-decoration: none; color: #172033; font-size: 13px; background: none; border: 0; text-align: left; cursor: pointer; border-radius: 0; }
    .dropdown-item:hover, .dropdown-item:focus { background: #f0f4f9; color: #163d68; outline: none; }
    .dropdown-menu button.danger, .dropdown-item.danger { display: flex; align-items: center; width: 100%; padding: 8px 14px; box-sizing: border-box; color: #8f2f2f; background: none; border: 0; text-align: left; cursor: pointer; border-radius: 0; font-size: 13px; font-weight: normal; }
    .dropdown-menu button.danger:hover, .dropdown-menu button.danger:focus, .dropdown-item.danger:hover, .dropdown-item.danger:focus { background: #fdecec; color: #782020; outline: none; }
    .dropdown-divider { height: 1px; background: #e4e9f0; margin: 4px 0; }
    .pagination { margin-top: 16px; }
    .user-nav { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
    .user-info { font-size: 14px; color: #344a66; }
    .role-badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 12px; font-weight: 600; margin-left: 4px; }
    .role-badge.admin { background: #e0edff; color: #163d68; border: 1px solid #b8d5ff; }
    .role-badge.user { background: #eaf5ea; color: #23632f; border: 1px solid #c2e3c5; }
    @media (max-width: 700px) { body { padding: 14px; } .search-row { grid-template-columns: 1fr; } .search-row button, .search-row .button { text-align: center; } }
  </style>
</head>
<body>
  <header class="app-header">
    <div>
      <h1>Quản lý hồ sơ thanh niên</h1>
      <p class="muted">Tra cứu, cập nhật hồ sơ và tạo tệp Word từ Excel</p>
    </div>
    {% if current_user %}
    <div class="user-nav">
      <span class="user-info">Xin chào, <strong>{{ current_user.username }}</strong>
        {% if current_user.is_admin %}
          <span class="role-badge admin">Quản trị viên</span>
        {% else %}
          <span class="role-badge user">Người dùng</span>
        {% endif %}
      </span>
      <a class="button secondary" href="{{ url_for('user_profile') }}">Tài khoản & Quyền hạn</a>
      {% if current_user.is_admin %}
        <a class="button secondary" href="{{ url_for('admin_dashboard') }}">Quản trị người dùng</a>
      {% endif %}
      <a class="button secondary" href="{{ url_for('logout') }}">Đăng xuất</a>
    </div>
    {% endif %}
  </header>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <p class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</p>
    {% endfor %}
  {% endwith %}
  {% if error %}<p class="notice error">{{ error }}</p>{% endif %}
  {% if not loaded %}
    <section class="card"><h2>Chọn tệp Excel</h2><p>Chọn tệp Excel gốc để ứng dụng đọc trực tiếp. Thay đổi đã lưu trong Excel sẽ xuất hiện sau khi làm mới.</p><div class="actions"><a class="button" href="{{ url_for('choose_workbook') }}">Chọn tệp Excel</a></div>
      <details class="advanced"><summary>Hoặc tải lên bản sao tệp Excel</summary><form class="panel" action="{{ url_for('load_uploaded_workbook') }}" method="post" enctype="multipart/form-data"><label>Tệp Excel<input type="file" name="workbook" accept=".xlsx,.xlsm" required></label><p><button type="submit">Dùng bản sao đã tải lên</button></p></form></details>
    </section>
  {% else %}
    <section class="card summary">
      <div>
        <strong>{{ workbook_name }}</strong>
        <div class="status">Trang tính: {{ sheet_name }} · {{ record_count }} hồ sơ · Đồng bộ SQLite lúc {{ loaded_at }}</div>
      </div>
      <div class="actions">
        <a class="button secondary" href="{{ url_for('refresh') }}">Làm mới</a>
        <a class="button secondary" href="{{ url_for('open_excel') }}">Mở bằng Excel</a>
        {% if current_user and current_user.can_create %}
          <a class="button" href="{{ url_for('new_person') }}">Thêm hồ sơ</a>
        {% endif %}
        <a class="button secondary" href="{{ url_for('choose_workbook') }}">Đổi tệp Excel</a>
        {% if current_user and current_user.can_read %}
          <button type="submit" form="batch-form">Tạo Word hàng loạt (đã chọn)</button>
        {% endif %}
      </div>
    </section>
    {% if read_only %}<p class="notice">Tệp .xlsm chỉ có thể xem và tạo Word trong ứng dụng. Hãy dùng Excel để lưu thay đổi.</p>{% endif %}
    {% if current_user and not current_user.can_read %}
      <section class="card">
        <h2>Không có quyền xem dữ liệu</h2>
        <p class="notice error">Tài khoản của bạn chưa được cấp quyền xem dữ liệu hồ sơ. Vui lòng liên hệ Quản trị viên để được cấp quyền.</p>
      </section>
    {% else %}
      <form class="card quick-search" action="{{ url_for('index') }}" method="get"><h2>Tìm kiếm hồ sơ</h2><div class="search-row"><label>Họ tên, STT hoặc CCCD<input name="q" value="{{ query }}" placeholder="Ví dụ: Nguyễn, 12 hoặc số CCCD" autofocus></label><button type="submit">Tìm kiếm</button><a class="button secondary" href="{{ url_for('index') }}">Xóa tìm kiếm</a></div>
        <details class="advanced" {% if filters_active %}open{% endif %}><summary>Bộ lọc nâng cao{% if filters_active %} đang được áp dụng{% endif %}</summary><div class="filter-grid"><label>Năm sinh<select name="birth_year"><option value="">Tất cả</option>{% for value in options.birth_year %}<option value="{{ value }}" {% if filters.birth_year == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Nghề nghiệp<select name="occupation"><option value="">Tất cả</option>{% for value in options.occupation %}<option value="{{ value }}" {% if filters.occupation == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Trình độ văn hóa<select name="education"><option value="">Tất cả</option>{% for value in options.education %}<option value="{{ value }}" {% if filters.education == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Dân tộc<select name="ethnicity"><option value="">Tất cả</option>{% for value in options.ethnicity %}<option value="{{ value }}" {% if filters.ethnicity == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Tôn giáo<select name="religion"><option value="">Tất cả</option>{% for value in options.religion %}<option value="{{ value }}" {% if filters.religion == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Địa chỉ<input name="address" value="{{ filters.address }}" placeholder="Thường trú hoặc nơi ở hiện nay"></label></div><p class="actions"><button type="submit">Áp dụng bộ lọc</button><a class="button secondary" href="{{ url_for('index') }}">Xóa bộ lọc</a></p></details>
      </form>
      <div class="status">Hiển thị {{ shown_start }}–{{ shown_end }} trong tổng số {{ filtered_count }} hồ sơ phù hợp</div>
      <form id="batch-form" action="{{ url_for('generate_batch') }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"></form>
      <div class="table-wrap">
        <table>
          <thead><tr><th><input type="checkbox" onclick="document.querySelectorAll('.row-select').forEach(function(box){box.checked=this.checked;}, this)"></th><th>STT</th><th>Họ và tên</th><th>Năm sinh</th><th>CCCD</th><th>Nghề nghiệp</th><th>Thao tác</th></tr></thead>
          <tbody>
            {% for record in records %}
            <tr>
              <td><input class="row-select" type="checkbox" name="stt" value="{{ record.stt }}" form="batch-form"></td>
              <td>{{ record.stt }}</td>
              <td>{{ record.name }}</td>
              <td>{{ record.birth_year }}</td>
              <td>{{ record.citizen_id }}</td>
              <td>{{ record.occupation }}</td>
              <td class="actions-cell">
                <div class="action-dropdown">
                  <button type="button" class="button secondary dropdown-toggle" onclick="toggleActionDropdown(this, event)">Thao tác</button>
                  <div class="dropdown-menu">
                    {% if current_user and current_user.can_update %}
                      <a class="dropdown-item" href="{{ url_for('edit_person', stt=record.stt) }}">Sửa</a>
                    {% endif %}
                    {% if current_user and current_user.can_read %}
                      <a class="dropdown-item" href="{{ url_for('preview_person', stt=record.stt) }}">Xem trước</a>
                      <form class="inline-form" action="{{ url_for('generate', stt=record.stt) }}" method="post">
                        <input type="hidden" name="csrf_token" value="{{ csrf }}">
                        <button type="submit" class="dropdown-item">Tạo Word</button>
                      </form>
                    {% endif %}
                    {% if current_user and current_user.is_admin %}
                      <a class="dropdown-item" href="{{ url_for('admin_audit_log', stt=record.stt) }}">Nhật ký hoạt động</a>
                    {% endif %}
                    {% if current_user and current_user.can_delete %}
                      <div class="dropdown-divider"></div>
                      <form class="inline-form" action="{{ url_for('delete_person_route', stt=record.stt) }}" method="post" onsubmit="return confirm('Bạn có chắc chắn muốn xóa hồ sơ STT {{ record.stt }} - {{ record.name }}? (Dữ liệu sẽ được ẩn - xóa mềm)');">
                        <input type="hidden" name="csrf_token" value="{{ csrf }}">
                        <input type="hidden" name="signature" value="{{ signature }}">
                        <button type="submit" class="button danger">Xóa</button>
                      </form>
                    {% endif %}
                  </div>
                </div>
              </td>
            </tr>
            {% else %}
            <tr><td colspan="7">Không tìm thấy hồ sơ phù hợp.</td></tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
      <p class="muted">Chọn nhiều hồ sơ rồi bấm "Tạo Word hàng loạt" để in chung một tệp Word gộp cho cả lô — mỗi tờ sẽ dùng phần trống của trang I-IV để in phần V-VI của người ngay trước, tiết kiệm giấy khi gấp thành tập.</p>
      {% if pages > 1 %}<nav class="pagination" aria-label="Phân trang">{% for page_number in range(1, pages + 1) %}<a class="button {{ 'secondary' if page_number != page else '' }}" href="{{ page_urls[page_number] }}">{{ page_number }}</a>{% endfor %}</nav>{% endif %}
    {% endif %}
    <script>
      function toggleActionDropdown(button, event) {
        event.stopPropagation();
        var dropdown = button.closest('.action-dropdown');
        var wasOpen = dropdown.classList.contains('open');
        document.querySelectorAll('.action-dropdown.open').forEach(function(el) {
          el.classList.remove('open');
        });
        if (!wasOpen) {
          dropdown.classList.add('open');
        }
      }
      document.addEventListener('click', function(e) {
        if (!e.target.closest('.action-dropdown')) {
          document.querySelectorAll('.action-dropdown.open').forEach(function(el) {
            el.classList.remove('open');
          });
        }
      });
      const currentSignature = {{ signature|tojson }}; setInterval(async () => { try { const response = await fetch({{ url_for('api_status')|tojson }}, {cache: 'no-store'}); const status = await response.json(); if (status.signature && currentSignature && status.signature !== currentSignature) window.location.reload(); } catch (error) { console.warn('Không thể kiểm tra thay đổi tệp Excel', error); } }, 5000);
    </script>
  {% endif %}
</body>
</html>
"""



FORM_PAGE = r"""
<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{{ title }} · Quản lý hồ sơ thanh niên</title>
<style>
  :root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; } body { max-width: 1080px; margin: 0 auto; padding: 28px; } h1 { color: #163d68; margin-bottom: 4px; }
  form { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; } .section-card { border: 1px solid #dbe3ee; border-radius: 9px; margin: 14px 0; overflow: hidden; } summary { cursor: pointer; display: flex; justify-content: space-between; gap: 12px; padding: 14px 16px; color: #163d68; font-weight: 700; background: #f7f9fc; } .section-card[open] summary { border-bottom: 1px solid #dbe3ee; background: #eef4fa; } .field-count { color: #637089; font-size: 14px; font-weight: 500; } .section-body { padding: 16px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 14px; } label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; } input, textarea, button, .button { font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 10px 11px; box-sizing: border-box; } input, textarea { width: 100%; } textarea { min-height: 84px; resize: vertical; } .wide { grid-column: 1 / -1; } .actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 22px; } button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; } .button.secondary { color: #1e5b91; background: #e9f1f8; } .notice { padding: 12px 14px; border-radius: 8px; background: #fdecec; color: #8f2f2f; } .muted { color: #637089; } .required { color: #a52a2a; } @media (max-width: 700px) { body { padding: 14px; } .grid { grid-template-columns: 1fr; } }
</style></head><body>
<h1>{{ title }}</h1><p class="muted">Tệp Excel: {{ workbook_name }} · Trang tính: {{ sheet_name }}</p><p class="muted"><span class="required">*</span> Trường bắt buộc</p>
{% if message %}<p class="notice">{{ message }}</p>{% endif %}{% if errors %}<div class="notice"><strong>Chưa thể lưu hồ sơ:</strong><ul>{% for error in errors %}<li>{{ error }}</li>{% endfor %}</ul></div>{% endif %}
<form action="{{ url_for('save_person_route') }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"><input type="hidden" name="original_stt" value="{{ original_stt }}"><input type="hidden" name="signature" value="{{ signature }}"><input type="hidden" name="view" value="form">
  {% for section in sections %}<details class="section-card" {% if section.name in open_sections %}open{% endif %}><summary><span>{{ section.name }}</span><span class="field-count">{{ section.fields|length }} trường</span></summary><div class="section-body"><div class="grid">{% for field in section.fields %}<label class="{{ 'wide' if field.input_type == 'textarea' else '' }}">{{ field.label }}{% if field.required %} <span class="required">*</span>{% endif %}{% if field.input_type == 'textarea' %}<textarea name="{{ field.name }}" {% if read_only %}readonly{% endif %}>{{ values.get(field.name, '') }}</textarea>{% else %}<input type="{{ field.input_type }}" name="{{ field.name }}" value="{{ values.get(field.name, '') }}" {% if field.required %}required{% endif %} {% if read_only %}readonly{% endif %}>{% endif %}</label>{% endfor %}</div></div></details>{% endfor %}
  <div class="actions">{% if not read_only %}<button type="submit">Lưu vào Excel</button>{% endif %}<a class="button secondary" href="{{ preview_url }}">Xem trước biểu mẫu</a><a class="button secondary" href="{{ url_for('index') }}">Quay lại danh sách</a></div>
</form></body></html>
"""


def render_form(title, record=None, values=None, errors=None, message=None):
    definitions = field_definitions()
    errors = errors or []
    open_sections = {"Thông tin cá nhân"}
    for error in errors:
        folded_error = fold(error)
        for definition in definitions:
            if fold(definition["header"]) in folded_error:
                open_sections.add(definition["section"])
    grouped = []
    for section_name, _ in FORM_GROUPS:
        grouped.append({"name": section_name, "fields": [field for field in definitions if field["section"] == section_name]})
    return render_template_string(
        FORM_PAGE,
        title=title,
        workbook_name=state["workbook"].name if state["workbook"] else "",
        sheet_name=state["sheet"] or "",
        sections=grouped,
        values=values if values is not None else form_values(record),
        errors=errors,
        open_sections=open_sections,
        message=message,
        original_stt=record["stt"] if record else "",
        signature=signature_token(state["signature"]),
        csrf=csrf_token(),
        read_only=not state["workbook"] or state["workbook"].suffix.lower() == ".xlsm",
        preview_url=url_for("preview_person", stt=record["stt"]) if record else url_for("new_person_preview"),
    )


PREVIEW_PAGE = r"""
<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{{ title }} · Quản lý hồ sơ thanh niên</title>
<style>
  :root { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #dde3ec; }
  body { max-width: 1500px; margin: 0 auto; padding: 24px; }
  .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: space-between; margin-bottom: 14px; }
  .toolbar h1 { color: #163d68; margin: 0 0 2px; font-size: 20px; } .muted { color: #4b5875; font-size: 13px; }
  .actions { display: flex; gap: 10px; flex-wrap: wrap; }
  button, .button { font: inherit; background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; padding: 10px 14px; border-radius: 7px; }
  .button.secondary { color: #1e5b91; background: #e9f1f8; }
  .notice { padding: 12px 14px; border-radius: 8px; background: #fdecec; color: #8f2f2f; margin-bottom: 10px; }
  .notice.success { background: #e9f4eb; color: #265b31; }

  .sheet { background: #fff; margin: 0 auto 24px; padding: 44px 56px; box-shadow: 0 6px 28px rgba(23,32,51,0.18); font-family: "Times New Roman", Georgia, "Noto Serif", serif; font-size: 14.5px; line-height: 1.55; color: #111; }
  .center { text-align: center; } .bold { font-weight: 700; } .upper { text-transform: uppercase; }
  .doc-meta { text-align: right; font-size: 12px; color: #555; margin: -10px 0 6px; white-space: pre-line; }
  .doc-title { font-size: 21px; letter-spacing: 1px; margin: 10px 0 20px; }
  .underline { text-decoration: underline; }
  .columns { display: flex; gap: 40px; align-items: flex-start; margin-top: 6px; }
  .col { flex: 1; min-width: 0; }
  h3.section { font-weight: 700; margin: 18px 0 8px; font-size: 14.5px; }
  h3.section:first-child { margin-top: 0; }
  p.line { margin: 5px 0; display: flex; flex-wrap: wrap; column-gap: 12px; row-gap: 3px; align-items: baseline; }
  p.note { font-style: italic; font-size: 13px; color: #333; margin: 6px 0; }
  .f { border: 0; border-bottom: 1px solid #24344a; background: transparent; font: inherit; padding: 0 2px; min-width: 110px; cursor: text; }
  .f:focus { outline: none; background: #eef4fb; }
  .f:disabled { border-bottom-style: dashed; color: #a33; font-style: italic; background: transparent; }
  .manual { border-bottom: 1px solid #24344a; display: inline-block; min-width: 90px; }
  .sub { margin-left: 14px; }
  .signature-row { display: flex; justify-content: flex-end; margin-top: 26px; }
  .signature-box { text-align: center; min-width: 230px; }
  .signature-gap { height: 64px; }
  details.page2 { margin-top: 18px; border-top: 1px dashed #bcc4d1; padding-top: 12px; }
  details.page2 summary { cursor: pointer; color: #1e5b91; font-weight: 700; }
  .page2-body p { margin: 6px 0; }

  .extra-card, .section-card { max-width: 1500px; margin: 0 auto 20px; background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; overflow: hidden; }
  .extra-card summary, .section-card summary { cursor: pointer; display: flex; justify-content: space-between; gap: 12px; padding: 14px 16px; color: #163d68; font-weight: 700; background: #f7f9fc; }
  .field-count { color: #637089; font-size: 14px; font-weight: 500; }
  .section-body { padding: 16px; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 14px; }
  .grid label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; font-size: 14px; }
  .grid input, .grid textarea { font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 9px 10px; box-sizing: border-box; width: 100%; }
  .grid textarea { min-height: 76px; resize: vertical; } .wide { grid-column: 1 / -1; }

  @media (max-width: 900px) { .columns { flex-direction: column; } .sheet { padding: 24px 18px; } .signature-row { justify-content: center; } }
</style></head><body>
  <div class="toolbar">
    <div><h1>{{ title }}</h1><p class="muted">Tệp Excel: {{ workbook_name }} · Trang tính: {{ sheet_name }} · STT: {{ stt_value }}</p></div>
    <div class="actions">{% if not read_only %}<button type="submit" form="preview-form">Lưu vào Excel</button>{% endif %}<a class="button secondary" href="{{ form_url }}">Xem dạng danh sách</a><a class="button secondary" href="{{ url_for('index') }}">Quay lại danh sách</a></div>
  </div>
  {% if message %}<p class="notice success">{{ message }}</p>{% endif %}
  {% if errors %}<div class="notice"><strong>Chưa thể lưu hồ sơ:</strong><ul>{% for error in errors %}<li>{{ error }}</li>{% endfor %}</ul></div>{% endif %}

  {% macro f(field, cls="") %}{% if field.name %}<input class="f {{ cls }}" type="text" name="{{ field.name }}" value="{{ field.value }}" form="preview-form" {% if read_only %}readonly{% endif %}>{% else %}<input class="f" type="text" value="" disabled title="Cột này chưa có trong tệp Excel">{% endif %}{% endmacro %}

  <form id="preview-form" action="{{ url_for('save_person_route') }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"><input type="hidden" name="original_stt" value="{{ original_stt }}"><input type="hidden" name="signature" value="{{ signature }}"><input type="hidden" name="view" value="preview"></form>

  <div class="sheet">
    <p class="center bold">CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</p>
    <p class="center bold underline">Độc lập - Tự do - Hạnh phúc</p>
    <p class="doc-meta">Biểu số 09/GNN-2025{{ "\n" }}Khổ biểu 29,7 x 42 cm</p>
    <p class="center bold doc-title">LÝ LỊCH<br>NGHĨA VỤ QUÂN SỰ</p>

    <div class="columns">
      <div class="col">
        <h3 class="section">I. SƠ YẾU LÝ LỊCH</h3>
        <p class="line">Họ, chữ đệm và tên khai sinh (viết chữ in hoa): {{ f(doc.birth_name, "upper") }}</p>
        <p class="line">Họ, chữ đệm và tên thường dùng: {{ f(doc.common_name) }}</p>
        <p class="line">Sinh ngày: {{ f(doc.birth_day) }} tháng: {{ f(doc.birth_month) }} năm: {{ f(doc.birth_year) }} &nbsp; Giới tính (nam, nữ): <span class="manual"></span></p>
        <p class="line">Số thẻ căn cước/CCCD: {{ f(doc.citizen_id) }}</p>
        <p class="line">Nơi đăng ký khai sinh: {{ f(doc.birthplace) }}</p>
        <p class="line">Quê quán: {{ f(doc.hometown) }}</p>
        <p class="line">Dân tộc: {{ f(doc.ethnicity) }} &nbsp; Tôn giáo: {{ f(doc.religion) }} &nbsp; Quốc tịch: Việt Nam</p>
        <p class="line">Nơi thường trú của gia đình: {{ f(doc.permanent_address) }}</p>
        <p class="line">Nơi ở hiện tại của bản thân: {{ f(doc.current_address) }}</p>
        <p class="line">Thành phần gia đình: {{ f(doc.family_background) }} &nbsp; Bản thân: {{ f(doc.personal_status) }}</p>
        <p class="line">Trình độ giáo dục phổ thông: {{ f(doc.education) }}</p>
        <p class="line">Trình độ đào tạo: {{ f(doc.specialization) }} &nbsp; Ngoại ngữ: <span class="manual"></span></p>
        <p class="line">Chuyên ngành đào tạo: {{ f(doc.training_major) }}</p>
        <p class="line">Ngày vào Đảng CSVN: <span class="manual"></span> &nbsp; Chính thức: <span class="manual"></span></p>
        <p class="line">Ngày vào Đoàn TNCS Hồ Chí Minh: <span class="manual"></span></p>
        <p class="line">Khen thưởng: <span class="manual"></span> &nbsp; Kỷ luật: <span class="manual"></span></p>
        <p class="line">Nghề nghiệp: {{ f(doc.occupation) }} &nbsp; Lương: <span class="manual"></span> &nbsp; Ngạch: <span class="manual"></span> &nbsp; Bậc: <span class="manual"></span></p>
        <p class="line">Nơi làm việc, (học tập): {{ f(doc.workplace) }}</p>
        <p class="line">Đã đi nước ngoài (tên nước, thời gian, lý do): <span class="manual" style="min-width:260px;"></span></p>
        <p class="line">Họ tên cha: {{ f(doc.father_name) }} &nbsp; (sống, chết): {{ f(doc.father_life_status) }}</p>
        <p class="line">Sinh ngày: {{ f(doc.father_birth_day) }} tháng {{ f(doc.father_birth_month) }} năm {{ f(doc.father_birth_year) }} &nbsp; Nghề nghiệp: {{ f(doc.father_occupation) }}</p>
        <p class="line">Họ tên mẹ: {{ f(doc.mother_name) }} &nbsp; (sống, chết): {{ f(doc.mother_life_status) }}</p>
        <p class="line">Sinh ngày: {{ f(doc.mother_birth_day) }} tháng {{ f(doc.mother_birth_month) }} năm {{ f(doc.mother_birth_year) }} &nbsp; Nghề nghiệp: {{ f(doc.mother_occupation) }}</p>
        <p class="line">Họ tên vợ (chồng): {{ f(doc.spouse_name) }} &nbsp; Sinh ngày: <span class="manual"></span> tháng: <span class="manual"></span> năm: {{ f(doc.spouse_year) }}</p>
        <p class="line">Nghề nghiệp: {{ f(doc.spouse_occupation) }} &nbsp; Bản thân đã có: {{ f(doc.spouse_children) }} con</p>
        <p class="line">Cha mẹ có {{ f(doc.sibling_count) }} người con, {{ f(doc.brother_count) }} trai, {{ f(doc.sister_count) }} gái; bản thân là con thứ: {{ f(doc.birth_order) }}</p>

        <div class="signature-row"><div class="signature-box"><p class="bold">NGƯỜI KHAI</p><p class="muted">(Ký ghi rõ họ tên)</p><div class="signature-gap"></div>{{ f(doc.common_name) }}</div></div>
      </div>

      <div class="col">
        <h3 class="section">II. TÌNH HÌNH KINH TẾ, CHÍNH TRỊ CỦA GIA ĐÌNH</h3>
        <p class="note">(Của cha đẻ, mẹ đẻ hoặc người trực tiếp nuôi dưỡng của bản thân và của vợ hoặc chồng; anh chị em ruột; con đẻ, con nuôi theo quy định của pháp luật; nghề nghiệp, tình hình kinh tế, chính trị của từng người qua các thời kỳ).</p>
        <p class="line">1. Họ tên cha: {{ f(doc.father_name) }} &nbsp; Sinh năm: {{ f(doc.father_birth_year) }}</p>
        <p class="line sub">- Nghề nghiệp: {{ f(doc.father_occupation) }}</p>
        <p class="line sub">- Trước 30/4/1975: {{ f(doc.father_before_1975) }}</p>
        <p class="line sub">- Sau 30/4/1975: {{ f(doc.father_after_1975) }}</p>
        <p class="line sub">- Hiện nay: {{ f(doc.father_current) }}</p>
        <p class="line sub">Tình hình kinh tế: Ổn định</p>
        <p class="note sub">Thái độ chính trị: Luôn chấp hành tốt đường lối, chủ trương của Đảng, chính sách, pháp luật của Nhà nước, quy định của địa phương nơi cư trú</p>
        <p class="line">2. Họ tên mẹ: {{ f(doc.mother_name) }} &nbsp; Sinh năm: {{ f(doc.mother_birth_year) }}</p>
        <p class="line sub">- Nghề nghiệp: {{ f(doc.mother_occupation) }}</p>
        <p class="line sub">- Trước 30/4/1975: {{ f(doc.mother_before_1975) }}</p>
        <p class="line sub">- Sau 30/4/1975: {{ f(doc.mother_after_1975) }}</p>
        <p class="line sub">- Hiện nay: {{ f(doc.mother_current) }}</p>
        <p class="line sub">Tình hình kinh tế: Ổn định</p>
        <p class="note sub">Thái độ chính trị: Luôn chấp hành tốt đường lối, chủ trương của Đảng, chính sách, pháp luật của Nhà nước, quy định của địa phương nơi cư trú</p>
        <h3 class="section" style="margin-top:14px;">3. Anh/Chị em ruột</h3>
        {% for sibling in siblings %}<p class="line sub">- Anh/chị/em {{ loop.index }}: {{ f(sibling) }}</p>{% endfor %}

        <h3 class="section">III. TÌNH HÌNH KINH TẾ, CHÍNH TRỊ, QUÁ TRÌNH CÔNG TÁC CỦA BẢN THÂN</h3>
        <p class="line">1. Quá trình công tác của bản thân</p>
        <p class="line sub">Cấp 1: {{ f(doc.education_level_1) }}</p>
        <p class="line sub">Cấp 2: {{ f(doc.education_level_2) }}</p>
        <p class="line sub">Cấp 3: {{ f(doc.education_level_3) }}</p>
        <p class="line sub">Hiện nay: {{ f(doc.current_history) }}</p>
        <p class="line">2. Thái độ chính trị</p>
        <p class="note sub">Bản thân luôn chấp hành tốt mọi đường lối chủ trương của Đảng, chính sách và pháp luật của Nhà nước, quy định của địa phương nơi cư trú</p>
      </div>
    </div>

    <details class="page2">
      <summary>Xem trang 2 — Nhận xét và kết luận (điền tay sau khi in, không chỉnh sửa ở đây)</summary>
      <div class="page2-body">
        <h3 class="section">IV. NHẬN XÉT VÀ KẾT LUẬN CỦA CÔNG AN PHƯỜNG</h3>
        <p>Kết luận tiêu chuẩn chính trị theo Điều 4, Thông tư 105/20216/TTLT-BQP-BCA: đủ điều kiện giao quân.</p>
        <h3 class="section">V. KẾT LUẬN CỦA BAN CHỈ HUY QUÂN SỰ PHƯỜNG</h3>
        <h3 class="section">VI. KẾT LUẬN CỦA HỘI ĐỒNG NVQS PHƯỜNG TRƯỚC KHI CÔNG DÂN NHẬP NGŨ</h3>
      </div>
    </details>
  </div>

  <details class="extra-card" {% if extra_fields %}open{% endif %}>
    <summary><span>Trường bổ sung (không hiển thị trong biểu mẫu in)</span><span class="field-count">{{ extra_fields|length }} trường</span></summary>
    <div class="section-body"><div class="grid">{% for field in extra_fields %}<label class="{{ 'wide' if field.input_type == 'textarea' else '' }}">{{ field.label }}{% if field.required %} <span class="required">*</span>{% endif %}{% if field.input_type == 'textarea' %}<textarea name="{{ field.name }}" form="preview-form" {% if read_only %}readonly{% endif %}>{{ values.get(field.name, '') }}</textarea>{% else %}<input type="{{ field.input_type }}" name="{{ field.name }}" form="preview-form" value="{{ values.get(field.name, '') }}" {% if read_only %}readonly{% endif %}>{% endif %}</label>{% endfor %}</div></div>
  </details>

  <script>
    (function () {
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const resize = (input) => {
        if (!input.classList.contains('f')) return;
        const raw = input.value || input.placeholder || '';
        const text = input.classList.contains('upper') ? raw.toUpperCase() : raw;
        const style = getComputedStyle(input);
        ctx.font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
        const width = ctx.measureText(text).width;
        input.style.width = Math.max(width + 16, 110) + 'px';
      };
      document.querySelectorAll('input.f[name], .grid input[name], .grid textarea[name]').forEach((el) => {
        resize(el);
        el.addEventListener('input', () => {
          resize(el);
          document.querySelectorAll(`[name="${CSS.escape(el.name)}"]`).forEach((mirror) => {
            if (mirror !== el) { mirror.value = el.value; resize(mirror); }
          });
        });
      });
    })();
  </script>
</body></html>
"""


GUEST_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Tra cứu hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 900px; margin: 0 auto; padding: 28px; }
    h1 { color: #163d68; margin: 0 0 6px; } .muted { color: #637089; }
    .card { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 24px; margin: 16px 0; box-shadow: 0 2px 10px #1720330c; }
    label { display: flex; flex-direction: column; gap: 6px; font-weight: 600; }
    input, textarea, select, button { font: inherit; border-radius: 7px; border: 1px solid #bdc9d9; padding: 10px 12px; box-sizing: border-box; }
    input, textarea { width: 100%; background: #fff; }
    button, .button { background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-block; padding: 10px 16px; border-radius: 7px; white-space: nowrap; }
    .button.secondary { color: #1e5b91; background: #e9f1f8; }
    .notice { padding: 12px 14px; border-radius: 8px; margin: 12px 0; }
    .notice.error { background: #fdecec; color: #8f2f2f; }
    .notice.success { background: #e9f4eb; color: #265b31; }
    .notice.info { background: #e0edff; color: #163d68; border: 1px solid #b8d5ff; }
    .search-box { display: flex; gap: 12px; align-items: end; }
    .search-box label { flex: 1; }
    .section-card { border: 1px solid #dbe3ee; border-radius: 9px; margin: 14px 0; overflow: hidden; }
    .section-card summary { cursor: pointer; display: flex; justify-content: space-between; gap: 12px; padding: 14px 16px; color: #163d68; font-weight: 700; background: #f7f9fc; }
    .section-card[open] summary { border-bottom: 1px solid #dbe3ee; background: #eef4fa; }
    .field-count { color: #637089; font-size: 14px; font-weight: 500; }
    .section-body { padding: 16px; }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 14px; }
    .wide { grid-column: 1 / -1; }
    textarea { min-height: 84px; resize: vertical; }
    .actions { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 18px; }
    .required { color: #a52a2a; }
    .readonly-field { background: #f7f9fc; border-color: #dbe3ee; color: #344a66; }
    .info-tooltip { position: relative; display: inline-block; cursor: help; }
    .info-tooltip .tooltip-text { visibility: hidden; background: #163d68; color: #fff; text-align: center; padding: 8px 12px; border-radius: 6px; font-size: 13px; font-weight: 400; position: absolute; z-index: 1; bottom: 125%; left: 50%; transform: translateX(-50%); white-space: nowrap; box-shadow: 0 2px 8px rgba(0,0,0,.15); opacity: 0; transition: opacity .2s; }
    .info-tooltip .tooltip-text::after { content: ""; position: absolute; top: 100%; left: 50%; margin-left: -5px; border: 5px solid transparent; border-top-color: #163d68; }
    .info-tooltip:hover .tooltip-text, .info-tooltip:focus .tooltip-text { visibility: visible; opacity: 1; }
    @media (max-width: 700px) {
      body { padding: 14px; }
      .search-box { flex-direction: column; }
      .grid { grid-template-columns: 1fr; }
      .info-tooltip .tooltip-text { white-space: normal; min-width: 200px; left: 0; transform: none; }
    }
  </style>
</head>
<body>
  <h1>🔍 Tra cứu hồ sơ thanh niên</h1>
  <p class="muted">Nhập số CCCD để kiểm tra thông tin hồ sơ hoặc tạo mới</p>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <p class="notice {{ 'error' if category == 'error' else 'success' if category == 'success' else 'info' }}">{{ message }}</p>
    {% endfor %}
  {% endwith %}
  <div class="card">
    <form action="{{ url_for('guest_lookup_post') }}" method="post">
      <input type="hidden" name="csrf_token" value="{{ csrf }}">
      <div class="search-box">
        <label>Số căn cước công dân (CCCD)
          <input type="text" name="cccd" value="{{ cccd }}" placeholder="Nhập 12 chữ số CCCD" pattern="[0-9]{12}" title="CCCD gồm 12 chữ số" required autofocus>
        </label>
        <button type="submit">Tra cứu</button>
      </div>
    </form>
  </div>

  {% if mode == 'readonly' and record %}
  <div class="notice info">
    <span class="info-tooltip" tabindex="0">
      ℹ️ Hồ sơ đã tồn tại — Chỉ hiển thị thông tin.
      <span class="tooltip-text">Kiểm tra thông tin và liên hệ admin nếu cần sửa đổi</span>
    </span>
  </div>
  {% for section in sections %}
  <details class="section-card" {% if loop.first %}open{% endif %}>
    <summary><span>{{ section.name }}</span><span class="field-count">{{ section.fields|length }} trường</span></summary>
    <div class="section-body"><div class="grid">
      {% for field in section.fields %}
      <label class="{{ 'wide' if field.input_type == 'textarea' else '' }}">
        {{ field.label }}
        {% if field.input_type == 'textarea' %}
        <textarea class="readonly-field" readonly>{{ values.get(field.name, '') }}</textarea>
        {% else %}
        <input type="{{ field.input_type }}" class="readonly-field" value="{{ values.get(field.name, '') }}" readonly>
        {% endif %}
      </label>
      {% endfor %}
    </div></div>
  </details>
  {% endfor %}

  {% elif mode == 'create' %}
  <p class="notice info">Không tìm thấy hồ sơ với CCCD <strong>{{ cccd }}</strong>. Vui lòng điền thông tin bên dưới để tạo hồ sơ mới.</p>
  <form action="{{ url_for('guest_save') }}" method="post">
    <input type="hidden" name="csrf_token" value="{{ csrf }}">
    <input type="hidden" name="cccd" value="{{ cccd }}">
    <input type="hidden" name="signature" value="{{ signature }}">
    {% for section in sections %}
    <details class="section-card" {% if loop.first %}open{% endif %}>
      <summary><span>{{ section.name }}</span><span class="field-count">{{ section.fields|length }} trường</span></summary>
      <div class="section-body"><div class="grid">
        {% for field in section.fields %}
        <label class="{{ 'wide' if field.input_type == 'textarea' else '' }}">
          {{ field.label }}{% if field.required %} <span class="required">*</span>{% endif %}
          {% if field.input_type == 'textarea' %}
          <textarea name="{{ field.name }}" {% if field.name == cccd_field_name %}readonly{% endif %}>{{ values.get(field.name, '') }}</textarea>
          {% else %}
          <input type="{{ field.input_type }}" name="{{ field.name }}" value="{{ values.get(field.name, '') }}" {% if field.required %}required{% endif %} {% if field.name == cccd_field_name %}readonly{% endif %}>
          {% endif %}
        </label>
        {% endfor %}
      </div></div>
    </details>
    {% endfor %}
    {% if errors %}
    <div class="notice error"><strong>Chưa thể lưu hồ sơ:</strong><ul>{% for error in errors %}<li>{{ error }}</li>{% endfor %}</ul></div>
    {% endif %}
    <div class="actions">
      <button type="submit">Gửi hồ sơ</button>
      <a class="button secondary" href="{{ url_for('guest_lookup') }}">Quay lại</a>
    </div>
  </form>
  {% endif %}
</body>
</html>
"""


def render_preview(title, record=None, values=None, errors=None, message=None):
    values = values if values is not None else form_values(record)

    def field_for(header):
        column = state["headers"].get(header)
        if not column:
            return {"name": None, "value": ""}
        name = f"column_{column}"
        return {"name": name, "value": values.get(name, "")}

    doc = {key: field_for(header) for key, header in PREVIEW_FIELD_MAP}
    siblings = [field_for(header) for header in PREVIEW_SIBLING_HEADERS]
    mapped_headers = {header for _, header in PREVIEW_FIELD_MAP} | set(PREVIEW_SIBLING_HEADERS)
    definitions = field_definitions()
    extra_fields = [definition for definition in definitions if definition["header"] not in mapped_headers]
    return render_template_string(
        PREVIEW_PAGE,
        title=title,
        workbook_name=state["workbook"].name if state["workbook"] else "",
        sheet_name=state["sheet"] or "",
        stt_value=record["stt"] if record else "(tự động)",
        doc=doc,
        siblings=siblings,
        extra_fields=extra_fields,
        values=values,
        errors=errors or [],
        message=message,
        original_stt=record["stt"] if record else "",
        signature=signature_token(state["signature"]),
        csrf=csrf_token(),
        read_only=not state["workbook"] or state["workbook"].suffix.lower() == ".xlsm",
        form_url=url_for("edit_person", stt=record["stt"]) if record else url_for("new_person"),
    )


def render_guest_form(mode, cccd="", record=None, values=None, errors=None):
    """Render the guest lookup page in readonly or create mode."""
    from quan_ly_ho_so.forms.fields import field_definitions, form_values
    definitions = field_definitions()
    grouped = []
    for section_name, _ in FORM_GROUPS:
        grouped.append({
            "name": section_name,
            "fields": [field for field in definitions if field["section"] == section_name],
        })
    # Determine the CCCD field name
    cccd_column = state["headers"].get("Căn cước")
    cccd_field_name = f"column_{cccd_column}" if cccd_column else ""
    if values is None:
        if mode == "readonly" and record:
            values = form_values(record)
        elif mode == "create":
            values = form_values(None)
            # Pre-fill CCCD
            if cccd_field_name:
                values[cccd_field_name] = cccd
        else:
            values = {}
    return render_template_string(
        GUEST_PAGE,
        mode=mode,
        cccd=cccd,
        record=record,
        sections=grouped,
        values=values,
        errors=errors or [],
        signature=signature_token(state["signature"]),
        csrf=csrf_token(),
        cccd_field_name=cccd_field_name,
    )


AUDIT_LOG_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Nhật ký hoạt động · Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 1240px; margin: 0 auto; padding: 28px; } h1 { margin: 0 0 6px; color: #163d68; }
    .muted { color: #637089; } .status { font-size: 14px; color: #637089; }
    .card { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: 0 2px 10px #1720330c; }
    button, .button { font: inherit; background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-block; padding: 10px 13px; border-radius: 7px; white-space: nowrap; }
    .button.secondary { color: #1e5b91; background: #e9f1f8; }
    .table-wrap { overflow-x: auto; } table { width: 100%; min-width: 700px; border-collapse: collapse; } th, td { text-align: left; padding: 10px; border-bottom: 1px solid #e4e9f0; vertical-align: top; font-size: 14px; }
    th { color: #344a66; background: #eef3f8; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 10px; font-size: 12px; font-weight: 600; }
    .badge.create { background: #e0f7e9; color: #1a7f37; }
    .badge.update { background: #fff3cd; color: #856404; }
    .badge.delete { background: #fdecec; color: #8f2f2f; }
    .change-list { font-size: 13px; color: #444; margin: 4px 0 0; padding: 0; list-style: none; }
    .change-list li { margin: 2px 0; }
    .change-list .field { font-weight: 600; color: #163d68; }
    .change-list .old { color: #8f2f2f; text-decoration: line-through; }
    .change-list .new { color: #1a7f37; }
    .pagination { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 16px; }
    .actions { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; margin-bottom: 16px; }
    @media (max-width: 700px) { body { padding: 14px; } }
  </style>
</head>
<body>
  <h1>Nhật ký hoạt động</h1>
  <p class="muted">{% if stt %}Lịch sử thao tác cho riêng hồ sơ STT {{ stt }}{% else %}Lịch sử thao tác tạo, sửa, xóa hồ sơ thanh niên{% endif %}</p>
  <div class="actions">
    {% if current_user and current_user.is_admin %}
      <a class="button secondary" href="{{ url_for('admin_dashboard') }}">← Quản trị người dùng</a>
    {% endif %}
    <a class="button secondary" href="{{ url_for('index') }}">← Danh sách hồ sơ</a>
    {% if stt %}
      <a class="button secondary" href="{{ url_for('admin_audit_log') }}">Xem toàn bộ nhật ký</a>
    {% endif %}
  </div>
  <p class="status">Tổng cộng {{ total }} lượt thao tác{% if stt %} cho hồ sơ STT {{ stt }}{% endif %}</p>
  <div class="card">
    <div class="table-wrap">
      <table>
        <thead><tr><th>Thời gian</th><th>Người thực hiện</th><th>Hành động</th><th>STT</th><th>Họ tên</th><th>Chi tiết thay đổi</th></tr></thead>
        <tbody>
          {% for entry in entries %}
          <tr>
            <td>{{ entry.timestamp[:19].replace('T', ' ') }}</td>
            <td><strong>{{ entry.username }}</strong></td>
            <td>
              {% if entry.action == 'create' %}<span class="badge create">Tạo mới</span>
              {% elif entry.action == 'update' %}<span class="badge update">Cập nhật</span>
              {% elif entry.action == 'delete' %}<span class="badge delete">Xóa mềm</span>
              {% else %}<span class="badge">{{ entry.action }}</span>{% endif %}
            </td>
            <td>{{ entry.stt }}</td>
            <td>{{ entry.record_name }}</td>
            <td>
              {% if entry.details and entry.details.changes is defined %}
                <ul class="change-list">
                  {% for change in entry.details.changes %}
                  <li><span class="field">{{ change.field }}</span>: <span class="old">{{ change['from'] or '(trống)' }}</span> → <span class="new">{{ change.to or '(trống)' }}</span></li>
                  {% endfor %}
                </ul>
              {% elif entry.details and entry.details.fields is defined %}
                <ul class="change-list">
                  {% for field, value in entry.details.fields.items() %}
                  <li><span class="field">{{ field }}</span>: <span class="new">{{ value }}</span></li>
                  {% endfor %}
                </ul>
              {% elif entry.details and entry.details.reason is defined %}
                <span class="muted">{{ entry.details.reason }}</span>
              {% endif %}
            </td>
          </tr>
          {% else %}
          <tr><td colspan="6">Chưa có lượt thao tác nào{% if stt %} cho hồ sơ STT {{ stt }}{% endif %}.</td></tr>
          {% endfor %}
        </tbody>
      </table>
    </div>
    {% if pages > 1 %}
    <nav class="pagination">
      {% for p in range(1, pages + 1) %}
      <a class="button {{ 'secondary' if p != page else '' }}" href="{{ url_for('admin_audit_log', page=p, stt=stt) if stt else url_for('admin_audit_log', page=p) }}">{{ p }}</a>
      {% endfor %}
    </nav>
    {% endif %}
  </div>
</body>
</html>
"""


def page_url(page_number, params):
    query = dict(params)
    query["page"] = page_number
    return url_for("index", **query)

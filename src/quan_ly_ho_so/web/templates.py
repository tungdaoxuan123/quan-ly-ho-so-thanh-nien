"""HTML page templates and the renderers that feed them from workbook state."""

from flask import render_template_string, url_for

from quan_ly_ho_so.config import FORM_GROUPS, PREVIEW_FIELD_MAP, PREVIEW_SIBLING_HEADERS
from quan_ly_ho_so.forms.enums import NvqsStatus
from quan_ly_ho_so.forms.fields import field_definitions, form_values
from quan_ly_ho_so.security import csrf_token
from quan_ly_ho_so.state import state
from quan_ly_ho_so.utils.text import fold
from quan_ly_ho_so.workbook.cache import signature_token

BASE_CSS = r"""
    :root { color-scheme: light; --bg: #f3f5f9; --surface: #fff; --border: #dde3ec; --border-strong: #c3ccda; --text: #1a2438; --muted: #64718a; --primary: #1d4f91; --primary-hover: #173f75; --primary-soft: #e7eef8; --ok: #1f5c33; --ok-bg: #e8f5ec; --err: #9a2b2b; --err-bg: #fdeceb; --shadow: 0 1px 2px rgba(20,32,56,.06), 0 4px 14px rgba(20,32,56,.05); font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; color: var(--text); background: var(--bg); }
    *, *::before, *::after { box-sizing: border-box; }
    body { margin: 0; font-size: 15px; line-height: 1.5; background: var(--bg); }
    h1, h2, h3 { color: #12213d; letter-spacing: -.01em; } h1 { font-size: 22px; margin: 0 0 4px; } h2 { font-size: 17px; margin: 0 0 6px; }
    .topbar { background: #12294b; } .topbar-inner { max-width: 1240px; margin: 0 auto; padding: 14px 28px; display: flex; align-items: center; gap: 14px; }
    .brand-mark { width: 38px; height: 38px; border-radius: 10px; background: rgba(255,255,255,.14); color: #fff; display: grid; place-items: center; font-weight: 700; font-size: 14px; letter-spacing: .04em; }
    .topbar h1 { color: #fff; font-size: 17px; margin: 0; } .topbar-sub { font-size: 12.5px; color: #b9c7de; }
    .container { max-width: 1240px; margin: 0 auto; padding: 24px 28px 48px; }
    .muted, .status { color: var(--muted); } .status { font-size: 14px; } .required { color: #b3261e; }
    label { display: flex; flex-direction: column; gap: 6px; font-weight: 600; font-size: 13.5px; color: #2b3852; }
    input, select, textarea { font: inherit; font-weight: 400; color: var(--text); background: #fff; border: 1px solid var(--border-strong); border-radius: 8px; padding: 9px 11px; width: 100%; transition: border-color .15s, box-shadow .15s; }
    input::placeholder, textarea::placeholder { color: #93a0b5; }
    input:focus, select:focus, textarea:focus { outline: none; border-color: var(--primary); box-shadow: 0 0 0 3px rgba(29,79,145,.18); }
    input[readonly], textarea[readonly] { background: #f3f5f9; }
    input[type=checkbox] { width: 16px; height: 16px; padding: 0; accent-color: var(--primary); cursor: pointer; }
    button, .button { font: inherit; font-weight: 600; font-size: 14px; display: inline-flex; align-items: center; justify-content: center; gap: 6px; background: var(--primary); color: #fff; border: 1px solid var(--primary); border-radius: 8px; padding: 9px 15px; cursor: pointer; text-decoration: none; white-space: nowrap; transition: background .15s, border-color .15s; }
    button:hover, .button:hover { background: var(--primary-hover); border-color: var(--primary-hover); }
    button.secondary, .button.secondary { background: #fff; color: var(--primary); border-color: var(--border-strong); }
    button.secondary:hover, .button.secondary:hover { background: var(--primary-soft); border-color: var(--primary); }
    button:focus-visible, .button:focus-visible, summary:focus-visible { outline: none; box-shadow: 0 0 0 3px rgba(29,79,145,.3); }
    button:disabled { opacity: .45; cursor: not-allowed; background: var(--primary); border-color: var(--primary); }
    .card, form.panel { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: var(--shadow); }
    .actions { display: flex; gap: 10px; flex-wrap: wrap; align-items: center; }
    .notice { padding: 12px 16px; border-radius: 8px; background: var(--ok-bg); color: var(--ok); border: 1px solid rgba(31,92,51,.2); margin: 0 0 14px; }
    .notice.error { background: var(--err-bg); color: var(--err); border-color: rgba(154,43,43,.2); }
    summary { cursor: pointer; }
    @media (max-width: 700px) { .topbar-inner, .container { padding-left: 16px; padding-right: 16px; } }
"""

PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Quản lý hồ sơ thanh niên</title>
  <style>
/*BASE_CSS*/
    .summary { display: flex; gap: 16px; flex-wrap: wrap; align-items: center; justify-content: space-between; margin-top: 0; }
    .file-name { font-size: 17px; font-weight: 700; color: #12213d; margin-bottom: 6px; } .meta { display: flex; gap: 8px; flex-wrap: wrap; }
    .chip { background: var(--primary-soft); color: #23446f; border-radius: 999px; padding: 3px 11px; font-size: 12.5px; font-weight: 600; }
    .search-row, .filter-grid { display: grid; gap: 12px; } .search-row { grid-template-columns: minmax(260px, 1fr) auto auto; align-items: end; margin-top: 10px; } .filter-grid { grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); margin: 14px 0; }
    details.advanced { border-top: 1px solid var(--border); margin-top: 18px; padding-top: 14px; }
    details.advanced > summary, .row-menu > summary { list-style: none; } details.advanced > summary::-webkit-details-marker, .row-menu > summary::-webkit-details-marker { display: none; }
    details.advanced > summary { color: var(--primary); font-weight: 600; } details.advanced > summary::before { content: ""; display: inline-block; width: 7px; height: 7px; margin-right: 9px; border-right: 2px solid currentColor; border-bottom: 2px solid currentColor; transform: rotate(-45deg); transition: transform .15s; } details.advanced[open] > summary::before { transform: rotate(45deg) translate(-2px, -2px); }
    .table-card { padding: 0; overflow: hidden; }
    .table-toolbar { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; justify-content: space-between; padding: 14px 20px; border-bottom: 1px solid var(--border); } .batch { display: flex; gap: 14px; align-items: center; }
    .table-wrap { overflow-x: auto; } table { width: 100%; min-width: 820px; border-collapse: collapse; }
    th, td { text-align: left; padding: 12px 16px; border-bottom: 1px solid #edf0f5; vertical-align: middle; }
    th { color: #4a5872; background: #f7f9fc; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; white-space: nowrap; }
    tbody tr:hover { background: #f8fafd; } tbody tr:last-child td { border-bottom: 0; }
    td.stt { color: var(--muted); font-variant-numeric: tabular-nums; } td.name { font-weight: 600; color: #12213d; } td.cccd { font-variant-numeric: tabular-nums; }
    td.actions-cell, th.actions-head { text-align: right; white-space: nowrap; } td.empty { text-align: center; color: var(--muted); padding: 40px 16px; }
    .table-hint { margin: 0; padding: 14px 20px; border-top: 1px solid var(--border); background: #f7f9fc; font-size: 13px; }
    .row-menu > summary { padding: 6px 12px; font-size: 13px; } .row-menu > summary::after { content: ""; width: 6px; height: 6px; margin-left: 4px; border-right: 2px solid currentColor; border-bottom: 2px solid currentColor; transform: rotate(45deg) translateY(-2px); }
    .row-menu-list { position: fixed; z-index: 10; min-width: 160px; background: #fff; border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 10px 28px rgba(20,32,56,.18); padding: 5px; display: flex; flex-direction: column; text-align: left; }
    .row-menu-list a, .row-menu-list button { display: block; width: 100%; text-align: left; padding: 8px 12px; background: none; color: var(--text); border: 0; border-radius: 6px; font-size: 14px; font-weight: 500; cursor: pointer; text-decoration: none; }
    .row-menu-list a:hover, .row-menu-list button:hover { background: var(--primary-soft); color: var(--primary); }
    .pagination { display: flex; gap: 6px; flex-wrap: wrap; justify-content: center; margin-top: 20px; } .pagination .button { min-width: 38px; padding: 7px 10px; }
    @media (max-width: 700px) { .search-row { grid-template-columns: 1fr; } .search-row button, .search-row .button { width: 100%; } }
  </style>
</head>
<body>
  <header class="topbar"><div class="topbar-inner"><span class="brand-mark">HS</span><div><h1>Quản lý hồ sơ thanh niên</h1><div class="topbar-sub">Tra cứu, cập nhật hồ sơ và tạo tệp Word từ Excel</div></div></div></header>
  <main class="container">
  {% with messages = get_flashed_messages(with_categories=true) %}{% for category, message in messages %}<p class="notice {{ 'error' if category == 'error' else '' }}">{{ message }}</p>{% endfor %}{% endwith %}
  {% if error %}<p class="notice error">{{ error }}</p>{% endif %}
  {% if not loaded %}
    <section class="card"><h2>Chọn tệp Excel</h2><p>Chọn tệp Excel gốc để ứng dụng đọc trực tiếp. Thay đổi đã lưu trong Excel sẽ xuất hiện sau khi làm mới.</p><div class="actions"><a class="button" href="{{ url_for('choose_workbook') }}">Chọn tệp Excel</a></div>
      <details class="advanced"><summary>Hoặc tải lên bản sao tệp Excel</summary><form class="panel" action="{{ url_for('load_uploaded_workbook') }}" method="post" enctype="multipart/form-data"><label>Tệp Excel<input type="file" name="workbook" accept=".xlsx,.xlsm" required></label><p><button type="submit">Dùng bản sao đã tải lên</button></p></form></details>
    </section>
  {% else %}
    <section class="card summary"><div><div class="file-name">{{ workbook_name }}</div><div class="meta"><span class="chip">Trang tính: {{ sheet_name }}</span><span class="chip">{{ record_count }} hồ sơ</span><span class="chip">Đồng bộ SQLite lúc {{ loaded_at }}</span></div></div><div class="actions"><a class="button secondary" href="{{ url_for('refresh') }}">Làm mới</a><a class="button secondary" href="{{ url_for('open_excel') }}">Mở bằng Excel</a><a class="button secondary" href="{{ url_for('choose_workbook') }}">Đổi tệp Excel</a><a class="button" href="{{ url_for('new_person') }}">Thêm hồ sơ</a></div></section>
    {% if read_only %}<p class="notice">Tệp .xlsm chỉ có thể xem và tạo Word trong ứng dụng. Hãy dùng Excel để lưu thay đổi.</p>{% endif %}
    <form class="card quick-search" action="{{ url_for('index') }}" method="get"><h2>Tìm kiếm hồ sơ</h2><div class="search-row"><label>Họ tên, STT hoặc CCCD<input name="q" value="{{ query }}" placeholder="Ví dụ: Nguyễn, 12 hoặc số CCCD" autofocus></label><button type="submit">Tìm kiếm</button><a class="button secondary" href="{{ url_for('index') }}">Xóa tìm kiếm</a></div>
      <details class="advanced" {% if filters_active %}open{% endif %}><summary>Bộ lọc nâng cao{% if filters_active %} đang được áp dụng{% endif %}</summary><div class="filter-grid"><label>Năm sinh<select name="birth_year"><option value="">Tất cả</option>{% for value in options.birth_year %}<option value="{{ value }}" {% if filters.birth_year == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Nghề nghiệp<select name="occupation"><option value="">Tất cả</option>{% for value in options.occupation %}<option value="{{ value }}" {% if filters.occupation == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Trình độ văn hóa<select name="education"><option value="">Tất cả</option>{% for value in options.education %}<option value="{{ value }}" {% if filters.education == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Dân tộc<select name="ethnicity"><option value="">Tất cả</option>{% for value in options.ethnicity %}<option value="{{ value }}" {% if filters.ethnicity == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Tôn giáo<select name="religion"><option value="">Tất cả</option>{% for value in options.religion %}<option value="{{ value }}" {% if filters.religion == value %}selected{% endif %}>{{ value }}</option>{% endfor %}</select></label><label>Diện<select name="dien"><option value="">Tất cả</option><option value="{{ unset_dien }}" {% if filters.dien == unset_dien %}selected{% endif %}>&lt;chưa có&gt;</option>{% for code, text in dien_options %}<option value="{{ code }}" {% if filters.dien == code %}selected{% endif %}>{{ text }}</option>{% endfor %}</select></label><label>Địa chỉ<input name="address" value="{{ filters.address }}" placeholder="Thường trú hoặc nơi ở hiện nay"></label></div><p class="actions"><button type="submit">Áp dụng bộ lọc</button><a class="button secondary" href="{{ url_for('index') }}">Xóa bộ lọc</a></p></details>
    </form>
    <form id="batch-form" action="{{ url_for('generate_batch') }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"></form>
    <section class="card table-card">
      <div class="table-toolbar"><span class="status">Hiển thị {{ shown_start }}–{{ shown_end }} trong tổng số {{ filtered_count }} hồ sơ phù hợp</span><div class="batch"><span id="selected-count" class="status">Chưa chọn hồ sơ</span><button type="submit" form="batch-form" id="batch-btn" disabled>Tạo Word hàng loạt</button></div></div>
      <div class="table-wrap"><table><thead><tr><th><input type="checkbox" id="select-all" aria-label="Chọn tất cả"></th><th>STT</th><th>Họ và tên</th><th>Năm sinh</th><th>CCCD</th><th>Nghề nghiệp</th><th>Diện</th><th class="actions-head">Thao tác</th></tr></thead><tbody>{% for record in records %}<tr><td><input class="row-select" type="checkbox" name="stt" value="{{ record.stt }}" form="batch-form" aria-label="Chọn hồ sơ {{ record.stt }}"></td><td class="stt">{{ record.stt }}</td><td class="name">{{ record.name }}</td><td>{{ record.birth_year }}</td><td class="cccd">{{ record.citizen_id }}</td><td>{{ record.occupation }}</td><td>{% if record.dien %}{{ record.dien }}{% else %}<span class="muted">&lt;chưa có&gt;</span>{% endif %}</td><td class="actions-cell"><details class="row-menu" ontoggle="positionRowMenu(this)"><summary class="button secondary">Thao tác</summary><div class="row-menu-list"><a href="{{ url_for('edit_person', stt=record.stt) }}">Sửa</a><a href="{{ url_for('preview_person', stt=record.stt) }}">Xem trước</a><form action="{{ url_for('generate', stt=record.stt) }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"><button type="submit">Tạo Word</button></form></div></details></td></tr>{% else %}<tr><td colspan="8" class="empty">Không tìm thấy hồ sơ phù hợp.</td></tr>{% endfor %}</tbody></table></div>
      <p class="table-hint muted">Chọn nhiều hồ sơ rồi bấm "Tạo Word hàng loạt" để in chung một tệp Word gộp cho cả lô — mỗi tờ sẽ dùng phần trống của trang I-IV để in phần V-VI của người ngay trước, tiết kiệm giấy khi gấp thành tập.</p>
    </section>
    {% if pages > 1 %}<nav class="pagination" aria-label="Phân trang">{% for page_number in range(1, pages + 1) %}<a class="button {{ 'secondary' if page_number != page else '' }}" href="{{ page_urls[page_number] }}">{{ page_number }}</a>{% endfor %}</nav>{% endif %}
    <script>
      function positionRowMenu(d) { if (!d.open) return; document.querySelectorAll('.row-menu[open]').forEach(function(o){ if (o !== d) o.open = false; }); var r = d.getBoundingClientRect(), m = d.querySelector('.row-menu-list'); m.style.left = r.left + 'px'; m.style.top = (r.bottom + 4) + 'px'; }
      document.addEventListener('click', function(e){ document.querySelectorAll('.row-menu[open]').forEach(function(o){ if (!o.contains(e.target)) o.open = false; }); });
      function updateSelection() { var n = document.querySelectorAll('.row-select:checked').length; document.getElementById('selected-count').textContent = n ? 'Đã chọn ' + n + ' hồ sơ' : 'Chưa chọn hồ sơ'; document.getElementById('batch-btn').disabled = !n; }
      document.getElementById('select-all').addEventListener('change', function(){ var on = this.checked; document.querySelectorAll('.row-select').forEach(function(b){ b.checked = on; }); updateSelection(); });
      document.querySelectorAll('.row-select').forEach(function(b){ b.addEventListener('change', updateSelection); });
      updateSelection();
      window.addEventListener('scroll', function(){ document.querySelectorAll('.row-menu[open]').forEach(function(o){ o.open = false; }); }, true);
    </script>
    <script>const currentSignature = {{ signature|tojson }}; setInterval(async () => { try { const response = await fetch({{ url_for('api_status')|tojson }}, {cache: 'no-store'}); const status = await response.json(); if (status.signature && currentSignature && status.signature !== currentSignature) window.location.reload(); } catch (error) { console.warn('Không thể kiểm tra thay đổi tệp Excel', error); } }, 5000);</script>
  {% endif %}
  </main>
</body>
</html>
"""


FORM_PAGE = r"""
<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{{ title }} · Quản lý hồ sơ thanh niên</title>
<style>
/*BASE_CSS*/
  .container { max-width: 1080px; }
  form.record { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 20px; box-shadow: var(--shadow); margin-top: 16px; }
  .section-card { border: 1px solid var(--border); border-radius: 10px; margin: 0 0 14px; overflow: hidden; background: #fff; }
  .section-card > summary { list-style: none; display: flex; justify-content: space-between; align-items: center; gap: 12px; padding: 14px 18px; color: #12213d; font-weight: 700; background: #f7f9fc; }
  .section-card > summary::-webkit-details-marker { display: none; }
  .section-card > summary > span:first-child::before { content: ""; display: inline-block; width: 7px; height: 7px; margin-right: 10px; border-right: 2px solid var(--primary); border-bottom: 2px solid var(--primary); transform: rotate(-45deg); transition: transform .15s; }
  .section-card[open] > summary > span:first-child::before { transform: rotate(45deg) translate(-2px, -2px); }
  .section-card[open] > summary { border-bottom: 1px solid var(--border); background: var(--primary-soft); }
  .field-count { color: var(--muted); font-size: 13px; font-weight: 500; } .section-body { padding: 18px; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 16px; } textarea { min-height: 84px; resize: vertical; } .wide { grid-column: 1 / -1; }
  .form-actions { position: sticky; bottom: 0; display: flex; gap: 10px; flex-wrap: wrap; margin: 20px -20px -20px; padding: 14px 20px; background: rgba(255,255,255,.96); border-top: 1px solid var(--border); border-radius: 0 0 12px 12px; backdrop-filter: blur(4px); }
  .status-card { border: 1px solid #cddcec; background: #f2f7fc; border-radius: 9px; padding: 14px 16px; margin-bottom: 14px; } .status-card label { max-width: 520px; } .status-card .hint { color: #4b5875; font-size: 13px; font-weight: 500; margin: 6px 0 0; }
  @media (max-width: 700px) { .grid { grid-template-columns: 1fr; } }
</style></head><body>
<header class="topbar"><div class="topbar-inner"><span class="brand-mark">HS</span><div><div class="topbar-title" style="color:#fff;font-weight:700;font-size:17px">Quản lý hồ sơ thanh niên</div><div class="topbar-sub">Tra cứu, cập nhật hồ sơ và tạo tệp Word từ Excel</div></div></div></header>
<main class="container">
<h1>{{ title }}</h1><p class="muted">Tệp Excel: {{ workbook_name }} · Trang tính: {{ sheet_name }}</p><p class="muted"><span class="required">*</span> Trường bắt buộc</p>
{% if message %}<p class="notice">{{ message }}</p>{% endif %}{% if errors %}<div class="notice error"><strong>Chưa thể lưu hồ sơ:</strong><ul>{% for error in errors %}<li>{{ error }}</li>{% endfor %}</ul></div>{% endif %}
<form class="record" action="{{ url_for('save_person_route') }}" method="post"><input type="hidden" name="csrf_token" value="{{ csrf }}"><input type="hidden" name="original_stt" value="{{ original_stt }}"><input type="hidden" name="signature" value="{{ signature }}"><input type="hidden" name="view" value="form">
  <div class="status-card"><label>Diện (nghĩa vụ quân sự)<select name="nvqs_type" {% if read_only %}disabled{% endif %}><option value="">&lt;chưa có&gt;</option>{% for code, text in nvqs_options %}<option value="{{ code }}" {% if nvqs_type == code %}selected{% endif %}>{{ text }}</option>{% endfor %}</select></label><p class="hint">Diện được lưu cùng hồ sơ trong cơ sở dữ liệu của ứng dụng.</p></div>
  {% for section in sections %}<details class="section-card" {% if section.name in open_sections %}open{% endif %}><summary><span>{{ section.name }}</span><span class="field-count">{{ section.fields|length }} trường</span></summary><div class="section-body"><div class="grid">{% for field in section.fields %}<label class="{{ 'wide' if field.input_type == 'textarea' else '' }}"><span>{{ field.label }}{% if field.required %} <span class="required">*</span>{% endif %}</span>{% if field.input_type == 'textarea' %}<textarea name="{{ field.name }}" {% if read_only %}readonly{% endif %}>{{ values.get(field.name, '') }}</textarea>{% else %}<input type="{{ field.input_type }}" name="{{ field.name }}" value="{{ values.get(field.name, '') }}" {% if field.required %}required{% endif %} {% if read_only %}readonly{% endif %}>{% endif %}</label>{% endfor %}</div></div></details>{% endfor %}
  <div class="form-actions">{% if not read_only %}<button type="submit">Lưu vào Excel</button>{% endif %}<a class="button secondary" href="{{ preview_url }}">Xem trước biểu mẫu</a><a class="button secondary" href="{{ url_for('index') }}">Quay lại danh sách</a></div>
</form></main></body></html>
"""


def render_form(title, record=None, values=None, errors=None, message=None, nvqs_type=None):
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
        nvqs_options=NvqsStatus.options(),
        nvqs_type=nvqs_type if nvqs_type is not None else ((record or {}).get("type") or ""),
    )


PREVIEW_PAGE = r"""
<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>{{ title }} · Quản lý hồ sơ thanh niên</title>
<style>
/*BASE_CSS*/
  :root { background: #dde3ec; } body { background: #dde3ec; max-width: 1500px; margin: 0 auto; padding: 0 24px 24px; }
  .toolbar { position: sticky; top: 0; z-index: 20; display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; margin: 0 -24px 20px; padding: 12px 24px; background: #fff; border-bottom: 1px solid var(--border); box-shadow: 0 2px 10px rgba(20,32,56,.08); }
  .toolbar h1 { font-size: 18px; margin: 0 0 2px; } .muted { color: var(--muted); font-size: 13px; } .toolbar p { margin: 0; }
  .notice.success { background: var(--ok-bg); color: var(--ok); } .notice { max-width: 1500px; }

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
  .f { width: auto; border-radius: 0; box-shadow: none; } .f:focus { outline: none; background: #eef4fb; box-shadow: none; }
  .f:disabled { border-bottom-style: dashed; color: #a33; font-style: italic; background: transparent; }
  .manual { border-bottom: 1px solid #24344a; display: inline-block; min-width: 90px; }
  .sub { margin-left: 14px; }
  .signature-row { display: flex; justify-content: flex-end; margin-top: 26px; }
  .signature-box { text-align: center; min-width: 230px; }
  .signature-gap { height: 64px; }
  details.page2 { margin-top: 18px; border-top: 1px dashed #bcc4d1; padding-top: 12px; }
  details.page2 summary { cursor: pointer; color: #1e5b91; font-weight: 700; }
  .page2-body p { margin: 6px 0; }

  .extra-card, .section-card { max-width: 1500px; margin: 0 auto 20px; background: #fff; border: 1px solid var(--border); border-radius: 12px; overflow: hidden; box-shadow: var(--shadow); }
  .extra-card summary, .section-card summary { cursor: pointer; display: flex; justify-content: space-between; gap: 12px; padding: 14px 16px; color: #12213d; font-weight: 700; background: #f7f9fc; }
  .field-count { color: var(--muted); font-size: 13px; font-weight: 500; }
  .section-body { padding: 18px; font-family: inherit; }
  .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 14px; }
  .grid label { display: flex; flex-direction: column; gap: 5px; font-weight: 600; font-size: 14px; }
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
    <div class="section-body"><div class="grid">{% for field in extra_fields %}<label class="{{ 'wide' if field.input_type == 'textarea' else '' }}"><span>{{ field.label }}{% if field.required %} <span class="required">*</span>{% endif %}</span>{% if field.input_type == 'textarea' %}<textarea name="{{ field.name }}" form="preview-form" {% if read_only %}readonly{% endif %}>{{ values.get(field.name, '') }}</textarea>{% else %}<input type="{{ field.input_type }}" name="{{ field.name }}" form="preview-form" value="{{ values.get(field.name, '') }}" {% if read_only %}readonly{% endif %}>{% endif %}</label>{% endfor %}</div></div>
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

PAGE = PAGE.replace("/*BASE_CSS*/", BASE_CSS)
FORM_PAGE = FORM_PAGE.replace("/*BASE_CSS*/", BASE_CSS)
PREVIEW_PAGE = PREVIEW_PAGE.replace("/*BASE_CSS*/", BASE_CSS)


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


def page_url(page_number, params):
    query = dict(params)
    query["page"] = page_number
    return url_for("index", **query)

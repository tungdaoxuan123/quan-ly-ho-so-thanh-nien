"""Templates for Admin Pending Approval Review Queue."""

APPROVALS_PAGE = r"""
<!doctype html>
<html lang="vi">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Hàng đợi duyệt chỉnh sửa hồ sơ · Quản lý hồ sơ thanh niên</title>
  <style>
    :root { color-scheme: light; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #172033; background: #f5f7fb; }
    body { max-width: 1240px; margin: 0 auto; padding: 28px; }
    h1 { margin: 0 0 6px; color: #163d68; }
    .muted { color: #637089; font-size: 14px; }
    .app-header { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; justify-content: space-between; margin-bottom: 16px; }
    .actions { display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
    .card { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; padding: 20px; margin: 16px 0; box-shadow: 0 2px 10px #1720330c; }
    
    button, .button { font: inherit; background: #1e5b91; border: 0; color: #fff; cursor: pointer; text-decoration: none; display: inline-flex; align-items: center; justify-content: center; padding: 9px 14px; border-radius: 7px; white-space: nowrap; font-size: 14px; box-sizing: border-box; }
    .button.secondary { color: #1e5b91; background: #e9f1f8; }
    .button.active { background: #163d68; font-weight: 600; cursor: default; }
    .button.danger, button.danger { background: #dc2626; color: #fff; }
    .button.danger:hover, button.danger:hover { background: #b91c1c; }
    .button.approve-btn, button.approve-btn { background: #16a34a; color: #fff; }
    .button.approve-btn:hover, button.approve-btn:hover { background: #15803d; }

    .notice { padding: 12px 14px; border-radius: 8px; margin-bottom: 14px; }
    .notice.success { background: #e9f4eb; color: #265b31; border: 1px solid #bbf7d0; }
    .notice.error { background: #fdecec; color: #8f2f2f; border: 1px solid #fecaca; }
    .notice.info { background: #e0edff; color: #163d68; border: 1px solid #b8d5ff; }

    /* Summary Stats */
    .stats-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 16px; }
    .stat-card { background: #fff; border: 1px solid #dbe3ee; border-radius: 10px; padding: 14px 18px; box-shadow: 0 1px 4px #17203308; }
    .stat-val { font-size: 24px; font-weight: 700; color: #163d68; margin-top: 4px; }
    .stat-val.pending { color: #d97706; }
    .stat-label { font-size: 13px; color: #637089; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px; }

    /* Filter tabs & search */
    .filter-bar { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; justify-content: space-between; margin: 16px 0; }
    .tabs { display: flex; gap: 6px; flex-wrap: wrap; }
    .tab-item { display: inline-flex; align-items: center; gap: 6px; padding: 8px 14px; border-radius: 8px; font-size: 14px; font-weight: 500; text-decoration: none; color: #475569; background: #e2e8f0; }
    .tab-item:hover { background: #cbd5e1; }
    .tab-item.active { background: #1e5b91; color: #fff; font-weight: 600; }
    .tab-badge { background: #d97706; color: #fff; font-size: 11px; padding: 2px 6px; border-radius: 10px; font-weight: 700; }

    .search-box { display: flex; gap: 8px; flex: 1; max-width: 400px; }
    .search-box input { flex: 1; font: inherit; font-size: 14px; padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 7px; background: #fff; }
    .search-box input:focus { outline: none; border-color: #1e5b91; box-shadow: 0 0 0 2px #1e5b9122; }

    /* Request Cards */
    .request-card { background: #fff; border: 1px solid #dbe3ee; border-radius: 12px; margin-bottom: 16px; overflow: hidden; box-shadow: 0 2px 8px #17203308; }
    .request-card.pending { border-left: 5px solid #f59e0b; }
    .request-card.approved { border-left: 5px solid #16a34a; }
    .request-card.rejected { border-left: 5px solid #dc2626; }
    
    .card-header { padding: 14px 18px; background: #f8fafc; border-bottom: 1px solid #edf2f7; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }
    .card-title { font-size: 15px; font-weight: 700; color: #163d68; }
    .card-meta { font-size: 13px; color: #64748b; margin-top: 3px; }

    .badge { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
    .badge.pending { background: #fef3c7; color: #b45309; border: 1px solid #fde68a; }
    .badge.approved { background: #dcfce7; color: #15803d; border: 1px solid #bbf7d0; }
    .badge.rejected { background: #fee2e2; color: #b91c1c; border: 1px solid #fecaca; }

    .card-body { padding: 18px; }

    /* Diff View */
    .diff-container { background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 16px; }
    .diff-title { font-size: 13.5px; font-weight: 700; color: #334155; margin-bottom: 8px; display: flex; align-items: center; gap: 6px; }
    .diff-table { width: 100%; border-collapse: collapse; font-size: 13.5px; background: #fff; border-radius: 6px; overflow: hidden; border: 1px solid #e2e8f0; }
    .diff-table th { background: #f1f5f9; padding: 9px 12px; border-bottom: 1px solid #cbd5e1; text-align: left; color: #475569; font-weight: 600; }
    .diff-table td { padding: 9px 12px; border-bottom: 1px solid #e2e8f0; vertical-align: top; }
    .diff-table tr:last-child td { border-bottom: 0; }
    .diff-field { font-weight: 600; color: #1e3a8a; width: 28%; }
    .diff-old { width: 36%; color: #991b1b; background: #fff1f2; }
    .diff-old del { text-decoration: line-through; }
    .diff-new { width: 36%; color: #166534; background: #f0fdf4; font-weight: 600; }
    .diff-new ins { text-decoration: none; }

    /* Review Form */
    .review-form { background: #f8fafc; border-top: 1px solid #edf2f7; padding: 14px 18px; }
    .review-grid { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
    .review-grid label { flex: 1; min-width: 280px; font-size: 13px; font-weight: 600; color: #334155; display: flex; flex-direction: column; gap: 4px; }
    .review-grid input[type="text"] { font: inherit; font-size: 14px; padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 6px; background: #fff; width: 100%; box-sizing: border-box; }
    .review-grid input[type="text"]:focus { outline: none; border-color: #1e5b91; box-shadow: 0 0 0 2px #1e5b9122; }
    .btn-group { display: flex; gap: 8px; align-items: flex-end; }

    .reviewed-history { background: #f8fafc; border-top: 1px solid #edf2f7; padding: 12px 18px; font-size: 13px; color: #475569; }

    .empty-queue { text-align: center; padding: 48px 20px; color: #64748b; }
    .empty-queue-icon { font-size: 40px; margin-bottom: 10px; }

    .pagination { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 16px; align-items: center; }
    .pagination .page-info { margin-left: auto; font-size: 13px; color: #637089; }
    
    @media (max-width: 768px) {
      body { padding: 14px; }
      .app-header { flex-direction: column; align-items: flex-start; }
      .filter-bar { flex-direction: column; align-items: stretch; }
      .search-box { max-width: none; }
      .card-header { flex-direction: column; align-items: flex-start; }
      .review-grid { flex-direction: column; align-items: stretch; }
      .btn-group { justify-content: flex-end; }
    }
  </style>
</head>
<body>
  <header class="app-header">
    <div>
      <h1>📋 Hàng đợi xét duyệt hồ sơ</h1>
      <p class="muted">Kiểm tra và phê duyệt các yêu cầu cập nhật thông tin hồ sơ được gửi từ trang tra cứu</p>
    </div>
    <div class="actions">
      <a class="button secondary" href="{{ url_for('admin_audit_log') }}">Nhật ký hoạt động</a>
      {% if current_user and current_user.is_admin %}
      <a class="button secondary" href="{{ url_for('admin_dashboard') }}">Quản trị người dùng</a>
      {% endif %}
      <a class="button secondary" href="{{ url_for('index') }}">← Danh sách hồ sơ</a>
      <a class="button secondary" href="{{ url_for('logout') }}">Đăng xuất</a>
    </div>
  </header>

  {% with messages = get_flashed_messages(with_categories=true) %}
    {% for category, message in messages %}
      <div class="notice {{ category }}">{{ message }}</div>
    {% endfor %}
  {% endwith %}

  <!-- Stats row -->
  <div class="stats-row">
    <div class="stat-card">
      <div class="stat-label">Đang chờ xét duyệt</div>
      <div class="stat-val pending">{{ pending_count }}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Tổng yêu cầu đã ghi nhận</div>
      <div class="stat-val">{{ total_requests }}</div>
    </div>
  </div>

  <!-- Filter tabs and search -->
  <div class="filter-bar">
    <div class="tabs">
      <a class="tab-item {{ 'active' if current_status == 'pending' else '' }}" href="{{ url_for('admin_approvals_queue', status='pending', q=current_query) }}">
        Đang chờ duyệt {% if pending_count %}<span class="tab-badge">{{ pending_count }}</span>{% endif %}
      </a>
      <a class="tab-item {{ 'active' if current_status == 'approved' else '' }}" href="{{ url_for('admin_approvals_queue', status='approved', q=current_query) }}">
        Đã duyệt
      </a>
      <a class="tab-item {{ 'active' if current_status == 'rejected' else '' }}" href="{{ url_for('admin_approvals_queue', status='rejected', q=current_query) }}">
        Đã từ chối
      </a>
      <a class="tab-item {{ 'active' if current_status == 'all' else '' }}" href="{{ url_for('admin_approvals_queue', status='all', q=current_query) }}">
        Tất cả
      </a>
    </div>

    <form method="get" action="{{ url_for('admin_approvals_queue') }}" class="search-box">
      <input type="hidden" name="status" value="{{ current_status }}">
      <input type="text" name="q" value="{{ current_query }}" placeholder="Tìm theo CCCD, Họ tên, STT...">
      <button type="submit">Tìm</button>
      {% if current_query %}
      <a class="button secondary" href="{{ url_for('admin_approvals_queue', status=current_status) }}">Xóa lọc</a>
      {% endif %}
    </form>
  </div>

  <!-- Request List -->
  {% if requests %}
    {% for req in requests %}
    <div class="request-card {{ req.status }}">
      <div class="card-header">
        <div>
          <span class="card-title">Yêu cầu #{{ req.id }} · {{ req.record_name }}</span>
          <span class="muted">(CCCD: <strong>{{ req.cccd }}</strong>{% if req.stt %} · STT: {{ req.stt }}{% endif %})</span>
          <div class="card-meta">
            Gửi lúc: {{ req.submitted_at[:19].replace('T', ' ') }}
          </div>
        </div>
        <div>
          {% if req.status == 'pending' %}
          <span class="badge pending">⏳ Đang chờ duyệt</span>
          {% elif req.status == 'approved' %}
          <span class="badge approved">✓ Đã phê duyệt</span>
          {% elif req.status == 'rejected' %}
          <span class="badge rejected">✗ Đã từ chối</span>
          {% endif %}
        </div>
      </div>

      <div class="card-body">
        <!-- Diff View showing only modified fields -->
        <div class="diff-container">
          <div class="diff-title">
            <span>📝 So sánh thay đổi (Diff View)</span>
            <span class="muted" style="font-weight: normal;">— Chỉ hiển thị {{ req.changes|length }} trường có sự thay đổi</span>
          </div>
          <table class="diff-table">
            <thead>
              <tr>
                <th>Trường thông tin</th>
                <th>Giá trị hiện tại trong hồ sơ (Cũ)</th>
                <th>Giá trị đề xuất cập nhật (Mới)</th>
              </tr>
            </thead>
            <tbody>
              {% for change in req.changes %}
              <tr>
                <td class="diff-field">{{ change.field }}</td>
                <td class="diff-old"><del>{{ change['from'] or '(trống)' }}</del></td>
                <td class="diff-new"><ins>{{ change.to or '(trống)' }}</ins></td>
              </tr>
              {% else %}
              <tr><td colspan="3" class="muted" style="text-align: center;">Không phát hiện thay đổi nào</td></tr>
              {% endfor %}
            </tbody>
          </table>
        </div>
      </div>

      {% if req.status == 'pending' %}
      <!-- Review actions form -->
      <form action="{{ url_for('admin_review_request', request_id=req.id) }}" method="post" class="review-form">
        <input type="hidden" name="csrf_token" value="{{ csrf }}">
        <input type="hidden" name="return_status" value="{{ current_status }}">
        <input type="hidden" name="return_q" value="{{ current_query }}">
        <div class="review-grid">
          <label>
            Ghi chú xét duyệt / Lý do từ chối:
            <input type="text" name="note" placeholder="Nhập ghi chú hoặc lý do từ chối (tùy chọn khi duyệt, nên có khi từ chối)...">
          </label>
          <div class="btn-group">
            <button type="submit" name="action" value="approve" class="button approve-btn" onclick="return confirm('Xác nhận PHÊ DUYỆT yêu cầu này? Thông tin sẽ được cập nhật trực tiếp vào tệp Excel.');">✓ Phê duyệt</button>
            <button type="submit" name="action" value="reject" class="button danger" onclick="return confirm('Xác nhận TỪ CHỐI yêu cầu này?');">✗ Từ chối</button>
          </div>
        </div>
      </form>
      {% else %}
      <!-- Review history -->
      <div class="reviewed-history">
        <strong>Người xử lý:</strong> {{ req.reviewer or '(hệ thống)' }} · 
        <strong>Thời gian:</strong> {{ req.reviewed_at[:19].replace('T', ' ') if req.reviewed_at else '(chưa ghi nhận)' }}
        {% if req.reviewer_note %}
        <div style="margin-top: 4px;">
          <strong>{{ 'Lý do từ chối' if req.status == 'rejected' else 'Ghi chú' }}:</strong> {{ req.reviewer_note }}
        </div>
        {% endif %}
      </div>
      {% endif %}
    </div>
    {% endfor %}

    {% if total_pages > 1 %}
    <nav class="pagination">
      {% for p in range(1, total_pages + 1) %}
      <a class="button secondary {{ 'active' if p == current_page else '' }}" href="{{ url_for('admin_approvals_queue', status=current_status, q=current_query, page=p) }}">{{ p }}</a>
      {% endfor %}
      <span class="page-info">Trang {{ current_page }} / {{ total_pages }} (Tổng {{ total_filtered }} yêu cầu)</span>
    </nav>
    {% endif %}

  {% else %}
    <div class="card empty-queue">
      <div class="empty-queue-icon">✨</div>
      <h3>Không có yêu cầu nào trong hàng đợi</h3>
      <p class="muted">
        {% if current_status == 'pending' %}
        Hiện không có yêu cầu cập nhật nào đang chờ xét duyệt.
        {% else %}
        Không tìm thấy yêu cầu nào phù hợp với bộ lọc hiện tại.
        {% endif %}
      </p>
    </div>
  {% endif %}
</body>
</html>
"""


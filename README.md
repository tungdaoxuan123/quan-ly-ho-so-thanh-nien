# Ứng dụng TSQ A3

Ứng dụng chạy cục bộ để quản lý hồ sơ TSQ A3 từ tệp Excel. Excel là nguồn dữ liệu chính. Ứng dụng đồng bộ dữ liệu sang SQLite để tìm kiếm, lọc và phân trang mà không cần giữ toàn bộ hồ sơ trong RAM. Ứng dụng vẫn hỗ trợ thêm/sửa hồ sơ và tạo tệp Word cho từng người.

## Chạy trên Mac

Mở Terminal, chuyển tới thư mục dự án và chạy một lần:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Sau đó, nhấp đúp [Run TSQ A3 App.command](Run%20TSQ%20A3%20App.command), hoặc chạy:

```bash
.venv/bin/python tsq_a3_app.py
```

Trình duyệt sẽ mở tại `http://127.0.0.1:8765`.

## Chạy trên Windows

1. Trên máy Windows dùng để tạo gói, nhấp đúp `build_windows.bat`.
2. Sao chép toàn bộ thư mục dự án, gồm cả `dist\TSQ_A3`, cho người dùng.
3. Nhấp đúp `Run TSQ A3 App.bat`.
4. Chọn tệp Excel gốc trong ứng dụng.

Tệp thực thi Windows phải được tạo trên Windows. Thư mục `dist\TSQ_A3` chứa ứng dụng và các thư viện cần thiết.

## Cách sử dụng hằng ngày

1. Chọn tệp Excel gốc. Ứng dụng sẽ ghi nhớ đường dẫn và tạo cache SQLite trên máy đang dùng. Lần mở sau sẽ dùng lại cache nếu Excel chưa thay đổi.
2. Dùng ô tìm kiếm để tra cứu theo họ tên, STT, CCCD hoặc nội dung hồ sơ. Mở `Bộ lọc nâng cao` khi cần lọc theo năm sinh, nghề nghiệp, học vấn, dân tộc, tôn giáo hoặc địa chỉ.
3. Chọn `Sửa` để cập nhật hồ sơ, hoặc `Thêm hồ sơ` để tạo người mới. Hồ sơ mới được cấp STT kế tiếp.
4. Lưu và đóng Excel trước khi bấm `Lưu vào Excel` trong ứng dụng. Hệ thống tạo một tệp sao lưu có tên `.backup-YYYYMMDD-HHMMSS.xlsx` cạnh tệp gốc trước khi thay thế tệp.
5. Chọn `Tạo Word`. Tệp được tải về trình duyệt và lưu tại thư mục `Generated TSQ A3` cạnh tệp Excel, với tên `STT - Họ tên.docx`.

Nếu một người khác đã lưu thay đổi trong Excel, bấm `Làm mới`. Ứng dụng cũng tự phát hiện tệp đã thay đổi và đồng bộ lại SQLite. Khi tệp thay đổi trong lúc đang sửa hồ sơ, ứng dụng sẽ yêu cầu mở lại hồ sơ để tránh ghi đè dữ liệu mới.

SQLite chỉ là cache phục vụ ứng dụng. Trên Mac, cache nằm trong `~/Library/Application Support/TSQ_A3_Generator/cache.sqlite3`. Trên Windows, cache nằm trong `%LOCALAPPDATA%\TSQ_A3_Generator\cache.sqlite3`. Có thể xóa tệp cache khi ứng dụng đã tắt; ứng dụng sẽ tạo lại từ Excel vào lần chạy tiếp theo.

## Lưu ý về tệp Excel

- Trang tính đang dùng phải có cột `STT` và `TÊN THƯỜNG DÙNG` ở hàng 1. Dữ liệu hồ sơ bắt đầu từ hàng 3, theo tệp mẫu hiện tại.
- Tệp `.xlsx` hỗ trợ lưu trực tiếp từ ứng dụng. Tệp `.xlsm` chỉ hỗ trợ xem và tạo Word; hãy dùng Excel để chỉnh sửa để tránh ảnh hưởng nội dung VBA.
- Ứng dụng giữ nguyên các trang tính khác, định dạng ô và các cột không chỉnh sửa. Ứng dụng không xóa hồ sơ.
- Không dùng Excel và ứng dụng để lưu cùng lúc. Thay đổi trong Excel chưa lưu sẽ không thể hiển thị trong ứng dụng.
- Tìm kiếm, lọc và phân trang dùng SQLite nên chỉ nạp trang hiện tại vào RAM. Khi lưu thay đổi, thư viện Excel vẫn cần mở toàn bộ workbook; tệp Excel rất lớn có thể dùng nhiều RAM trong lúc lưu.

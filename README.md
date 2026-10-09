# Quản lý hồ sơ thanh niên

Ứng dụng chạy cục bộ để quản lý hồ sơ thanh niên từ tệp Excel. Excel là nguồn dữ liệu chính. Ứng dụng đồng bộ dữ liệu sang SQLite để tìm kiếm, lọc và phân trang mà không cần giữ toàn bộ hồ sơ trong RAM. Ứng dụng hỗ trợ thêm/sửa hồ sơ và tạo tệp Word cho từng người.

## Chạy trên Mac

Mở Terminal, chuyển tới thư mục dự án và chạy một lần:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Sau đó, nhấp đúp [Chạy Quản Lý Hồ Sơ.command](Chay%20Quan%20Ly%20Ho%20So.command), hoặc chạy:

```bash
.venv/bin/python run.py
```

Trình duyệt sẽ mở tại `http://127.0.0.1:8765`.

## Cho máy khác trong mạng nội bộ (LAN) dùng chung

Mặc định ứng dụng chỉ mở trên máy đang chạy (`127.0.0.1`). Để máy khác trong cùng mạng truy cập:

- **Windows:** nhấp đúp `Chay Quan Ly Ho So (LAN).bat`, đặt mật khẩu. Cửa sổ sẽ hiện địa chỉ dạng `http://192.168.x.x:8765`; nếu Windows hỏi về tường lửa, chọn *Private networks* → *Allow access*.
- **Mac / dòng lệnh:** `QLHS_HOST=0.0.0.0 QLHS_PASSWORD=<mật khẩu> .venv/bin/python3 run.py`.

Người dùng mở địa chỉ đó trên trình duyệt và nhập mật khẩu. Ứng dụng từ chối khởi động ở chế độ LAN nếu chưa đặt `QLHS_PASSWORD`, và chỉ nhận kết nối từ địa chỉ mạng nội bộ/riêng (dải `192.168.x.x`, `10.x.x.x`, `172.16–31.x.x`), không nhận từ Internet. Đặt `QLHS_PASSWORD` cũng bắt buộc đăng nhập cả khi chạy trên một máy.

Lưu ý: mọi người dùng chung một tệp Excel và một cơ sở dữ liệu trên máy chạy ứng dụng, chưa có phân quyền theo người dùng, và chưa có cơ chế khóa khi nhiều người lưu cùng lúc. Các nút "Mở bằng Excel" / "Đổi tệp Excel" tác động lên máy chạy ứng dụng.

## Chạy trên Windows

1. Trên máy Windows dùng để tạo gói, nhấp đúp `build_windows.bat`.
2. Sao chép toàn bộ thư mục dự án, gồm cả `dist\QuanLyHoSoThanhNien`, cho người dùng.
3. Nhấp đúp `Chay Quan Ly Ho So.bat`.
4. Chọn tệp Excel gốc trong ứng dụng.

Tệp thực thi Windows phải được tạo trên Windows. Thư mục `dist\QuanLyHoSoThanhNien` chứa ứng dụng và các thư viện cần thiết.

## Cách sử dụng hằng ngày

1. Chọn tệp Excel gốc. Ứng dụng sẽ ghi nhớ đường dẫn và tạo cache SQLite trên máy đang dùng. Lần mở sau sẽ dùng lại cache nếu Excel chưa thay đổi.
2. Dùng ô tìm kiếm để tra cứu theo họ tên, STT, CCCD hoặc nội dung hồ sơ. Mở `Bộ lọc nâng cao` khi cần lọc theo năm sinh, nghề nghiệp, học vấn, dân tộc, tôn giáo hoặc địa chỉ.
3. Chọn `Sửa` để cập nhật hồ sơ, hoặc `Thêm hồ sơ` để tạo người mới. Hồ sơ mới được cấp STT kế tiếp.
4. Lưu và đóng Excel trước khi bấm `Lưu vào Excel` trong ứng dụng. Hệ thống tạo một tệp sao lưu có tên `.backup-YYYYMMDD-HHMMSS.xlsx` cạnh tệp gốc trước khi thay thế tệp.
5. Chọn `Tạo Word`. Tệp được tải về trình duyệt và lưu tại thư mục `Hồ sơ thanh niên đã tạo` cạnh tệp Excel, với tên `STT - Họ tên.docx`.

Nếu một người khác đã lưu thay đổi trong Excel, bấm `Làm mới`. Ứng dụng cũng tự phát hiện tệp đã thay đổi và đồng bộ lại SQLite. Khi tệp thay đổi trong lúc đang sửa hồ sơ, ứng dụng sẽ yêu cầu mở lại hồ sơ để tránh ghi đè dữ liệu mới.

Cơ sở dữ liệu SQLite `cache.sqlite3` và `settings.json` nằm ngay trong thư mục của ứng dụng (cạnh `run.py`); nếu thư mục đó không ghi được, ứng dụng dùng thư mục dữ liệu người dùng (`~/Library/Application Support/Quan_Ly_Ho_So_Thanh_Nien` trên Mac, `%LOCALAPPDATA%\Quan_Ly_Ho_So_Thanh_Nien` trên Windows). Danh sách hồ sơ trong đó chỉ là cache dựng lại từ Excel, nhưng **Diện, Khu phố và tài liệu liên quan chỉ được lưu ở đây, không có trong Excel** — hãy sao lưu `cache.sqlite3` cùng tệp Excel. Xóa `cache.sqlite3` sẽ mất các dữ liệu đó; phần danh sách hồ sơ vẫn được tạo lại từ Excel vào lần chạy tiếp theo.

## Lưu ý về tệp Excel

- Trang tính đang dùng phải có cột `STT` và `TÊN THƯỜNG DÙNG` ở hàng 1. Dữ liệu hồ sơ bắt đầu từ hàng 3, theo tệp mẫu hiện tại.
- Tệp `.xlsx` hỗ trợ lưu trực tiếp từ ứng dụng. Tệp `.xlsm` chỉ hỗ trợ xem và tạo Word; hãy dùng Excel để chỉnh sửa để tránh ảnh hưởng nội dung VBA.
- Ứng dụng giữ nguyên các trang tính khác, định dạng ô và các cột không chỉnh sửa. Ứng dụng không xóa hồ sơ.
- Không dùng Excel và ứng dụng để lưu cùng lúc. Thay đổi trong Excel chưa lưu sẽ không thể hiển thị trong ứng dụng.
- Tìm kiếm, lọc và phân trang dùng SQLite nên chỉ nạp trang hiện tại vào RAM. Khi lưu thay đổi, thư viện Excel vẫn cần mở toàn bộ workbook; tệp Excel rất lớn có thể dùng nhiều RAM trong lúc lưu.

## Dành cho lập trình viên

Mã nguồn nằm trong `src/quan_ly_ho_so/` (src-layout), chia theo chức năng: `config.py`/`state.py`/`errors.py` (cấu hình và trạng thái dùng chung), `workbook/` (đọc/ghi Excel và cache SQLite), `forms/` (định nghĩa và kiểm tra trường biểu mẫu), `web/` (giao diện HTML), `word/` (tạo tệp Word), `app.py` (Flask app và route). `run.py` ở thư mục gốc chỉ thêm `src/` vào đường dẫn rồi khởi động ứng dụng — đây là tệp mà các launcher (`.command`/`.bat`) và `build_windows.bat` gọi tới.

Chạy bộ kiểm thử:

```bash
.venv/bin/python -m unittest discover -s tests -t . -v
```

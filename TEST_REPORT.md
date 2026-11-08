# Kết quả kiểm thử Windows — 08/11/2026

## Kết quả

- Bộ kiểm thử mới nhất: **78/78 đạt**, 22,731 giây.
- Bản ZIP offline sau khi build lại: **8 nhóm kiểm tra HTTP thực tế đạt**.
- Dữ liệu kiểm thử được tạo trong thư mục tạm; không dùng hồ sơ thật.

## Phạm vi đã kiểm tra

| Nhóm | Nội dung |
|---|---|
| Tài khoản | Đăng nhập, đăng xuất, sai mật khẩu, khóa đăng nhập, đổi/reset mật khẩu, buộc đổi mật khẩu, quản trị người dùng |
| Phân quyền | Đọc, tạo, sửa, xóa, xuất; tài khoản chỉ đọc; giới hạn truy cập trang quản trị |
| Excel | Đọc, thêm/sửa, bảo toàn cấu trúc, đồng bộ thay đổi bên ngoài, lỗi đồng bộ không mất cache, chống lưu biểu mẫu cũ |
| Tìm kiếm/lọc | Không dấu, hoa/thường, nhiều giá trị, ngày thay đổi, phân trang danh sách lớn |
| Hồ sơ | Kiểm tra CCCD, xóa mềm/khôi phục, Diện/Khu phố, nhập danh sách CA và tránh trùng |
| Duyệt thay đổi | Khách gửi yêu cầu, chống spam, chặn đổi thông tin định danh, duyệt/từ chối, xung đột sửa đồng thời, đổi STT, đổi thứ tự cột, sai nguồn Excel |
| Nhật ký | Ghi lịch sử thao tác, lọc người dùng/hành động/ngày, tìm kiếm, phân trang |
| Tài liệu | Xuất Word, xuất danh sách Excel, tệp đính kèm, luồng chuyển Word sang PDF với bộ chuyển đổi được giả lập |
| LAN | Sinh/gia hạn chứng chỉ khi đổi IP/tên máy, giữ CA, máy khách TLS im lặng không làm treo máy chủ |
| Gói offline | ZIP nguyên vẹn, đủ mẫu và thư viện, không chứa DB/uploads/khóa riêng, khởi động EXE từ thư mục làm việc khác |
| HTTP bản EXE | Khách tra cứu, chặn trang quản trị, đăng nhập, upload Excel, tìm kiếm/lọc, trang thêm/sửa/xem trước, profile, duyệt, nhật ký, status, xuất DOCX thực, từ chối CSRF sai, đăng xuất |

## Điều chỉnh sau kiểm thử

1. Cấu hình UTF-8 cho đầu ra launcher Python để tiếng Việt không lỗi khi ghi ra tệp trên Windows.
2. Kiểm tra quyền POSIX `0600` chỉ chạy trên hệ điều hành hỗ trợ; Windows dùng ACL, chưa kiểm toán ACL trong đợt này.
3. Sửa cấu hình build để mẫu Word/Excel và thư mục `template` nằm cạnh EXE, đúng đường dẫn ứng dụng đọc. Dùng console để hỗ trợ nhập mật khẩu khi mở LAN và xem lỗi.
4. Bổ sung `tests/check_offline_package.py` để tái kiểm tra ZIP qua HTTP thực, gồm xuất DOCX bằng mẫu đã đóng gói.
5. Hộp chọn Excel trên Windows chạy trong tiến trình riêng, với Tkinter trên luồng chính; bổ sung kiểm tra đường dẫn Unicode, hủy, lỗi và chế độ đóng gói.
6. Chuẩn hóa ký tự xuống dòng LF/CRLF khi so sánh thay đổi và kiểm tra xung đột duyệt hồ sơ. Kiểm thử xác nhận ghi chú không sửa không xuất hiện trong diff và giữ nguyên sau khi duyệt.

## Giới hạn

Đây là kiểm thử tự động ở lớp nghiệp vụ/HTTP trên máy hiện tại. Chưa thao tác trực quan mọi nút bằng trình duyệt; chưa thử máy Windows sạch khác, điện thoại qua LAN, tường lửa thực, chứng chỉ trên iPhone/Android, in máy in, hoặc Word/PDF thông qua Microsoft Word thực. Chưa ngắt mạng của máy chủ để kiểm tra offline; EXE được chạy độc lập từ bản giải nén, yêu cầu HTTP chỉ gửi tới localhost. Các bài test có giả lập tích hợp bên ngoài không chứng minh phần mềm bên ngoài đã được cài trên máy đích.

## Chạy lại

```powershell
$env:PYTHONUTF8='1'
.\.venv\Scripts\python.exe -m unittest discover -s tests -t . -v
.\.venv\Scripts\python.exe tests\check_offline_package.py
```

Nhật ký chi tiết tại `build/test-full.log` và `build/test-offline.log` (không đưa dữ liệu sinh ra vào Git).

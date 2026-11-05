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

Trình duyệt sẽ mở tại `http://127.0.0.1:3000`.

## Tài khoản, phân quyền và dùng chung trong mạng nội bộ (LAN)

Ứng dụng yêu cầu đăng nhập. Lần đầu chạy tạo tài khoản quản trị `admin` (mật khẩu mặc định `admin123` — **hãy đổi ngay**; xem `.env.example`). Quản trị viên tạo người dùng tại **Quản trị**, chọn nhanh mẫu quyền hoặc tích từng quyền:

- **Toàn quyền:** tạo, xem, sửa, xóa, xuất file.
- **Chỉ xem và in:** chỉ quyền *Xem* — xem hồ sơ, xem trước, tạo Word (in); không xuất Excel, không sửa.
- **Cập nhật:** *Tạo + Xem + Sửa*. Quyền *Xuất file* (Excel hàng loạt, Excel theo diện) và *Xóa* tích riêng.

**Nhập file CA** (nút *Nhập file CA*): đọc các tệp khu phố của bộ CA và cho xem trước, chưa ghi gì. Khi xác nhận, theo CCCD: người chưa có được thêm mới; người đã có được cập nhật các ô mà file CA có dữ liệu và khác hiện tại (ô trống trong file CA không xóa dữ liệu đang có, Khu phố lấy từ cột *Khu phố mới*). Tất cả ghi một lần, có bản sao lưu Excel trước đó và có nhật ký từng người (giá trị cũ → mới).

Mọi thao tác ghi (sửa, thêm, xóa/khôi phục, sửa hàng loạt, nhập file CA, tải tài liệu, xuất Excel) được ghi vào nhật ký theo người dùng; quản trị viên xem tại **Nhật ký**. Quyền được nạp lúc đăng nhập: sau khi đổi quyền, người đó cần đăng nhập lại.

Khách (không cần tài khoản) vào `/tra-cuu`, nhập số CCCD: nếu chưa có thì điền form tạo hồ sơ mới, nếu đã có thì chỉ được xem — khách không sửa được.

Mặc định ứng dụng chỉ mở trên máy đang chạy (`127.0.0.1`). Để máy khác trong cùng mạng truy cập:

- **macOS:** nhấp đúp `Chay Quan Ly Ho So (LAN).command`.
- **Windows:** nhấp đúp `Chay Quan Ly Ho So (LAN).bat`. Cửa sổ sẽ hiện địa chỉ dạng `https://192.168.x.x:3000`; nếu Windows hỏi về tường lửa, chọn *Private networks* → *Allow access*.
- **Mac / dòng lệnh:** `QLHS_HOST=0.0.0.0 .venv/bin/python3 run.py`.

Nếu tài khoản quản trị vẫn dùng mật khẩu mặc định, ứng dụng sẽ hỏi đặt mật khẩu mới (ít nhất 8 ký tự) ngay trong cửa sổ dòng lệnh trước khi mở cho người khác. Ở chế độ LAN, ứng dụng chỉ nhận kết nối từ địa chỉ mạng nội bộ/riêng (`192.168.x.x`, `10.x.x.x`, `172.16–31.x.x`), không nhận từ Internet.

**Địa chỉ cố định:** cửa sổ ứng dụng hiện địa chỉ dạng `https://ten-may.local:3000` (tên máy, macOS xem ở *Cài đặt hệ thống* → *Chung* → *Chia sẻ* → *Tên máy chủ cục bộ*). Địa chỉ này không đổi khi tắt máy hay kết nối lại wifi, khác với địa chỉ IP `192.168.x.x`. Thiết bị nào không mở được tên `.local` (một số máy Android cũ) thì dùng địa chỉ IP hiện ngay bên dưới.

**HTTPS trong LAN:** chế độ LAN chạy `https://` (Chrome trên điện thoại hay tự chuyển sang https và không mở được http). Lần chạy đầu ứng dụng tự tạo chứng chỉ trong thư mục `HTTPS/` (không đưa lên git; `ca-key.pem` và `key.pem` là khóa bí mật, không gửi cho ai). Lần đầu mỗi thiết bị sẽ báo *kết nối không riêng tư*: bấm *Nâng cao* → *Tiếp tục*. Muốn hết hẳn cảnh báo, cài chứng chỉ một lần trên từng thiết bị — tải tại `https://192.168.x.x:3000/chung-chi`:

- **iPhone:** mở địa chỉ trên bằng **Safari** → *Cho phép* tải hồ sơ → *Cài đặt* → *Đã tải về hồ sơ* → *Cài đặt*; rồi *Cài đặt* → *Cài đặt chung* → *Giới thiệu* → *Cài đặt tin cậy chứng nhận* → bật *Quan Ly Ho So - CA noi bo*. Chrome trên iPhone còn cần bật *Cài đặt* → *Chrome* → *Mạng cục bộ*.
- **Android:** tải tệp `QuanLyHoSo-CA.crt`, rồi *Cài đặt* → *Bảo mật* → *Mã hóa và thông tin xác thực* → *Cài đặt chứng chỉ* → *Chứng chỉ CA* → chọn tệp vừa tải.

Chứng chỉ này chỉ có hiệu lực cho địa chỉ mạng nội bộ, không dùng được để giả mạo trang web khác. Đổi địa chỉ IP thì ứng dụng tự cấp lại, không cần cài lại. Xóa thư mục `HTTPS/` thì phải cài lại trên mọi thiết bị.

### Máy Windows không có mạng (một nút bấm)

Chuẩn bị trên một máy có mạng (Mac/Linux, trong thư mục mã nguồn): chạy `./build_windows_offline.sh`. Lệnh này tạo `dist/QuanLyHoSo-Windows.zip` (~40 MB) gồm Python cho Windows, đủ thư viện và ứng dụng — không chứa dữ liệu Excel hay tài khoản của máy đang chạy.

Trên máy Windows 10/11 không có mạng:

1. Chép `QuanLyHoSo-Windows.zip` sang (USB), bấm chuột phải → *Extract All*. Chép file Excel tổng vào máy.
2. **Một lần duy nhất:** bấm chuột phải `Mo tuong lua Windows (Run as administrator).bat` → *Run as administrator*. Wifi không có internet thường bị Windows coi là mạng *Public* nên nếu bỏ bước này điện thoại sẽ không vào được.
3. Bấm chuột phải `Chay Quan Ly Ho So (LAN).bat` → *Send to* → *Desktop (create shortcut)*. Từ giờ chỉ cần nhấp đúp biểu tượng này trên Desktop. Lần đầu cửa sổ sẽ hỏi đặt mật khẩu admin mới.
4. Cửa sổ hiện địa chỉ dạng `https://192.168.x.x:3000` — điện thoại dùng địa chỉ IP này (không dùng tên `.local` trên Android). Đóng cửa sổ là tắt ứng dụng.

**Giữ địa chỉ IP cố định** để không phải báo địa chỉ mới cho mọi người — chọn một trong hai:

- **Một nút trên Windows:** nhấp đúp `Dat IP co dinh.bat` (bấm *Yes* khi Windows hỏi quyền). Máy giữ luôn IP đang có và hiện địa chỉ cho điện thoại. Muốn bỏ (ví dụ mang máy sang wifi khác): `Tra lai IP tu dong.bat`. Lưu ý: router không biết IP này đã có chủ, nên hiếm khi có thể cấp trùng cho thiết bị khác; nếu bị trùng, dùng cách dưới.
- **Giữ IP trên router (chắc chắn nhất):** vào trang quản trị router (thường `http://192.168.1.1`) → *DHCP Reservation / Static Lease / IP tĩnh* → gán cố định một IP cho máy Windows.

Tên máy (`https://<tên-máy>.local:3000`, hiện trong cửa sổ ứng dụng) không đổi khi IP đổi và dùng được trên iPhone, máy tính; Chrome trên Android thường không tìm được tên `.local`, nên điện thoại Android dùng địa chỉ IP.

### Điện thoại Android: Chrome không vào được (trình duyệt khác vào được)

Ứng dụng đã chạy https; khi Chrome báo *không thể truy cập trang web* thì điện thoại đang chặn Chrome, không phải do máy chủ. Thử nhanh: mở `http://192.168.1.1` (trang router) bằng Chrome — nếu cũng không vào được thì làm lần lượt:

1. **Wifi không có internet:** Android thấy wifi không có mạng sẽ tự gửi dữ liệu qua 4G, nên không tới được máy trong wifi. Khi điện thoại báo *Wi-Fi không có kết nối Internet*, bấm vào → *Vẫn kết nối* / *Không hỏi lại*. Tắt *Tự động chuyển sang dữ liệu di động* (Samsung: *Cài đặt → Kết nối → Wi-Fi → ⋮ → Intelligent Wi-Fi*; máy khác: *Wi-Fi → Nâng cao*). Cách đơn giản nhất: **tắt dữ liệu di động** khi dùng ứng dụng.
2. **Quyền mạng cục bộ của Chrome:** *Cài đặt → Ứng dụng → Chrome → Quyền* → cho phép *Thiết bị ở gần* / *Mạng cục bộ* (nếu có). Khi Chrome hỏi có cho truy cập thiết bị trong mạng cục bộ không, chọn *Cho phép*.
3. **Chặn mạng theo ứng dụng** (Xiaomi, Oppo, Realme, Vivo…): *Cài đặt → Ứng dụng → Chrome → Dữ liệu di động & Wi-Fi* → cho phép Wi-Fi. Tắt VPN, trình chặn quảng cáo, và *Cài đặt → Mạng → DNS riêng tư* đặt về *Tắt/Tự động*.
4. Mở được rồi mà Chrome báo *kết nối không riêng tư*: bấm *Nâng cao → Tiếp tục*, hoặc cài chứng chỉ tại `https://<địa chỉ máy>:3000/chung-chi` để hết cảnh báo.

### Đưa ra Internet (một nút)

Wifi công ty thường chặn các máy nói chuyện với nhau, hoặc người dùng ở ngoài mạng. Khi đó dùng `Chay Online.command` (macOS) hoặc `Chay Online.bat` (Windows). Cần cài `cloudflared` một lần (`brew install cloudflared` hoặc `winget install Cloudflare.cloudflared`).

File này sẽ: hỏi đặt mật khẩu admin mới nếu còn mật khẩu mặc định, bật ứng dụng (chỉ nghe ở `127.0.0.1`), rồi mở đường hầm Cloudflare và in địa chỉ `https://….trycloudflare.com` để gửi cho mọi người (khách tra cứu thêm `/tra-cuu`). Địa chỉ đổi mỗi lần bật; đóng cửa sổ là tắt truy cập từ ngoài. Máy phải đang mở.

Chế độ online (`QLHS_ONLINE=1`) tắt giới hạn mạng nội bộ nên bảo vệ bằng: bắt buộc đăng nhập, khóa 15 phút sau 5 lần sai mật khẩu, giới hạn 30 lượt tra cứu CCCD/10 phút và 5 hồ sơ khách tạo/giờ cho mỗi địa chỉ, cookie phiên chỉ gửi qua https.

Lưu ý: mọi người dùng chung một tệp Excel và một cơ sở dữ liệu trên máy chạy ứng dụng; chưa có cơ chế khóa khi nhiều người lưu cùng lúc. Các nút "Mở bằng Excel" / "Đổi tệp Excel" tác động lên máy chạy ứng dụng.

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

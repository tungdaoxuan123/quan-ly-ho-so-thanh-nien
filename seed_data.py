"""Seed script: populates sample youth records, demo user accounts, and initial audit logs."""

from __future__ import annotations

import sys
from pathlib import Path

# Add src to python path
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from openpyxl import Workbook, load_workbook

from quan_ly_ho_so.auth.db import auth_database_connection, create_user
from quan_ly_ho_so.audit import log_action
from quan_ly_ho_so.config import FORM_GROUPS, DELETED_AT_HEADER
from quan_ly_ho_so.workbook.manager import set_workbook

WORKBOOK_PATH = Path(__file__).resolve().parent / "Danh_Sach_Thanh_Nien.xlsx"

SAMPLE_YOUTHS = [
    {
        "STT": 1,
        "TÊN THƯỜNG DÙNG": "NGUYỄN VĂN AN",
        "Tên khai sinh": "Nguyễn Văn An",
        "Ngày sinh": 15,
        "Tháng sinh": 5,
        "Năm sinh": 2005,
        "Căn cước": "001205000001",
        "Nơi đăng ký khai sinh (2 cấp)": "Phường Hàng Bông, Quận Hoàn Kiếm, TP Hà Nội",
        "Quê quán (2 cấp)": "Xã Kim Chung, Huyện Hoài Đức, TP Hà Nội",
        "Dân tộc": "Kinh",
        "Tôn giáo": "Không",
        "Thường trú": "Số 12 phố Tràng Thi, Phường Hàng Trống, Quận Hoàn Kiếm, TP Hà Nội",
        "Nơi ở hiện nay": "Số 12 phố Tràng Thi, Phường Hàng Trống, Quận Hoàn Kiếm, TP Hà Nội",
        "Thành phần": "Cán bộ công chức",
        "Bản thân": "Sinh viên",
        "Văn hóa": "12/12",
        "Chuyên môn": "Đại học",
        "Năm tốt nghiệp": 2027,
        "Ngành đào tạo": "Công nghệ thông tin",
        "Nơi làm việc, học tập": "Trường Đại học Bách Khoa Hà Nội",
        "Tên cha": "Nguyễn Văn Hùng",
        "Sống / Chết": "Sống",
        "Năm sinh cha": 1975,
        "Nghề nghiệp cha": "Kỹ sư xây dựng",
        "Tên mẹ": "Trần Thị Mai",
        "Sống / chết (mẹ)": "Sống",
        "Năm sinh mẹ": 1978,
        "Nghề nghiệp mẹ": "Giáo viên",
        "cha mẹ có bao nhiêu con": 2,
        "mấy trai": 1,
        "mấy gái": 1,
        "là Con thứ": 1,
        "Cấp 1": "2011-2016: Tiểu học Thăng Long, Hà Nội",
        "Cấp 2": "2016-2020: THCS Trưng Vương, Hà Nội",
        "Cấp 3": "2020-2023: THPT Việt Đức, Hà Nội",
        "Hiện nay": "2023-nay: Sinh viên Đại học Bách Khoa Hà Nội",
    },
    {
        "STT": 2,
        "TÊN THƯỜNG DÙNG": "TRẦN VĂN BÌNH",
        "Tên khai sinh": "Trần Văn Bình",
        "Ngày sinh": 22,
        "Tháng sinh": 8,
        "Năm sinh": 2004,
        "Căn cước": "048204000002",
        "Nơi đăng ký khai sinh (2 cấp)": "Phường Hải Châu 1, Quận Hải Châu, TP Đà Nẵng",
        "Quê quán (2 cấp)": "Huyện Điện Bàn, Tỉnh Quảng Nam",
        "Dân tộc": "Kinh",
        "Tôn giáo": "Không",
        "Thường trú": "Số 45 đường Lê Lợi, Phường Thạch Thang, Quận Hải Châu, TP Đà Nẵng",
        "Nơi ở hiện nay": "Số 45 đường Lê Lợi, Phường Thạch Thang, Quận Hải Châu, TP Đà Nẵng",
        "Thành phần": "Tiểu thương",
        "Bản thân": "Thợ điện tử",
        "Văn hóa": "12/12",
        "Chuyên môn": "Cao đẳng nghề",
        "Năm tốt nghiệp": 2025,
        "Ngành đào tạo": "Điện tử công nghiệp",
        "Nơi làm việc, học tập": "Công ty TNHH Cơ điện Miền Trung",
        "Tên cha": "Trần Văn Đức",
        "Sống / Chết": "Sống",
        "Năm sinh cha": 1972,
        "Nghề nghiệp cha": "Lái xe",
        "Tên mẹ": "Lê Thị Thảo",
        "Sống / chết (mẹ)": "Sống",
        "Năm sinh mẹ": 1976,
        "Nghề nghiệp mẹ": "Kinh doanh tự do",
        "cha mẹ có bao nhiêu con": 3,
        "mấy trai": 2,
        "mấy gái": 1,
        "là Con thứ": 2,
        "Cấp 1": "2010-2015: Tiểu học Trần Cao Vân, Đà Nẵng",
        "Cấp 2": "2015-2019: THCS Nguyễn Huệ, Đà Nẵng",
        "Cấp 3": "2019-2022: THPT Phan Châu Trinh, Đà Nẵng",
        "Hiện nay": "2022-nay: Làm việc tại Công ty Cơ điện Miền Trung",
    },
    {
        "STT": 3,
        "TÊN THƯỜNG DÙNG": "LÊ HOÀNG CƯỜNG",
        "Tên khai sinh": "Lê Hoàng Cường",
        "Ngày sinh": 10,
        "Tháng sinh": 11,
        "Năm sinh": 2003,
        "Căn cước": "079203000003",
        "Nơi đăng ký khai sinh (2 cấp)": "Phường Bến Nghé, Quận 1, TP Hồ Chí Minh",
        "Quê quán (2 cấp)": "Huyện Củ Chi, TP Hồ Chí Minh",
        "Dân tộc": "Kinh",
        "Tôn giáo": "Không",
        "Thường trú": "Số 88 đường Nguyễn Du, Phường Bến Nghé, Quận 1, TP Hồ Chí Minh",
        "Nơi ở hiện nay": "Số 88 đường Nguyễn Du, Phường Bến Nghé, Quận 1, TP Hồ Chí Minh",
        "Thành phần": "Viên chức",
        "Bản thân": "Kỹ sư phần mềm",
        "Văn hóa": "Đại học",
        "Chuyên môn": "Kỹ sư",
        "Năm tốt nghiệp": 2025,
        "Ngành đào tạo": "Khoa học máy tính",
        "Nơi làm việc, học tập": "Tập đoàn Viễn thông FPT",
        "Tên cha": "Lê Hoàng Dũng",
        "Sống / Chết": "Sống",
        "Năm sinh cha": 1970,
        "Nghề nghiệp cha": "Bác sĩ",
        "Tên mẹ": "Nguyễn Thị Ngọc",
        "Sống / chết (mẹ)": "Sống",
        "Năm sinh mẹ": 1974,
        "Nghề nghiệp mẹ": "Dược sĩ",
        "cha mẹ có bao nhiêu con": 1,
        "mấy trai": 1,
        "mấy gái": 0,
        "là Con thứ": 1,
        "Cấp 1": "2009-2014: Tiểu học Nguyễn Bỉnh Khiêm, Quận 1",
        "Cấp 2": "2014-2018: THCS Lê Quý Đôn, Quận 3",
        "Cấp 3": "2018-2021: THPT Chuyên Lê Hồng Phong, TP HCM",
        "Hiện nay": "2021-nay: Làm việc tại FPT",
    },
    {
        "STT": 4,
        "TÊN THƯỜNG DÙNG": "PHẠM MINH ĐỨC",
        "Tên khai sinh": "Phạm Minh Đức",
        "Ngày sinh": 3,
        "Tháng sinh": 2,
        "Năm sinh": 2006,
        "Căn cước": "001206000004",
        "Nơi đăng ký khai sinh (2 cấp)": "Phường Láng Thượng, Quận Đống Đa, TP Hà Nội",
        "Quê quán (2 cấp)": "Huyện Nam Trực, Tỉnh Nam Định",
        "Dân tộc": "Kinh",
        "Tôn giáo": "Không",
        "Thường trú": "Số 25 ngõ 1194 đường Láng, Đống Đa, TP Hà Nội",
        "Nơi ở hiện nay": "Số 25 ngõ 1194 đường Láng, Đống Đa, TP Hà Nội",
        "Thành phần": "Lao động",
        "Bản thân": "Học sinh",
        "Văn hóa": "12/12",
        "Năm tốt nghiệp": 2024,
        "Nơi làm việc, học tập": "Đang chờ nhập ngũ NVQS",
        "Tên cha": "Phạm Văn Thành",
        "Sống / Chết": "Sống",
        "Năm sinh cha": 1977,
        "Nghề nghiệp cha": "Công nhân cơ khí",
        "Tên mẹ": "Vũ Thị Hằng",
        "Sống / chết (mẹ)": "Sống",
        "Năm sinh mẹ": 1980,
        "Nghề nghiệp mẹ": "May mặc",
        "cha mẹ có bao nhiêu con": 2,
        "mấy trai": 2,
        "mấy gái": 0,
        "là Con thứ": 1,
        "Cấp 1": "2012-2017: Tiểu học Láng Thượng",
        "Cấp 2": "2017-2021: THCS Cầu Giấy",
        "Cấp 3": "2021-2024: THPT Cầu Giấy",
        "Hiện nay": "Sẵn sàng thi hành NVQS",
    },
    {
        "STT": 5,
        "TÊN THƯỜNG DÙNG": "VÕ THÀNH ĐẠT",
        "Tên khai sinh": "Võ Thành Đạt",
        "Ngày sinh": 18,
        "Tháng sinh": 9,
        "Năm sinh": 2005,
        "Căn cước": "092205000005",
        "Nơi đăng ký khai sinh (2 cấp)": "Phường An Khánh, Quận Ninh Kiều, TP Cần Thơ",
        "Quê quán (2 cấp)": "Huyện Châu Thành, Tỉnh Hậu Giang",
        "Dân tộc": "Kinh",
        "Tôn giáo": "Không",
        "Thường trú": "Số 67 đường 30 Tháng 4, Phường Hưng Lợi, Ninh Kiều, Cần Thơ",
        "Nơi ở hiện nay": "Số 67 đường 30 Tháng 4, Phường Hưng Lợi, Ninh Kiều, Cần Thơ",
        "Thành phần": "Nông dân",
        "Bản thân": "Sinh viên",
        "Văn hóa": "12/12",
        "Chuyên môn": "Đại học",
        "Năm tốt nghiệp": 2027,
        "Ngành đào tạo": "Nông nghiệp công nghệ cao",
        "Nơi làm việc, học tập": "Đại học Cần Thơ",
        "Tên cha": "Võ Văn Sang",
        "Sống / Chết": "Sống",
        "Năm sinh cha": 1973,
        "Nghề nghiệp cha": "Làm vườn",
        "Tên mẹ": "Đặng Thị Hoa",
        "Sống / chết (mẹ)": "Sống",
        "Năm sinh mẹ": 1977,
        "Nghề nghiệp mẹ": "Nội trợ",
        "cha mẹ có bao nhiêu con": 2,
        "mấy trai": 1,
        "mấy gái": 1,
        "là Con thứ": 1,
        "Cấp 1": "2011-2016: Tiểu học Mạc Đĩnh Chi, Cần Thơ",
        "Cấp 2": "2016-2020: THCS Đoàn Thị Điểm, Cần Thơ",
        "Cấp 3": "2020-2023: THPT Châu Văn Liêm, Cần Thơ",
        "Hiện nay": "Sinh viên Đại học Cần Thơ",
    },
    {
        "STT": 6,
        "TÊN THƯỜNG DÙNG": "HOÀNG QUỐC VIỆT",
        "Tên khai sinh": "Hoàng Quốc Việt",
        "Ngày sinh": 28,
        "Tháng sinh": 12,
        "Năm sinh": 2002,
        "Căn cước": "031202000006",
        "Nơi đăng ký khai sinh (2 cấp)": "Phường Ngô Quyền, TP Hải Phòng",
        "Quê quán (2 cấp)": "Huyện Thủy Nguyên, TP Hải Phòng",
        "Dân tộc": "Kinh",
        "Tôn giáo": "Không",
        "Thường trú": "Số 102 đường Lạch Tray, Quận Ngô Quyền, TP Hải Phòng",
        "Nơi ở hiện nay": "Số 102 đường Lạch Tray, Quận Ngô Quyền, TP Hải Phòng",
        "Thành phần": "Công nhân",
        "Bản thân": "Nhân viên cảng",
        "Văn hóa": "12/12",
        "Chuyên môn": "Cao đẳng",
        "Năm tốt nghiệp": 2023,
        "Ngành đào tạo": "Khai thác cảng",
        "Nơi làm việc, học tập": "Cảng Hải Phòng",
        "Tên cha": "Hoàng Văn Hậu",
        "Sống / Chết": "Sống",
        "Năm sinh cha": 1968,
        "Nghề nghiệp cha": "Công nhân cảng",
        "Tên mẹ": "Bùi Thị Quyên",
        "Sống / chết (mẹ)": "Sống",
        "Năm sinh mẹ": 1971,
        "Nghề nghiệp mẹ": "Y tá",
        "cha mẹ có bao nhiêu con": 2,
        "mấy trai": 2,
        "mấy gái": 0,
        "là Con thứ": 2,
        "Cấp 1": "2008-2013: Tiểu học Lê Hồng Phong, Hải Phòng",
        "Cấp 2": "2013-2017: THCS Chu Văn An, Hải Phòng",
        "Cấp 3": "2017-2020: THPT Thái Phiên, Hải Phòng",
        "Hiện nay": "Làm việc tại Cảng Hải Phòng",
        "Đã xóa lúc": "2026-10-07T10:00:00",  # Sample soft-deleted record!
    },
]

DEMO_USERS = [
    {
        "username": "canbo_nhaplieu",
        "password": "user1234",
        "can_create": True,
        "can_read": True,
        "can_update": True,
        "can_delete": False,
        "must_change_password": False,
    },
    {
        "username": "canbo_kiemtra",
        "password": "user1234",
        "can_create": False,
        "can_read": True,
        "can_update": False,
        "can_delete": False,
        "must_change_password": False,
    },
    {
        "username": "canbo_toanquyen",
        "password": "user1234",
        "can_create": True,
        "can_read": True,
        "can_update": True,
        "can_delete": True,
        "must_change_password": False,
    },
]


def seed_excel():
    print(f"Creating Excel workbook at: {WORKBOOK_PATH}")
    headers = ["STT"]
    for _, section_headers in FORM_GROUPS:
        for header in section_headers:
            if header not in headers:
                headers.append(header)
    if DELETED_AT_HEADER not in headers:
        headers.append(DELETED_AT_HEADER)

    wb = Workbook()
    sheet = wb.active
    sheet.title = "Hồ sơ"
    sheet.append(headers)
    sheet.append([])  # Blank row 2

    for record in SAMPLE_YOUTHS:
        row = [record.get(h) for h in headers]
        sheet.append(row)

    wb.save(WORKBOOK_PATH)
    wb.close()
    print(f"-> Successfully seeded {len(SAMPLE_YOUTHS)} youth records into {WORKBOOK_PATH.name}")


def seed_users():
    print("Seeding demo user accounts into users.db...")
    conn = auth_database_connection()
    conn.close()

    for user_info in DEMO_USERS:
        try:
            u = create_user(
                username=user_info["username"],
                password=user_info["password"],
                can_create=user_info["can_create"],
                can_read=user_info["can_read"],
                can_update=user_info["can_update"],
                can_delete=user_info["can_delete"],
                must_change_password=user_info["must_change_password"],
            )
            print(f"-> Created user '{u['username']}' with permissions: "
                  f"create={u['can_create']}, read={u['can_read']}, update={u['can_update']}, delete={u['can_delete']}")
        except ValueError as e:
            print(f"-> User '{user_info['username']}' already exists ({e})")


def seed_audit_logs():
    print("Seeding initial CRUD audit logs...")
    log_action("admin", "create", "1", "NGUYỄN VĂN AN", {
        "fields": {"TÊN THƯỜNG DÙNG": "NGUYỄN VĂN AN", "Năm sinh": 2005, "Căn cước": "001205000001"}
    })
    log_action("canbo_nhaplieu", "create", "2", "TRẦN VĂN BÌNH", {
        "fields": {"TÊN THƯỜNG DÙNG": "TRẦN VĂN BÌNH", "Năm sinh": 2004, "Căn cước": "048204000002"}
    })
    log_action("canbo_nhaplieu", "update", "2", "TRẦN VĂN BÌNH", {
        "changes": [{"field": "Nơi ở hiện nay", "from": "Quảng Nam", "to": "Số 45 đường Lê Lợi, TP Đà Nẵng"}]
    })
    log_action("admin", "delete", "6", "HOÀNG QUỐC VIỆT", {
        "reason": "soft_delete"
    })
    print("-> Successfully seeded audit logs.")


def activate_seeded_workbook():
    print("Activating seeded workbook in application state and SQLite cache...")
    set_workbook(WORKBOOK_PATH, persist=True)
    print("-> Workbook state and cache successfully synchronized!")


def main():
    print("========================================")
    print("   QUẢN LÝ HỒ SƠ THANH NIÊN - SEED DATA")
    print("========================================")
    seed_excel()
    seed_users()
    seed_audit_logs()
    activate_seeded_workbook()
    print("========================================")
    print("Seeding completed successfully!")
    print("Default accounts:")
    print(" - admin / admin123 (Quản trị viên toàn quyền)")
    print(" - canbo_nhaplieu / user1234 (Thêm, xem, sửa)")
    print(" - canbo_kiemtra / user1234 (Chỉ xem)")
    print(" - canbo_toanquyen / user1234 (Thêm, xem, sửa, xóa mềm)")
    print(f"Sample Excel: {WORKBOOK_PATH}")
    print("========================================")


if __name__ == "__main__":
    main()

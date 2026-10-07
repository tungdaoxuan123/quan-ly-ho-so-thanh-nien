"""Static configuration: paths, header sets, and form/preview field layout."""

import sys
from pathlib import Path


def resource_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Dev mode: this file lives at src/quan_ly_ho_so/config.py, so go up to the
    # repo root, where run.py and the .docx templates live. A frozen build
    # instead bundles those next to the executable via PyInstaller --add-data.
    return Path(__file__).resolve().parents[2]


BASE_DIR = resource_dir()
UPLOAD_DIR = BASE_DIR / "uploads"
TEMPLATE = BASE_DIR / "Mau_Ho_So_Thanh_Nien.docx"
LIST_TEMPLATE = BASE_DIR / "Mau_Danh_Sach_Thanh_Nien.xlsx"
DIEN_TEMPLATE_DIR = BASE_DIR / "template"
DEFAULT_OUTPUT_NAME = "Hồ sơ thanh niên đã tạo"
MAX_RECORDS_PER_PAGE = 50
# A per-diện export covers the whole workbook, not one page of it.
MAX_EXPORT_RECORDS = 100000
CACHE_SEARCH_SCOPE = "identity-v2"
SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
# Sentinel for the "Diện" filter, since an empty value already means "no filter".
UNSET_DIEN_FILTER = "__none__"
# The quarters (khu phố) this ward is divided into. Observed 1-55 in the current records; edit this
# single line if the list changes.
QUARTER_CHOICES = [f"Khu phố {number}" for number in range(1, 56)]
FILTER_DB_COLUMNS = {
    "birth_year": ("birth_year", "birth_year_key"),
    "occupation": ("occupation", "occupation_key"),
    "education": ("education", "education_key"),
    "ethnicity": ("ethnicity", "ethnicity_key"),
    "religion": ("religion", "religion_key"),
}

NUMERIC_HEADERS = {
    "Ngày sinh", "Tháng sinh", "Năm sinh", "Năm tốt nghiệp",
    "Ngày sinh cha", "tháng sinh cha", "năm sinh cha",
    "Ngày sinh mẹ", "Tháng sinh mẹ", "Năm sinh mẹ", "năm sinh vợ",
    "Số người con", "cha mẹ có bao nhiêu con", "mấy trai", "mấy gái",
}

LONG_TEXT_HEADERS = {
    "Nơi đăng ký khai sinh (2 cấp)", "Quê quán (2 cấp)", "Thường trú", "Nơi ở hiện nay",
    "Bản thân", "Nơi làm việc, học tập", "Trước 30-4 (cha)\nghi 3 cấp và 2 cấp",
    "sau 30-4 (cha)\nghi 2 cấp", "hiện nay (cha)\nghi 2 cấp",
    "Trước 30-4 (mẹ)\nghi 3 cấp và 2 cấp", "sau 30-4 (mẹ)\nghi 2 cấp",
    "hiện nay (mẹ)\nghi 2 cấp", "Anh chị em 1", "Anh chị em 2", "Anh chị em 3",
    "Anh chị em 4", "Anh chị em 5", "Anh chị em 6", "Anh chị em 7", "Cấp 1", "Cấp 2",
    "Cấp 3", "Hiện nay", "GHI CHÚ\n(Về Sức khỏe; gia cảnh)",
    "TẤT CẢ THÔNG TIN KHÁC\nLIÊN QUAN ĐẾN THANH NIÊN",
}

FORM_GROUPS = [
    ("Thông tin cá nhân", [
        "TÊN THƯỜNG DÙNG", "Tên khai sinh", "Ngày sinh", "Tháng sinh", "Năm sinh", "Căn cước",
        "Nơi đăng ký khai sinh (2 cấp)", "Quê quán (2 cấp)", "Dân tộc", "Tôn giáo", "Thành phần", "Bản thân",
    ]),
    ("Địa chỉ và công việc", [
        "Thường trú", "Nơi ở hiện nay", "Nghề nghiệp", "Văn hóa", "Chuyên môn", "Năm tốt nghiệp",
        "Ngành đào tạo", "Nơi làm việc, học tập",
    ]),
    ("Thông tin cha", [
        "Tên cha", "Sống / Chết", "Ngày sinh cha", "tháng sinh cha", "năm sinh cha", "Nghề nghiệp cha",
        "Trước 30-4 (cha)\nghi 3 cấp và 2 cấp", "sau 30-4 (cha)\nghi 2 cấp", "hiện nay (cha)\nghi 2 cấp",
    ]),
    ("Thông tin mẹ", [
        "Tên mẹ", "Sống / chết (mẹ)", "Ngày sinh mẹ", "Tháng sinh mẹ", "Năm sinh mẹ", "Nghề nghiệp mẹ",
        "Trước 30-4 (mẹ)\nghi 3 cấp và 2 cấp", "sau 30-4 (mẹ)\nghi 2 cấp", "hiện nay (mẹ)\nghi 2 cấp",
    ]),
    ("Vợ, con và anh chị em", [
        "Tên vợ", "năm sinh vợ", "nghề nghiệp vợ", "Số người con", "cha mẹ có bao nhiêu con", "mấy trai",
        "mấy gái", "là Con thứ", "Anh chị em 1", "Anh chị em 2", "Anh chị em 3", "Anh chị em 4",
        "Anh chị em 5", "Anh chị em 6", "Anh chị em 7",
    ]),
    ("Học tập và ghi chú", [
        "Cấp 1", "Cấp 2", "Cấp 3", "Hiện nay", "GHI CHÚ\n(Về Sức khỏe; gia cảnh)",
        "TẤT CẢ THÔNG TIN KHÁC\nLIÊN QUAN ĐẾN THANH NIÊN",
    ]),
]

PREVIEW_FIELD_MAP = [
    ("birth_name", "Tên khai sinh"),
    ("common_name", "TÊN THƯỜNG DÙNG"),
    ("birth_day", "Ngày sinh"),
    ("birth_month", "Tháng sinh"),
    ("birth_year", "Năm sinh"),
    ("citizen_id", "Căn cước"),
    ("birthplace", "Nơi đăng ký khai sinh (2 cấp)"),
    ("hometown", "Quê quán (2 cấp)"),
    ("ethnicity", "Dân tộc"),
    ("religion", "Tôn giáo"),
    ("permanent_address", "Thường trú"),
    ("current_address", "Nơi ở hiện nay"),
    ("family_background", "Thành phần"),
    ("personal_status", "Bản thân"),
    ("education", "Văn hóa"),
    ("specialization", "Chuyên môn"),
    ("training_major", "Ngành đào tạo"),
    ("occupation", "Nghề nghiệp"),
    ("workplace", "Nơi làm việc, học tập"),
    ("father_name", "Tên cha"),
    ("father_life_status", "Sống / Chết"),
    ("father_birth_day", "Ngày sinh cha"),
    ("father_birth_month", "tháng sinh cha"),
    ("father_birth_year", "năm sinh cha"),
    ("father_occupation", "Nghề nghiệp cha"),
    ("father_before_1975", "Trước 30-4 (cha)\nghi 3 cấp và 2 cấp"),
    ("father_after_1975", "sau 30-4 (cha)\nghi 2 cấp"),
    ("father_current", "hiện nay (cha)\nghi 2 cấp"),
    ("mother_name", "Tên mẹ"),
    ("mother_life_status", "Sống / chết (mẹ)"),
    ("mother_birth_day", "Ngày sinh mẹ"),
    ("mother_birth_month", "Tháng sinh mẹ"),
    ("mother_birth_year", "Năm sinh mẹ"),
    ("mother_occupation", "Nghề nghiệp mẹ"),
    ("mother_before_1975", "Trước 30-4 (mẹ)\nghi 3 cấp và 2 cấp"),
    ("mother_after_1975", "sau 30-4 (mẹ)\nghi 2 cấp"),
    ("mother_current", "hiện nay (mẹ)\nghi 2 cấp"),
    ("spouse_name", "Tên vợ"),
    ("spouse_year", "năm sinh vợ"),
    ("spouse_occupation", "nghề nghiệp vợ"),
    ("spouse_children", "Số người con"),
    ("sibling_count", "cha mẹ có bao nhiêu con"),
    ("brother_count", "mấy trai"),
    ("sister_count", "mấy gái"),
    ("birth_order", "là Con thứ"),
    ("education_level_1", "Cấp 1"),
    ("education_level_2", "Cấp 2"),
    ("education_level_3", "Cấp 3"),
    ("current_history", "Hiện nay"),
]

PREVIEW_SIBLING_HEADERS = [f"Anh chị em {number}" for number in range(1, 8)]

"""Enumerated values selectable from dropdown fields in the record form."""

from enum import Enum


class NvqsStatus(Enum):
    """"Diện" NVQS: where a youth currently stands in the military service process."""

    MIEN_DANG_KY = "Miễn Đăng Ký"
    MIEN_GOI_NHAP_NGU = "Miễn Gọi Nhập Ngũ"
    TAM_HOAN_SUC_KHOE = "Tạm Hoạn Sức Khỏe"
    TAM_HOAN_HSSV = "Tạm Hoãn HSSV"
    TAM_HOAN_CHINH_SACH = "Tạm Hoãn Chính Sách"
    TAM_HOAN_DAN_QUAN_THUONG_TRUC = "Tạm Hoãn Dân Quân Thường Trực"
    TAM_HOAN_HOC_VAN_THAP = "Tạm Hoãn Học Vấn Thấp"
    KHONG_TUYEN_CHON = "Không Tuyển Chọn"
    CHUA_XET_TUYEN = "Chưa Xét Tuyển"
    DU_DIEU_KIEN = "Đủ Điều Kiện"
    DAN_QUAN = "Dân Quân"
    DU_HOC = "Du Học"
    GIA_CANH = "Gia Cảnh"
    LAO_DONG_NUOC_NGOAI = "Lao Động Nước Ngoài"
    TON_GIAO = "Tôn Giáo"
    VANG_MAT_DIA_PHUONG = "Vắng Mặt Địa Phương"

    @classmethod
    def options(cls):
        """(stored type code, display text) pairs for building the dropdown."""
        return [(member.name, member.value) for member in cls]

    @classmethod
    def type_for(cls, value):
        """Stable identifier for a display value (e.g. "HEALTH_CHECKED"), for cheap storage/lookup.

        Uses the member name rather than an enum index so a value stays correct even if the
        list above is reordered. Returns None for blank or unrecognized text (free-form legacy data).
        """
        for member in cls:
            if member.value == value:
                return member.name
        return None

    @classmethod
    def value_for(cls, type_code):
        """Display text for a stored type code, or "" when unset or no longer defined."""
        member = cls.__members__.get(type_code or "")
        return member.value if member else ""

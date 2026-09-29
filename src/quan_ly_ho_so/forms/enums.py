"""Enumerated values selectable from dropdown fields in the record form."""

from enum import Enum


class NvqsStatus(Enum):
    """"Diện" NVQS: where a youth currently stands in the military service process."""

    NOT_REGISTERED = "Chưa đăng ký NVQS"
    REGISTERED = "Đã đăng ký NVQS"
    ELIGIBLE_FOR_HEALTH_CHECK = "Đủ điều kiện gọi khám sức khỏe"
    HEALTH_CHECKED = "Đã khám sức khỏe"
    ELIGIBLE_FOR_ENLISTMENT = "Đủ điều kiện nhập ngũ"
    DEFERRED = "Tạm hoãn gọi nhập ngũ"
    EXEMPTED_FROM_ENLISTMENT = "Miễn gọi nhập ngũ"
    ENLISTMENT_DECISION_ISSUED = "Đã có quyết định gọi nhập ngũ"
    ENLISTED = "Đã nhập ngũ"
    DISCHARGED_OR_RESERVE = "Đã xuất ngũ / ngạch dự bị"
    NOT_SUBJECT_TO_REGISTRATION = "Không thuộc diện đăng ký"
    EXEMPTED_FROM_REGISTRATION = "Miễn đăng ký NVQS"

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

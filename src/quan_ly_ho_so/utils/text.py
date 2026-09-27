"""Vietnamese-aware text normalization shared by the cache and form layers."""

import unicodedata

from quan_ly_ho_so.word.export import text


def normalized_header(value):
    return text(value).replace("\r\n", "\n")


def display_value(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d")
    return text(value)


def fold(value):
    """Case- and accent-insensitive text used for Vietnamese search."""
    value = display_value(value).casefold().replace("đ", "d")
    return "".join(char for char in unicodedata.normalize("NFD", value) if unicodedata.category(char) != "Mn")

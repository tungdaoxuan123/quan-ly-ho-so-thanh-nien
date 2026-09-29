"""Turn uploaded paperwork into page images, so related documents can live in the database."""

import io
import logging
import tempfile
from pathlib import Path

from quan_ly_ho_so.errors import WorkbookError

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff"}
PDF_SUFFIXES = {".pdf"}
WORD_SUFFIXES = {".doc", ".docx", ".docm", ".rtf", ".odt"}
SUPPORTED_SUFFIXES = IMAGE_SUFFIXES | PDF_SUFFIXES | WORD_SUFFIXES

MAX_PAGES = 30
FULL_MAX_EDGE = 1700
THUMBNAIL_MAX_EDGE = 360
JPEG_QUALITY = 85
WD_FORMAT_PDF = 17


class AttachmentError(WorkbookError):
    """Raised with a message meant for the person using the app."""


def _require_pillow():
    try:
        from PIL import Image
    except ImportError as error:
        raise AttachmentError("Thiếu thư viện Pillow để xử lý ảnh. Hãy cài lại phụ thuộc của ứng dụng.") from error
    return Image


def _encode(image, max_edge):
    """Scale `image` down to `max_edge` and return it as JPEG bytes."""
    Image = _require_pillow()
    frame = image.convert("RGB") if image.mode not in ("RGB", "L") else image
    frame = frame.copy()
    frame.thumbnail((max_edge, max_edge), Image.LANCZOS)
    buffer = io.BytesIO()
    frame.save(buffer, format="JPEG", quality=JPEG_QUALITY, optimize=True)
    return buffer.getvalue()


def _pdf_pages(pdf_path):
    try:
        import pypdfium2 as pdfium
    except ImportError as error:
        raise AttachmentError("Thiếu thư viện pypdfium2 để đọc tệp PDF. Hãy cài lại phụ thuộc của ứng dụng.") from error

    document = pdfium.PdfDocument(str(pdf_path))
    try:
        images = []
        for index in range(min(len(document), MAX_PAGES)):
            page = document[index]
            width, height = page.get_size()
            longest = max(width, height) or 1
            # get_size() is in points; render at whatever zoom lands near the stored width.
            scale = min(max(FULL_MAX_EDGE / longest, 1.0), 4.0)
            images.append(page.render(scale=scale).to_pil())
        return images
    finally:
        document.close()


def _word_to_pdf(source, target):
    """Drive the copy of Word installed on this machine; it is the only faithful converter here."""
    try:
        import pythoncom
        import win32com.client
    except ImportError as error:
        raise AttachmentError(
            "Không chuyển được tệp Word vì thiếu pywin32. Hãy lưu tệp thành PDF rồi tải lên."
        ) from error

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        document = word.Documents.Open(str(source), ReadOnly=True)
        try:
            document.SaveAs(str(target), FileFormat=WD_FORMAT_PDF)
        finally:
            document.Close(False)
    except AttachmentError:
        raise
    except Exception as error:
        raise AttachmentError(
            f"Không mở được tệp Word bằng Microsoft Word trên máy này ({error}). Hãy lưu tệp thành PDF rồi tải lên."
        ) from error
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception as error:
                logging.warning("Could not close Word: %s", error)
        pythoncom.CoUninitialize()


def document_pages(data, filename):
    """Return [(full JPEG bytes, thumbnail JPEG bytes)] for every page of an uploaded document."""
    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise AttachmentError(f"Không hỗ trợ tệp {suffix or filename}. Hãy tải lên tệp Word, PDF hoặc ảnh.")
    if not data:
        raise AttachmentError("Tệp tải lên rỗng.")

    Image = _require_pillow()
    if suffix in IMAGE_SUFFIXES:
        try:
            images = [Image.open(io.BytesIO(data))]
        except Exception as error:
            raise AttachmentError(f"Không đọc được tệp ảnh: {error}") from error
        return [(_encode(images[0], FULL_MAX_EDGE), _encode(images[0], THUMBNAIL_MAX_EDGE))]

    with tempfile.TemporaryDirectory(prefix="ho-so-lien-quan-") as folder:
        source = Path(folder) / f"upload{suffix}"
        source.write_bytes(data)
        if suffix in WORD_SUFFIXES:
            pdf_path = Path(folder) / "upload.pdf"
            _word_to_pdf(source, pdf_path)
        else:
            pdf_path = source
        images = _pdf_pages(pdf_path)
        if not images:
            raise AttachmentError("Tệp không có trang nào để chuyển thành ảnh.")
        return [(_encode(image, FULL_MAX_EDGE), _encode(image, THUMBNAIL_MAX_EDGE)) for image in images]

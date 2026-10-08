"""Check uploaded paperwork before it is filed: originals kept, Word converted to PDF, thumbnails on demand."""

import io
import logging
import tempfile
from pathlib import Path

from quan_ly_ho_so.errors import WorkbookError

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff"}
PDF_SUFFIXES = {".pdf"}
WORD_SUFFIXES = {".doc", ".docx", ".docm", ".rtf", ".odt"}
SUPPORTED_SUFFIXES = IMAGE_SUFFIXES | PDF_SUFFIXES | WORD_SUFFIXES

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


def _pdf_thumbnail(pdf_bytes):
    """Render page 1 of a PDF as a small JPEG for the document grid; also proves the PDF opens."""
    try:
        import pypdfium2 as pdfium
    except ImportError as error:
        raise AttachmentError("Thiếu thư viện pypdfium2 để đọc tệp PDF. Hãy cài lại phụ thuộc của ứng dụng.") from error

    try:
        document = pdfium.PdfDocument(pdf_bytes)
    except Exception as error:
        raise AttachmentError(f"Không đọc được tệp PDF: {error}") from error
    try:
        if not len(document):
            raise AttachmentError("Tệp PDF không có trang nào.")
        page = document[0]
        width, height = page.get_size()
        # get_size() is in points; render at a zoom that lands near the thumbnail edge.
        scale = THUMBNAIL_MAX_EDGE * 2 / (max(width, height) or 1)
        return _encode(page.render(scale=scale).to_pil(), THUMBNAIL_MAX_EDGE)
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


def prepare_document(data, filename):
    """Check an upload and return (name to store it under, bytes to store).

    Images and PDFs are kept byte for byte; Word-family files are saved as PDF.
    """
    suffix = Path(filename).suffix.lower()
    if suffix not in SUPPORTED_SUFFIXES:
        raise AttachmentError(f"Không hỗ trợ tệp {suffix or filename}. Hãy tải lên tệp Word, PDF hoặc ảnh.")
    if not data:
        raise AttachmentError("Tệp tải lên rỗng.")

    if suffix in IMAGE_SUFFIXES:
        Image = _require_pillow()
        try:
            Image.open(io.BytesIO(data)).load()
        except Exception as error:
            raise AttachmentError(f"Không đọc được tệp ảnh: {error}") from error
        return filename, data

    if suffix in WORD_SUFFIXES:
        with tempfile.TemporaryDirectory(prefix="ho-so-lien-quan-") as folder:
            source = Path(folder) / f"upload{suffix}"
            target = Path(folder) / "upload.pdf"
            source.write_bytes(data)
            _word_to_pdf(source, target)
            data = target.read_bytes()
        filename = f"{Path(filename).stem}.pdf"
    _pdf_thumbnail(data)  # raises AttachmentError if the PDF cannot be opened
    return filename, data


def document_thumbnail(path):
    """Small JPEG of a stored image, or of page 1 of a stored PDF."""
    if path.suffix.lower() in PDF_SUFFIXES:
        return _pdf_thumbnail(path.read_bytes())
    Image = _require_pillow()
    try:
        return _encode(Image.open(path), THUMBNAIL_MAX_EDGE)
    except Exception as error:
        raise AttachmentError(f"Không đọc được tệp ảnh: {error}") from error

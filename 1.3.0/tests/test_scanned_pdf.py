import os

import pytest
from PIL import Image, ImageDraw, ImageFont

import pdf_extractor
from settings_manager import set_setting

LINES = [
    "ATTENTION US",
    "Shipping Date: 08/03/2026",
    "Sales Order #: 123456",
    "Customer and Address: Acme Farms, 1 Main St, Salinas, CA",
    "Lot ID: L-77",
    "Packaging Style: 12x1ct clam",
    "Number of Plants: 240 heads (12 heads x 20 cases)",
]


def _font(size):
    arial = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "arial.ttf")
    if os.path.exists(arial):
        return ImageFont.truetype(arial, size)
    return ImageFont.load_default(size=size)


def _scanned_pdf(path, lines):
    """A letter-size page at 300 dpi containing only an image, like a scan."""
    page = Image.new("RGB", (2550, 3300), "white")
    draw = ImageDraw.Draw(page)
    font = _font(46)
    for index, line in enumerate(lines):
        draw.text((150, 200 + index * 90), line, fill="black", font=font)
    page.save(path, "PDF", resolution=300)
    return str(path)


def test_scanned_pdf_ocr_can_be_disabled(tmp_path):
    set_setting(("ocr", "ocr_scanned_pdfs"), False)
    result = pdf_extractor.process_pdf(_scanned_pdf(tmp_path / "blank.pdf", LINES))
    assert result["error"] == "No extractable text found (PDF may be scanned/image-based)."


@pytest.mark.skipif(
    pdf_extractor.pytesseract is None or not pdf_extractor.TESSERACT_PATH,
    reason="Tesseract OCR is not available",
)
def test_scanned_pdf_is_read_with_ocr(tmp_path):
    result = pdf_extractor.process_pdf(_scanned_pdf(tmp_path / "scan.pdf", LINES))

    assert result["error"] is None
    assert result["ocr_source"] is True
    assert result["template"] == "Template A"
    fields = result["records"][0]
    assert fields["Date"] == "August 3, 2026"
    assert fields["Sales Order Number"] == "123456"
    assert fields["Lot Code"] == "L-77"
    assert fields["Pack Style"] == "12x1ct Clam"
    assert pdf_extractor.SCANNED_PDF_WARNING in result["record_warnings"][0]
    assert result["record_status"][0] != "Verified"


@pytest.mark.skipif(
    pdf_extractor.pytesseract is None or not pdf_extractor.TESSERACT_PATH,
    reason="Tesseract OCR is not available",
)
def test_blank_scanned_pdf_reports_error(tmp_path):
    result = pdf_extractor.process_pdf(_scanned_pdf(tmp_path / "empty.pdf", []))
    assert result["error"] == "No extractable text found, and OCR could not read any text from this PDF."

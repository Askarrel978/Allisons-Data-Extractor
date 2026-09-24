"""
pdf_extractor.py

Reads the three known shipment layouts and returns the six fields needed
for the Excel output:
    Date, Sales Order Number, Customer and Location, Pack Style,
    Number of Plants, Lot Code

Template markers:
    Template A -> ATTENTION US
    Template B -> LOT#
    Template C -> *Lettuce

Templates A/C are text PDFs. Template B is an image with a fixed two-column
shipment table, so its cells are OCR'd individually for better accuracy.
"""

import os
import re
import shutil
from pathlib import Path

import pdfplumber

try:
    from PIL import Image, ImageOps
    import pytesseract
except ImportError:  # handled cleanly at runtime if image input is used
    Image = None
    ImageOps = None
    pytesseract = None


def _configure_tesseract():
    """Find the Windows Tesseract executable automatically when possible."""
    if pytesseract is None:
        return None

    # 1) If Tesseract is already on PATH, use it.
    found = shutil.which("tesseract") or shutil.which("tesseract.exe")
    if found:
        pytesseract.pytesseract.tesseract_cmd = found
        return found

    # 2) Check common Windows install locations.
    candidates = [
        os.environ.get("TESSERACT_CMD"),
        r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
        r"C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Tesseract-OCR", "tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate

    return None


TESSERACT_PATH = _configure_tesseract()

HEADERS = [
    "Date",
    "Sales Order Number",
    "Customer and Location",
    "Pack Style",
    "Number of Plants",
    "Lot Code",
]

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

# ---------------------------------------------------------------------------
# 1. TEMPLATE DETECTION
# ---------------------------------------------------------------------------
TEMPLATE_MARKERS = {
    "Template A": ["ATTENTION US"],
    "Template B": ["LOT#"],
    "Template C": ["*Lettuce"],
}


def detect_template(text: str) -> str:
    lower = text.lower()
    for template_name, markers in TEMPLATE_MARKERS.items():
        for marker in markers:
            if marker.lower() in lower:
                return template_name
    return "Generic"


# ---------------------------------------------------------------------------
# 2. NORMALIZATION HELPERS
# ---------------------------------------------------------------------------

def _clean_label(value: str) -> str:
    value = (value or "").replace("\x00", "")
    value = re.sub(r"\s+", " ", value).strip().lower()
    return value.rstrip(":").strip()


def _clean_value(field: str, value: str) -> str:
    value = (value or "").replace("\x00", "").strip()

    if field == "Customer and Location":
        lines = [re.sub(r"\s+", " ", line).strip() for line in value.splitlines()]
        value = ", ".join(line for line in lines if line)
    elif field == "Pack Style":
        lines = [re.sub(r"\s+", " ", line).strip() for line in value.splitlines()]
        value = " | ".join(line for line in lines if line)
    else:
        value = re.sub(r"\s+", " ", value).strip()

    if field == "Sales Order Number":
        value = re.sub(r"^SO\s*#\s*", "", value, flags=re.IGNORECASE).strip()

    # IMPORTANT: keep the complete Number of Plants cell, including units and
    # anything in parentheses, e.g. "240 heads (12 heads x 20 cases)".
    return value


TABLE_LABELS = {
    "Date": {"shipping date", "date"},
    "Sales Order Number": {
        "sales order #", "sales order", "sales order number", "sales order no.", "sales order no"
    },
    "Customer and Location": {
        "customer and address", "customer and location", "customer/address", "customer"
    },
    "Pack Style": {"packaging style", "pack style", "packing"},
    "Number of Plants": {"number of plants", "plant count", "quantity of plants", "qty of plants"},
    "Lot Code": {"lot id", "lot#", "lot #", "lot code", "lot number"},
}


def extract_fields_from_tables(tables: list) -> dict:
    result = {header: "" for header in HEADERS}
    label_to_field = {}
    for field, labels in TABLE_LABELS.items():
        for label in labels:
            label_to_field[_clean_label(label)] = field

    for table in tables:
        for row in table or []:
            if not row or len(row) < 2:
                continue
            label = _clean_label(row[0])
            if not label:
                continue
            field = label_to_field.get(label)
            if not field:
                continue
            raw_value = "\n".join(
                str(cell).strip() for cell in row[1:] if cell is not None and str(cell).strip()
            )
            if raw_value and not result[field]:
                result[field] = _clean_value(field, raw_value)
    return result


# ---------------------------------------------------------------------------
# 3. PDF REGEX FALLBACK RULES
# ---------------------------------------------------------------------------
GENERIC_RULES = {
    "Date": [
        r"Shipping\s*Date\s*[:\-]\s*([0-9]{1,2}[/\-][0-9]{1,2}[/\-][0-9]{2,4})",
        r"Shipping\s*Date\s*[:\-]\s*([A-Za-z]+\s+[0-9]{1,2},?\s+[0-9]{4})",
        r"Date\s*[:\-]\s*([0-9]{1,2}[/\-][0-9]{1,2}[/\-][0-9]{2,4})",
        r"Date\s*[:\-]\s*([A-Za-z]+\s+[0-9]{1,2},?\s+[0-9]{4})",
    ],
    "Sales Order Number": [
        r"Sales\s*Order\s*(?:Number|No\.?|#)?\s*[:\-]\s*(?:SO\s*#\s*)?([A-Za-z0-9\-]+)",
        r"\bS\.?O\.?\s*(?:Number|No\.?|#)?\s*[:\-]\s*([A-Za-z0-9\-]+)",
    ],
    "Customer and Location": [
        r"Customer\s*and\s*(?:Address|Location)\s*:\s*([\s\S]*?)(?=\s*(?:Lot\s*(?:ID|#|Code|Number)|Grown\s*at|Exported\s*By)\s*:?)",
    ],
    "Pack Style": [
        r"Packaging\s*Style\s*:\s*([\s\S]*?)(?=\s*Number\s*of\s*Plants\s*:)",
        r"Pack\s*Style\s*[:\-]\s*([^\n]+)",
    ],
    "Number of Plants": [
        r"Number\s*of\s*Plants\s*[:\-]\s*([^\n]+)",
        r"(?:Qty|Quantity)\s*(?:of\s*Plants)?\s*[:\-]\s*([^\n]+)",
    ],
    "Lot Code": [
        r"Lot\s*ID\s*[:\-]\s*([A-Za-z0-9\-]+)",
        r"Lot\s*#\s*[:\-]?\s*([A-Za-z0-9\-]+)",
        r"Lot\s*Code\s*[:\-]\s*([A-Za-z0-9\-]+)",
        r"Lot\s*Number\s*[:\-]\s*([A-Za-z0-9\-]+)",
    ],
}

TEMPLATE_OVERRIDES = {
    "Template A": {
        "Lot Code": [r"Lot\s*ID\s*:\s*([A-Za-z0-9\-]+)"],
    },
    "Template B": {
        "Lot Code": [r"LOT\s*#\s*:?\s*([A-Za-z0-9\-]+)"],
    },
    "Template C": {
        "Lot Code": [r"Lot\s*(?:ID|#|Code|Number)\s*:?\s*([A-Za-z0-9\-]+)"],
    },
}


def _rules_for_template(template_name: str) -> dict:
    rules = dict(GENERIC_RULES)
    rules.update(TEMPLATE_OVERRIDES.get(template_name, {}))
    return rules


def extract_fields_regex(text: str, template_name: str) -> dict:
    rules = _rules_for_template(template_name)
    result = {header: "" for header in HEADERS}
    for field in HEADERS:
        for pattern in rules.get(field, []):
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                value = _clean_value(field, match.group(1))
                if value:
                    result[field] = value
                    break
    return result


def extract_fields(text: str, template_name: str, tables: list | None = None) -> dict:
    result = {header: "" for header in HEADERS}
    if tables:
        result.update(extract_fields_from_tables(tables))
    regex_fields = extract_fields_regex(text, template_name)
    for field in HEADERS:
        if not result[field] and regex_fields[field]:
            result[field] = regex_fields[field]
    return result


# ---------------------------------------------------------------------------
# 4. TEMPLATE B IMAGE EXTRACTION
# ---------------------------------------------------------------------------

def _ocr_cell(image, psm=6) -> str:
    """OCR one table cell. Upscaling makes the small screenshot text clearer."""
    if pytesseract is None or ImageOps is None:
        raise RuntimeError(
            "Template B image support requires the Python packages Pillow and pytesseract. "
            "Install them with: pip install pillow pytesseract"
        )
    if not TESSERACT_PATH:
        raise RuntimeError(
            "Tesseract OCR itself was not found. Install Tesseract OCR for Windows, "
            "or set the TESSERACT_CMD environment variable to the full path to tesseract.exe. "
            r"Common location: C:\Program Files\Tesseract-OCR\tesseract.exe"
        )

    gray = image.convert("L")
    gray = ImageOps.autocontrast(gray)
    scale = 5
    gray = gray.resize((max(1, gray.width * scale), max(1, gray.height * scale)))
    return pytesseract.image_to_string(gray, config=f"--psm {psm}").strip()


def _group_nearby(values: list[int]) -> list[int]:
    """Collapse adjacent detected grid-line pixels to one coordinate."""
    if not values:
        return []
    groups = [[values[0]]]
    for value in values[1:]:
        if value <= groups[-1][-1] + 1:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [round(sum(group) / len(group)) for group in groups]


def _detect_grid_lines(image) -> tuple[list[int], list[int]]:
    """Detect the bright horizontal/vertical rules in Template B's table."""
    gray = image.convert("L")
    width, height = gray.size
    pix = gray.load()

    # The screenshot uses light grid lines on a dark background.
    horizontal = []
    for y in range(height):
        bright = sum(1 for x in range(width) if pix[x, y] >= 180)
        if bright >= int(width * 0.75):
            horizontal.append(y)

    vertical = []
    for x in range(width):
        bright = sum(1 for y in range(height) if pix[x, y] >= 180)
        if bright >= int(height * 0.70):
            vertical.append(x)

    return _group_nearby(horizontal), _group_nearby(vertical)


def _template_b_field_from_label(label: str) -> str | None:
    normalized = _clean_label(label)
    compact = re.sub(r"[^a-z0-9#]+", " ", normalized).strip()

    if "shipping date" in compact:
        return "Date"
    if "sales order" in compact:
        return "Sales Order Number"
    if "customer" in compact and ("address" in compact or "location" in compact):
        return "Customer and Location"
    if "packaging style" in compact or "pack style" in compact:
        return "Pack Style"
    if "number of plants" in compact:
        return "Number of Plants"
    if "lot#" in normalized or compact == "lot#" or compact.startswith("lot #"):
        return "Lot Code"
    return None


def extract_template_b_image(path: str) -> tuple[dict, str, list[str]]:
    """
    Extract Template B from the image table itself.

    Instead of OCR'ing the entire screenshot as one block, detect the grid and OCR
    each right-hand value cell separately. This preserves values such as:
        240 heads (12 heads x 20 cases)
    and is much more reliable for the dark-background Template B screenshots.
    """
    if Image is None or pytesseract is None:
        raise RuntimeError(
            "Template B image support requires Pillow and pytesseract. "
            "Install with: pip install pillow pytesseract"
        )
    if not TESSERACT_PATH:
        raise RuntimeError(
            "Tesseract OCR itself was not found. Install Tesseract OCR for Windows, "
            "or set the TESSERACT_CMD environment variable to the full path to tesseract.exe. "
            r"Common location: C:\Program Files\Tesseract-OCR\tesseract.exe"
        )

    image = Image.open(path).convert("RGB")
    horizontal, vertical = _detect_grid_lines(image)

    if len(horizontal) < 8 or len(vertical) < 2:
        raise RuntimeError("Template B table grid could not be detected in this image.")

    # Use the interior divider between the label and value columns.
    width, height = image.size
    divider_candidates = [x for x in vertical if width * 0.20 < x < width * 0.75]
    divider = divider_candidates[0] if divider_candidates else vertical[0]
    right_edge = max(x for x in vertical if x > divider) if any(x > divider for x in vertical) else width - 1

    result = {header: "" for header in HEADERS}
    labels_seen = []
    raw_lines = []

    for top, bottom in zip(horizontal, horizontal[1:]):
        if bottom - top < 8:
            continue

        left_cell = image.crop((1, top + 1, max(2, divider - 1), bottom - 1))
        right_cell = image.crop((divider + 1, top + 1, max(divider + 2, right_edge - 1), bottom - 1))

        label = _ocr_cell(left_cell, psm=7)
        value = _ocr_cell(right_cell, psm=6)
        if label:
            labels_seen.append(label)
            raw_lines.append(label)
        if value:
            raw_lines.append(value)

        field = _template_b_field_from_label(label)
        if field and value:
            result[field] = _clean_value(field, value)

    raw_text = "\n".join(raw_lines)
    marker_found = any("lot#" in _clean_label(label) for label in labels_seen)

    warnings = []
    if not marker_found:
        warnings.append('Template B grid matched, but the "LOT#" marker was not read clearly by OCR.')

    # Small targeted correction for a common OCR error in this exact packaging format.
    # e.g. "12xict clam" -> "12x1ct clam". This does not alter other pack styles.
    pack = result.get("Pack Style", "")
    if pack:
        pack = re.sub(r"(?i)(\d+)x[iIlL]ct\b", r"\1x1ct", pack)
        result["Pack Style"] = pack

    return result, raw_text, warnings


# ---------------------------------------------------------------------------
# 5. PDF / TOP-LEVEL ENTRY POINT
# ---------------------------------------------------------------------------

def read_pdf(path: str) -> tuple[str, list]:
    text_parts = []
    tables = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text_parts.append(page.extract_text() or "")
            tables.extend(page.extract_tables() or [])
    return "\n".join(text_parts), tables


def read_pdf_text(path: str) -> str:
    text, _ = read_pdf(path)
    return text


def process_pdf(path: str) -> dict:
    """Backward-compatible entry point. Handles both PDFs and Template B images."""
    out = {
        "source_file": path,
        "template": "",
        "fields": {h: "" for h in HEADERS},
        "warnings": [],
        "raw_text": "",
        "error": None,
    }

    try:
        extension = Path(path).suffix.lower()

        if extension in IMAGE_EXTENSIONS:
            fields, raw_text, image_warnings = extract_template_b_image(path)
            out["raw_text"] = raw_text
            out["template"] = "Template B"
            out["fields"] = fields
            out["warnings"].extend(image_warnings)
        else:
            text, tables = read_pdf(path)
            out["raw_text"] = text
            if not text.strip():
                out["error"] = "No extractable text found (PDF may be scanned/image-based)."
                return out

            template_name = detect_template(text)
            out["template"] = template_name
            out["fields"] = extract_fields(text, template_name, tables)

        missing = [field for field, value in out["fields"].items() if not value]
        if missing:
            out["warnings"].append("Could not find: " + ", ".join(missing))

        if out["template"] == "Template C" and not out["fields"].get("Lot Code"):
            out["warnings"].append(
                "Template C sample does not contain a lot code, so Lot Code was left blank."
            )

        if out["template"] == "Generic":
            out["warnings"].append(
                "Template not recognized (using generic rules) - please review."
            )

    except Exception as e:
        out["error"] = f"Failed to process file: {e}"

    return out

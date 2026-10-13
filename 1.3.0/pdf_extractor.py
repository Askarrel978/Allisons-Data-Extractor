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
Scanned PDFs without selectable text are OCR'd page by page and then read
with the same text rules as other PDFs.
"""

import math
import os
import re
import shutil
import sys
import logging
from collections import Counter
from pathlib import Path
from datetime import datetime

import pdfplumber

from template_config import get_template_markers, get_template_settings
from settings_manager import load_settings

logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdffont").setLevel(logging.ERROR)

try:
    from PIL import Image, ImageOps, ImageSequence, ImageStat
    import pytesseract
except ImportError:  # handled cleanly at runtime if image input is used
    Image = None
    ImageOps = None
    ImageSequence = None
    ImageStat = None
    pytesseract = None


def _configure_tesseract():
    """Find the Windows Tesseract executable automatically when possible."""
    if pytesseract is None:
        return None

    # Prefer a portable Tesseract-OCR folder beside the program. This works
    # both while developing and after PyInstaller compiles the application.
    if getattr(sys, "frozen", False):
        app_dir = os.path.dirname(sys.executable)
    else:
        app_dir = os.path.dirname(os.path.abspath(__file__))

    settings = load_settings()
    override = str(settings.get("ocr", {}).get("tesseract_cmd_override", "") or "").strip()
    meipass = getattr(sys, "_MEIPASS", "")

    candidates = [
        override,
        os.environ.get("TESSERACT_CMD"),
        os.path.join(app_dir, "Tesseract-OCR", "tesseract.exe"),
        os.path.join(app_dir, "_internal", "Tesseract-OCR", "tesseract.exe"),
        os.path.join(meipass, "Tesseract-OCR", "tesseract.exe") if meipass else None,
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR", "tesseract.exe"),
        os.path.join(os.environ.get("LOCALAPPDATA", ""), "Tesseract-OCR", "tesseract.exe"),
        shutil.which("tesseract"),
        shutil.which("tesseract.exe"),
    ]
    for candidate in candidates:
        if candidate and os.path.isfile(candidate):
            pytesseract.pytesseract.tesseract_cmd = candidate
            return candidate

    return None


TESSERACT_PATH = _configure_tesseract()


def refresh_tesseract_configuration():
    """Re-read Tesseract settings and return the active executable path."""
    global TESSERACT_PATH
    TESSERACT_PATH = _configure_tesseract()
    return TESSERACT_PATH

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
def detect_template(text: str) -> str:
    lower = text.lower()
    for template_name, markers in get_template_markers().items():
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


def _configured_table_labels() -> dict:
    configured = load_settings().get("extraction", {}).get("table_labels", {})
    if not isinstance(configured, dict) or not configured:
        return TABLE_LABELS
    result = {}
    for field in HEADERS:
        values = configured.get(field, TABLE_LABELS.get(field, set()))
        result[field] = {str(value) for value in values if str(value).strip()}
    return result


def extract_fields_from_tables(tables: list) -> dict:
    result = {header: "" for header in HEADERS}
    label_to_field = {}
    for field, labels in _configured_table_labels().items():
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
    extraction = load_settings().get("extraction", {})
    configured_generic = extraction.get("generic_rules", {})
    configured_overrides = extraction.get("template_overrides", {})

    rules = {}
    for field in HEADERS:
        values = configured_generic.get(field, GENERIC_RULES.get(field, [])) if isinstance(configured_generic, dict) else GENERIC_RULES.get(field, [])
        rules[field] = [str(value) for value in values]

    override_source = configured_overrides.get(template_name, {}) if isinstance(configured_overrides, dict) else {}
    if not override_source:
        override_source = TEMPLATE_OVERRIDES.get(template_name, {})
    for field, values in override_source.items():
        rules[field] = [str(value) for value in values]
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
# 4. TEMPLATE B EXTRACTION (IMAGES, SCANNED PAGES AND PDF TABLES)
# ---------------------------------------------------------------------------
#
# A Template B source can hold any number of shipment tables: stacked,
# side by side, different widths, light or dark theme, surrounded by other
# email content. Images are handled in three steps:
#   1. find every table independently (_find_tables): a table is a run of
#      horizontal rules that are joined to each other by vertical rules;
#   2. read each table row by row (label cell + value cell);
#   3. assemble shipments (_TemplateBRecordBuilder): every table, every
#      "Plant Species" row and every repeated field starts a new shipment,
#      and an unreadable label is recovered from Template B's fixed row order
#      when the value fits that field.

# Template B's fixed row order. Entries that are not in HEADERS are rows that
# are not exported but still anchor the position of the rows around them.
TEMPLATE_B_ROW_ORDER = (
    "plant species",
    "country of origin",
    "Pack Style",
    "Number of Plants",
    "Date",
    "Lot Code",
    "Sales Order Number",
    "Customer and Location",
    "grown at",
    "exported by",
)

_OCR_MISSING_MESSAGE = (
    "Tesseract OCR itself was not found. Install Tesseract OCR for Windows, "
    "or set the TESSERACT_CMD environment variable to the full path to tesseract.exe. "
    r"Common location: C:\Program Files\Tesseract-OCR\tesseract.exe"
)


def _require_ocr():
    if Image is None or ImageOps is None or pytesseract is None:
        raise RuntimeError(
            "Template B image support requires the Python packages Pillow and pytesseract. "
            "Install them with: pip install pillow pytesseract"
        )
    if not TESSERACT_PATH:
        raise RuntimeError(_OCR_MISSING_MESSAGE)


def _strip_rule_marks(text: str) -> str:
    """Remove table-border fragments OCR reads as | [ ] at the ends of lines."""
    lines = [line.strip().strip("|[]{}¦").strip() for line in str(text or "").splitlines()]
    return "\n".join(line for line in lines if line)


def _ocr_cell(image, psm=6, scale=5) -> str:
    """OCR one table cell. Upscaling makes the small screenshot text clearer."""
    _require_ocr()
    gray = ImageOps.autocontrast(image.convert("L"))
    gray = gray.resize((max(1, round(gray.width * scale)), max(1, round(gray.height * scale))))
    return _strip_rule_marks(pytesseract.image_to_string(gray, config=f"--psm {psm}"))


def _ocr_cell_for_vote(image, psm: int, scale: float) -> str:
    """One voting read: dark text on white with a margin, as Tesseract prefers."""
    gray = ImageOps.autocontrast(image.convert("L"))
    if ImageStat.Stat(gray).mean[0] < 128:
        gray = ImageOps.invert(gray)
    big = gray.resize(
        (max(1, round(gray.width * scale)), max(1, round(gray.height * scale))),
        Image.Resampling.LANCZOS,
    )
    big = ImageOps.expand(big, border=max(4, int(big.height * 0.4)), fill=255)
    return _strip_rule_marks(pytesseract.image_to_string(big, config=f"--psm {psm}"))


def _ocr_cell_alternatives(image, scale=5):
    """Differently prepared OCR reads of a cell, used when the normal read fails validation."""
    gray = ImageOps.autocontrast(image.convert("L"))
    # Tesseract prefers dark text on a light background.
    if ImageStat.Stat(gray).mean[0] < 128:
        gray = ImageOps.invert(gray)
    for factor, binarize in ((scale - 1, False), (scale, True), (scale + 1, False)):
        factor = max(2, factor)
        big = gray.resize((max(1, gray.width * factor), max(1, gray.height * factor)), Image.Resampling.LANCZOS)
        if binarize:
            big = big.point(lambda p: 255 if p > 128 else 0)
        yield _strip_rule_marks(pytesseract.image_to_string(big, config="--psm 6"))


def _ocr_digits_cell(image, scale=6) -> str:
    """OCR a short numeric cell such as a sales order using a digits-only pass."""
    if pytesseract is None or ImageOps is None:
        return ""
    gray = ImageOps.autocontrast(image.convert("L"))
    gray = gray.resize((max(1, gray.width * scale), max(1, gray.height * scale)))
    text = pytesseract.image_to_string(
        gray,
        config="--psm 7 -c tessedit_char_whitelist=0123456789",
    ).strip()
    match = re.search(r"\d{5,9}", text)
    return match.group(0) if match else ""


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
    """Whole-image rule detection used by the original single-layout reader.

    Kept as a fallback for images where _find_tables finds nothing.
    """
    gray = image.convert("L")
    width, height = gray.size

    # The screenshot uses light grid lines on a dark background. Thresholds
    # are administrator-configurable through Advanced Configuration JSON.
    grid_settings = load_settings().get("extraction", {}).get("template_b_grid", {})
    bright_threshold = int(grid_settings.get("bright_threshold", 100))
    horizontal_coverage = float(grid_settings.get("horizontal_coverage", 0.75))
    vertical_coverage = float(grid_settings.get("vertical_coverage", 0.70))

    # Threshold once, then count bright pixels per row/column with bytes.count,
    # which runs in C instead of visiting every pixel from Python.
    binary = gray.point(lambda p: 255 if p >= bright_threshold else 0)
    rows = binary.tobytes()
    columns = binary.transpose(Image.Transpose.TRANSPOSE).tobytes()

    min_row = int(width * horizontal_coverage)
    horizontal = [
        y for y in range(height)
        if rows.count(255, y * width, (y + 1) * width) >= min_row
    ]

    min_column = int(height * vertical_coverage)
    vertical = [
        x for x in range(width)
        if columns.count(255, x * height, (x + 1) * height) >= min_column
    ]

    return _group_nearby(horizontal), _group_nearby(vertical)


# --- Table finding ---------------------------------------------------------

# A run of rule pixels, tolerating 1-2 pixel breaks from anti-aliasing/JPEG.
_RULE_RUN = re.compile(rb"\xff+(?:\x00{1,2}\xff+)*")


def _rule_mask(image):
    """Return (row-major bytes, column-major bytes) where table rules are 255.

    Light-theme images are inverted so that rules are always bright, and the
    threshold follows the background so faint grey borders still register.
    """
    gray = image.convert("L")
    grid_settings = load_settings().get("extraction", {}).get("template_b_grid", {})
    threshold = int(grid_settings.get("bright_threshold", 100))

    histogram = gray.histogram()
    half = gray.width * gray.height / 2
    running = 0
    background = 0
    for level, count in enumerate(histogram):
        running += count
        if running >= half:
            background = level
            break
    if background > 127:
        gray = ImageOps.invert(gray)
        threshold = min(threshold, (255 - background) + 25)

    binary = gray.point(lambda p: 255 if p >= threshold else 0)
    return binary.tobytes(), binary.transpose(Image.Transpose.TRANSPOSE).tobytes()


def _horizontal_rules(rows: bytes, width: int, height: int) -> list[tuple[int, int, int, int]]:
    """Find horizontal rules as (y_top, y_bottom, x_left, x_right)."""
    min_length = max(40, int(width * 0.08))
    found = []
    for y in range(height):
        start = y * width
        if rows.count(255, start, start + width) < min_length:
            continue
        best = max(_RULE_RUN.finditer(rows, start, start + width), key=lambda m: m.end() - m.start())
        if best.end() - best.start() >= min_length:
            found.append((y, best.start() - start, best.end() - start - 1))

    rules = []
    for y, x0, x1 in found:
        if rules and y == rules[-1][1] + 1 and _same_span(rules[-1], (y, y, x0, x1)):
            top, _bottom, left, right = rules[-1]
            rules[-1] = (top, y, min(left, x0), max(right, x1))
        else:
            rules.append((y, y, x0, x1))
    return rules


def _same_span(a, b) -> bool:
    tolerance = max(6, int(0.04 * max(a[3] - a[2], b[3] - b[2])))
    return abs(a[2] - b[2]) <= tolerance and abs(a[3] - b[3]) <= tolerance


def _rules_connected(columns: bytes, height: int, upper, lower) -> bool:
    """True when a vertical rule joins two horizontal rules (same table)."""
    gap_top = upper[1] + 1
    gap_bottom = lower[0] - 1
    span = gap_bottom - gap_top + 1
    if span <= 0:
        return True
    needed = span - max(1, span // 10)
    for x in range(max(upper[2], lower[2]), min(upper[3], lower[3]) + 1):
        base = x * height
        if columns.count(255, base + gap_top, base + gap_bottom + 1) >= needed:
            return True
    return False


def _vertical_runs(columns: bytes, height: int, x_from: int, x_to: int, top: int, bottom: int, coverage: float):
    """Vertical rules between ``top`` and ``bottom`` as [(x_left, x_right)]."""
    needed = coverage * (bottom - top + 1)
    xs = [
        x for x in range(x_from, x_to + 1)
        if columns.count(255, x * height + top, x * height + bottom + 1) >= needed
    ]
    runs = []
    for x in xs:
        if runs and x == runs[-1][1] + 1:
            runs[-1] = (runs[-1][0], x)
        else:
            runs.append((x, x))
    return runs


def _table_layout(rules, columns: bytes, width: int, height: int) -> dict | None:
    """Cell geometry for one table: label/value column bounds and row bands."""
    top = rules[0][1] + 1
    bottom = rules[-1][0] - 1
    if bottom - top < 10:
        return None
    lefts = sorted(rule[2] for rule in rules)
    rights = sorted(rule[3] for rule in rules)
    left = lefts[len(lefts) // 2]
    right = rights[len(rights) // 2]
    table_width = right - left + 1
    x_from, x_to = max(0, left - 3), min(width - 1, right + 3)

    divider = None
    verticals = []
    for coverage in (0.6, 0.35):
        verticals = _vertical_runs(columns, height, x_from, x_to, top, bottom, coverage)
        inner = [v for v in verticals if v[0] > left + 0.12 * table_width and v[1] < right - 0.12 * table_width]
        if inner:
            divider = inner[0]
            break
    estimated = divider is None
    if estimated:
        # Template B's label column is about 41% of the table width.
        middle = left + round(0.41 * table_width)
        divider = (middle, middle)

    edge = 0.05 * table_width + 3
    left_borders = [v for v in verticals if v[1] <= left + edge]
    right_borders = [v for v in verticals if v[0] >= right - edge]
    inner_left = left_borders[-1][1] + 1 if left_borders else left
    inner_right = right_borders[0][0] - 1 if right_borders else right

    rows = [
        (upper[1] + 1, lower[0])
        for upper, lower in zip(rules, rules[1:])
        if lower[0] - upper[1] - 1 >= 6
    ]
    if not rows or divider[0] - inner_left < 8 or inner_right - divider[1] < 8:
        return None
    # Keep a 2px margin from the rules: blurred (scaled/JPEG) rule edges that
    # fall under the threshold would otherwise be read as "|" characters.
    return {
        "label_x": (inner_left + 2, divider[0] - 1),
        "value_x": (divider[1] + 3, inner_right - 1),
        "rows": rows,
        "divider_estimated": estimated,
        "position": (top, left),
    }


def _find_tables(image) -> list[dict]:
    """Locate every ruled table in the image, in reading order."""
    width, height = image.size
    rows, columns = _rule_mask(image)
    groups = []
    for rule in _horizontal_rules(rows, width, height):
        best = None
        for group in groups:
            last = group[-1]
            if last[1] >= rule[0] or not _same_span(last, rule):
                continue
            if best is not None and best[-1][1] >= last[1]:
                continue
            if _rules_connected(columns, height, last, rule):
                best = group
        if best is not None:
            best.append(rule)
        else:
            groups.append([rule])

    layouts = []
    for group in groups:
        if len(group) < 3:
            continue
        layout = _table_layout(group, columns, width, height)
        if layout:
            layouts.append(layout)
    # Top to bottom; tables side by side (similar top) left to right.
    layouts.sort(key=lambda item: (item["position"][0] // 20, item["position"][1]))
    return layouts


def _legacy_layout(image) -> dict | None:
    """The original whole-image grid reader, used when no table is found."""
    horizontal, vertical = _detect_grid_lines(image)
    if len(horizontal) < 8 or len(vertical) < 2:
        return None
    width = image.width
    divider_candidates = [x for x in vertical if width * 0.20 < x < width * 0.75]
    divider = divider_candidates[0] if divider_candidates else vertical[0]
    right_edge = max((x for x in vertical if x > divider), default=width - 1)
    return {
        "label_x": (1, max(2, divider - 1)),
        "value_x": (divider + 1, max(divider + 2, right_edge - 1)),
        "rows": [(top + 1, bottom - 1) for top, bottom in zip(horizontal, horizontal[1:]) if bottom - top >= 8],
        "divider_estimated": False,
        "position": (0, 0),
    }


def _line_height(layout: dict) -> int:
    """Height in pixels of the table's shortest (one-line) row."""
    heights = [bottom - top for top, bottom in layout["rows"] if bottom - top >= 8]
    return min(heights) if heights else 20


def _ocr_scale(layout: dict) -> int:
    """Upscale factor for OCR: 5x for normal screenshots, less for high-resolution ones."""
    line_height = _line_height(layout)
    if line_height <= 24:
        return 5
    return max(2, min(5, math.ceil(115 / line_height)))


# --- Field reading -----------------------------------------------------------

def _fix_pack_ocr(text: str) -> str:
    """Correct recurring OCR misreads of compact pack styles such as 12x1ct clam."""
    if not text:
        return text
    # "12xict" / "12xlct" -> "12x1ct"
    text = re.sub(r"(?i)\b(\d+)\s*[x×]\s*[il|!]\s*ct\b", r"\1x1ct", text)
    # "12x1ict" / "8x2lct" -> "12x1ct" / "8x2ct" (a stray stroke before "ct")
    text = re.sub(r"(?i)\b(\d+)\s*[x×]\s*(\d+)\s*[il|!]\s*ct\b", r"\1x\2ct", text)
    # Small "clam" is routinely read as "dam" / "cam" / "ciam".
    text = re.sub(r"(?i)\b(\d+x\d+ct)\s*(?:dam|cam|ciam|c1am|clarn|olam|elam|diam)\b", r"\1 clam", text)
    return text


def _plant_count_parts(value: str):
    """Split "1520 heads (12 heads x 60 cases / 16 heads x 50 cases)" into (1520, [(12, 60), (16, 50)])."""
    total = re.match(r"\s*([\d,]+)\s*(?:heads?|plants?)?", value or "", re.IGNORECASE)
    breakdown = re.search(r"\((.*)\)", value or "")
    if not total or not total.group(1).strip(",") or not breakdown:
        return None
    groups = re.findall(
        r"(\d+)\s*(?:heads?|plants?)?\s*[x×*]\s*(\d+)\s*cases?",
        breakdown.group(1),
        re.IGNORECASE,
    )
    if not groups:
        return None
    return int(total.group(1).replace(",", "")), [(int(each), int(cases)) for each, cases in groups]


def _plant_count_ok(value: str) -> bool:
    parts = _plant_count_parts(value)
    return parts is not None and sum(each * cases for each, cases in parts[1]) == parts[0]


def _value_is_plausible(field: str, value: str) -> bool:
    """Does an extracted value have the shape its field should have?"""
    value = str(value or "").strip()
    if not value:
        return False
    if field == "Date":
        return normalize_date(value)[1] is None
    if field == "Sales Order Number":
        pattern = get_template_settings().get("Template B", {}).get("sales_order_pattern") or r"^\d{5,9}$"
        try:
            return re.fullmatch(pattern, value) is not None
        except re.error:
            return re.fullmatch(r"\d{5,9}", value) is not None
    if field == "Lot Code":
        return re.fullmatch(r"[A-Za-z0-9\-]{3,20}", value) is not None and any(ch.isdigit() for ch in value)
    if field == "Pack Style":
        return normalize_pack_style(value)[1] is None
    if field == "Number of Plants":
        return _plant_count_ok(value)
    if field == "Customer and Location":
        return value.count(",") >= 2
    return True


def _clean_template_b_value(field: str, text: str) -> str:
    # OCR reads the same apostrophe as ’ or ' depending on size; use one form
    # so the same customer is always written the same way.
    value = _clean_value(field, str(text or "").replace("’", "'").replace("‘", "'"))
    if field == "Pack Style":
        value = _fix_pack_ocr(value)
    return value


# Identity fields have no internal check (unlike Pack Style or the plant-count
# arithmetic), so a one-character OCR slip would go unnoticed. They are read
# with the text magnified to several heights (pixels per one-line row) and the
# majority wins; any disagreement is flagged for review. Different heights
# rarely make the same mistake, which is what makes the vote reliable. (On 200
# random codes in five fonts, dark and light: 98% correct, and all but one
# visually ambiguous J/I error were caught as disagreements.)
_VOTE_LINE_HEIGHTS = {
    "Sales Order Number": (60, 40, 85, 130, 50),
    "Lot Code": (60, 40, 85, 130, 50),
    "Date": (60, 40, 85),
    "Customer and Location": (60, 40, 85),
}
_SINGLE_LINE_FIELDS = ("Sales Order Number", "Lot Code", "Date")


def _clean_read(field: str, text: str) -> str:
    value = _clean_template_b_value(field, text)
    if field == "Sales Order Number":
        digits = re.sub(r"\D", "", value)
        if 5 <= len(digits) <= 9:
            value = digits
    return value


def _vote_key(value: str) -> str:
    return re.sub(r"[\s’'`]", "", value).casefold()


def _read_value_cell(field: str, cell, scale: int, line_height: int) -> tuple[str, str | None]:
    """OCR a value cell; return (value, review note or None).

    ``line_height`` is the height in pixels of a one-line row of this table.
    """
    psm = 7 if field in _SINGLE_LINE_FIELDS else 6
    if field in _VOTE_LINE_HEIGHTS:
        factors = [target / max(line_height, 1) for target in _VOTE_LINE_HEIGHTS[field]]
        reads = [_clean_read(field, _ocr_cell_for_vote(cell, psm, factor)) for factor in factors]
    else:
        factors = [scale]
        reads = [_clean_read(field, _ocr_cell(cell, psm=psm, scale=scale))]
    plausible = [value for value in reads if _value_is_plausible(field, value)]

    if not plausible and field == "Sales Order Number":
        # Small white digits on dark screenshots sometimes need a digits-only pass.
        plausible = [
            digits for digits in (_ocr_digits_cell(cell, scale=factor) for factor in factors[:3])
            if _value_is_plausible(field, digits)
        ]
    if not plausible:
        for alternative in _ocr_cell_alternatives(cell, scale):
            candidate = _clean_read(field, alternative)
            if _value_is_plausible(field, candidate):
                return candidate, None
        return reads[0], None

    votes = Counter(_vote_key(value) for value in plausible)
    # Majority first; on a tie, the read at the normal magnification wins.
    best = max(plausible, key=lambda value: (votes[_vote_key(value)], -plausible.index(value)))
    note = None
    if len(votes) > 1:
        others = sorted({value for value in plausible if _vote_key(value) != _vote_key(best)})
        note = (
            f'OCR readings of {field} disagreed ("{best}" vs "{others[0]}"); '
            f'"{best}" was used. Please verify against the source.'
        )
    return best, note


def _template_b_field_from_label(label: str) -> str | None:
    normalized = _clean_label(label)
    compact = re.sub(r"[^a-z0-9#]+", " ", normalized).strip()

    configured = load_settings().get("extraction", {}).get("template_b_label_terms", {})
    if isinstance(configured, dict):
        for field, terms in configured.items():
            if field not in HEADERS:
                continue
            for term in terms or []:
                term_norm = _clean_label(str(term))
                term_compact = re.sub(r"[^a-z0-9#]+", " ", term_norm).strip()
                if term_norm and (term_norm in normalized or (term_compact and term_compact in compact)):
                    return field

    # Fallbacks keep older configurations resilient to minor OCR damage.
    if "shipping date" in compact:
        return "Date"
    if "sales order" in compact or ("sales" in compact and "order" in compact):
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


def _is_template_b_record_start(label: str) -> bool:
    """Return True when OCR has found the first row of a Template B shipment."""
    normalized = _clean_label(label)
    compact = re.sub(r"[^a-z0-9]+", " ", normalized).strip()

    settings = get_template_settings().get("Template B", {})
    configured_terms = settings.get("record_start_terms", ["Plant Species"]) or ["Plant Species"]
    for term in configured_terms:
        term_compact = re.sub(r"[^a-z0-9]+", " ", _clean_label(str(term))).strip()
        if term_compact and term_compact in compact:
            return True

    # Common OCR fallback for Plant Species even when one or two letters are damaged.
    words = compact.split()
    return len(words) >= 2 and words[0].startswith("plant") and words[1].startswith("spec")


def _template_b_row_key(label: str) -> str | None:
    """Which Template B row a label belongs to (a HEADERS field or a TEMPLATE_B_ROW_ORDER key)."""
    if not str(label or "").strip():
        return None
    if _is_template_b_record_start(label):
        return "plant species"
    field = _template_b_field_from_label(label)
    if field:
        return field
    words = re.sub(r"[^a-z]+", " ", _clean_label(label))
    if "country" in words or "origin" in words:
        return "country of origin"
    if "grown" in words:
        return "grown at"
    if "export" in words:
        return "exported by"
    return None


def _finish_template_b_record(record: dict, labels_seen: list[str], warnings: list[str]) -> tuple[dict, list[str]]:
    """Clean one Template B shipment and attach record-specific warnings."""
    record = dict(record)
    if record.get("Pack Style"):
        record["Pack Style"] = _fix_pack_ocr(record["Pack Style"])

    record_warnings = list(warnings)
    marker_found = any("lot#" in _clean_label(label) or "lot #" in _clean_label(label) for label in labels_seen)
    if not marker_found:
        record_warnings.append('Template B record found, but the "LOT#" marker was not read clearly by OCR.')

    missing = [field for field, value in record.items() if field in HEADERS and not value]
    if missing:
        record_warnings.append("Could not find: " + ", ".join(missing))

    return record, record_warnings


class _TemplateBRecordBuilder:
    """Turns a stream of (label, value) rows into shipments.

    A new shipment starts at a "Plant Species" row, at a table boundary (when
    requested) and when a field that is already filled appears again. A row
    whose label could not be read is assigned to the next field in Template
    B's row order, but only when its value has that field's shape.
    """

    def __init__(self):
        self.records = []
        self.warnings = []
        self._current = None
        self._labels = []
        self._notes = []
        self._started_by_label = False
        self._position = None
        self._table_notes = []

    def new_table(self, separate_records=True, notes=None):
        if separate_records:
            self._flush()
        self._position = None
        self._table_notes = list(notes or [])

    def add_row(self, label: str, key: str | None, read_value) -> None:
        """``read_value(field)`` is only called for exported fields, so OCR is skipped for other rows."""
        if key == "plant species":
            self._flush()
            self._start()
            self._labels.append(label)
            self._started_by_label = True
            self._position = 0
            return
        if self._current is None:
            self._start()
        if label:
            self._labels.append(label)

        inferred = False
        if key is None:
            if self._position is None or self._position + 1 >= len(TEMPLATE_B_ROW_ORDER):
                return
            key = TEMPLATE_B_ROW_ORDER[self._position + 1]
            inferred = True
        if key in TEMPLATE_B_ROW_ORDER:
            self._position = TEMPLATE_B_ROW_ORDER.index(key)
        if key not in HEADERS:
            return

        value = read_value(key)
        note = None
        if isinstance(value, tuple):
            value, note = value
        if not value or (inferred and not _value_is_plausible(key, value)):
            return
        existing = self._current[key]
        if existing:
            if inferred or existing == value:
                return
            # The same field again: the next shipment, whose first rows were unreadable.
            self._flush()
            self._start()
            self._position = TEMPLATE_B_ROW_ORDER.index(key)
        self._current[key] = value
        if note:
            self._notes.append(note)
        if inferred:
            self._notes.append(
                f'The "{key}" label could not be read, so this value was placed by its row position. Please verify.'
            )

    def finish(self):
        self._flush()
        return self.records, self.warnings

    def _start(self):
        self._current = {header: "" for header in HEADERS}
        self._labels = []
        self._notes = list(self._table_notes)
        self._started_by_label = False

    def _flush(self):
        if self._current is None:
            return
        filled = sum(1 for header in HEADERS if self._current[header])
        # Keep anything that looks like a real shipment; drop stray UI tables.
        if filled >= 2 or (filled and (self._started_by_label or self._current["Sales Order Number"])):
            record, record_warnings = _finish_template_b_record(self._current, self._labels, self._notes)
            self.records.append(record)
            self.warnings.append(record_warnings)
        self._current = None


def _prepare_image(image):
    """Upright RGB image; transparency is flattened onto a contrasting background."""
    image = ImageOps.exif_transpose(image)
    if image.mode in ("RGBA", "LA", "PA") or (image.mode == "P" and "transparency" in image.info):
        rgba = image.convert("RGBA")
        alpha = rgba.getchannel("A")
        if alpha.getextrema()[0] < 255:
            opaque = alpha.point(lambda a: 255 if a > 128 else 0)
            if opaque.getbbox():
                mean = ImageStat.Stat(rgba.convert("L"), mask=opaque).mean[0]
            else:
                mean = 0
            background = Image.new("RGB", rgba.size, (0, 0, 0) if mean > 127 else (255, 255, 255))
            background.paste(rgba, mask=alpha)
            return background
        return rgba.convert("RGB")
    return image.convert("RGB")


def _read_tables(image, layouts, builder: _TemplateBRecordBuilder, raw_lines: list[str]) -> None:
    for layout in layouts:
        scale = _ocr_scale(layout)
        line_height = _line_height(layout)
        notes = []
        if layout["divider_estimated"]:
            notes.append("The table's column divider could not be detected and was estimated. Please verify.")
        builder.new_table(separate_records=True, notes=notes)
        label_left, label_right = layout["label_x"]
        value_left, value_right = layout["value_x"]
        for top, bottom in layout["rows"]:
            label_cell = image.crop((label_left, top, label_right, bottom))
            value_cell = image.crop((value_left, top, value_right, bottom))
            label = _ocr_cell(label_cell, psm=7, scale=scale)
            if label:
                raw_lines.append(label)

            def read_value(field, cell=value_cell):
                value, note = _read_value_cell(field, cell, scale, line_height)
                if value:
                    raw_lines.append(value)
                return value, note

            builder.add_row(label, _template_b_row_key(label), read_value)


def extract_template_b_images(images) -> tuple[list[dict], str, list[list[str]]]:
    """Extract every Template B shipment table found in the given images/pages."""
    _require_ocr()
    builder = _TemplateBRecordBuilder()
    raw_lines = []
    for image in images:
        image = _prepare_image(image)
        found = len(builder.records)
        _read_tables(image, _find_tables(image), builder, raw_lines)
        builder.new_table(separate_records=True)
        if len(builder.records) == found:
            # Nothing read with table finding: try the original whole-image grid.
            legacy = _legacy_layout(image)
            if legacy:
                _read_tables(image, [legacy], builder, raw_lines)
                builder.new_table(separate_records=True)

    records, warnings_by_record = builder.finish()
    if not records:
        raise RuntimeError(
            "No Template B shipment tables could be read from this image. Make sure the "
            "table is fully visible and not cropped, then try again."
        )
    return records, "\n".join(raw_lines), warnings_by_record


def _image_frames(path: str):
    """Every page of an image file (multi-page TIFFs have several)."""
    with Image.open(path) as image:
        for frame in ImageSequence.Iterator(image):
            yield frame.copy()


def extract_template_b_image(path: str) -> tuple[list[dict], str, list[list[str]]]:
    """
    Extract one or more Template B shipment tables from a single image.

    Every table found becomes its own shipment, so one image may contain 1, 2,
    5, or more Template B tables and each becomes a separate record/Excel row.
    """
    _require_ocr()
    return extract_template_b_images(_image_frames(path))


# --- Template B PDFs (printed or exported emails) -----------------------------

_PRIVATE_USE = re.compile(r"[\ue000-\uf8ff]")


def _pdf_cell_text(cell) -> str:
    return _PRIVATE_USE.sub("", str(cell or "")).strip()


def _pdf_label_key(text: str) -> str | None:
    if len(text) > 40 or "\n" in text:
        return None
    cleaned = _clean_label(text)
    squashed = re.sub(r"\s+", "", cleaned)
    for field, labels in _configured_table_labels().items():
        for label in labels:
            label = _clean_label(label)
            if cleaned == label or squashed == re.sub(r"\s+", "", label):
                return field
    return _template_b_row_key(text)


def _split_pdf_row(row) -> tuple[str, str | None, str]:
    """(label, key, value) for one PDF table row; labels split across cells are rejoined."""
    cells = [text for text in (_pdf_cell_text(cell) for cell in row or []) if text]
    if not cells:
        return "", None, ""
    if len(cells) == 1:
        # Either a label whose value cell is empty, or a value whose label was lost.
        key = _pdf_label_key(cells[0])
        return (cells[0], key, "") if key else ("", None, cells[0])
    for count in range(1, len(cells)):
        for label in ("".join(cells[:count]), " ".join(cells[:count])):
            key = _pdf_label_key(label)
            if key:
                return label, key, "\n".join(cells[count:])
    return "", None, cells[-1]


def extract_template_b_pdf_tables(tables: list) -> tuple[list[dict], list[list[str]]]:
    """Read every Template B shipment from tables pdfplumber found in a text PDF."""
    builder = _TemplateBRecordBuilder()
    for table in tables or []:
        # Tables on a PDF page can be fragments of one shipment, so a table
        # boundary only resets the row position.
        builder.new_table(separate_records=False)
        for row in table or []:
            label, key, value = _split_pdf_row(row)
            if not label and not value:
                continue
            builder.add_row(label, key, lambda field, text=value: _clean_template_b_value(field, text))
    return builder.finish()


def _pack_alias_key(value: str) -> str:
    """Canonical key used for Pack Style aliases without changing meaning."""
    value = str(value or "").strip().casefold()
    value = re.sub(r"\s+", " ", value)
    return value


def _smart_pack_words(text: str) -> str:
    """Human-friendly casing for packaging/product words while preserving numbers."""
    text = re.sub(r"\s+", " ", str(text or "").strip())
    if not text:
        return ""
    lower_words = {"of", "with", "and", "the", "roots", "root"}
    known_upper = {"us": "US", "usa": "USA"}
    out = []
    for word in text.split(" "):
        bare = word.strip()
        if not bare:
            continue
        low = bare.casefold()
        if low in known_upper:
            out.append(known_upper[low])
        elif low in lower_words:
            out.append(low)
        elif re.fullmatch(r"\d+ct", low):
            out.append(low)
        elif re.fullmatch(r"\d+x\d+ct", low):
            out.append(low)
        else:
            # Keep abbreviations/codes already containing mixed case or digits.
            if any(ch.isdigit() for ch in bare) or (bare.isupper() and len(bare) <= 4):
                out.append(bare)
            else:
                out.append(bare[:1].upper() + bare[1:].lower())
    return " ".join(out)


def _normalize_pack_segment(segment: str) -> tuple[str, bool]:
    segment = re.sub(r"\s+", " ", segment.strip(" /|,;"))
    if not segment:
        return "", False

    # Compact pack style, e.g. 12x1ctClam, 12 X 1 CT clam, 8x2ct clam.
    compact = re.fullmatch(r"(\d+)\s*[xX×]\s*(\d+)\s*ct\s*(.*)", segment, re.IGNORECASE)
    if compact:
        outer, inner, description = compact.groups()
        description = _smart_pack_words(description)
        value = f"{outer}x{inner}ct"
        if description:
            value += f" {description}"
        return value, True

    # Descriptive case style, e.g. "1 case of 24ct Green Butterhead Lettuce with roots".
    descriptive = re.fullmatch(
        r"(\d+)\s+cases?\s+of\s+(\d+)\s*ct\s+(.+)",
        segment,
        re.IGNORECASE,
    )
    if descriptive:
        cases, count, description = descriptive.groups()
        noun = "case" if int(cases) == 1 else "cases"
        return f"{cases} {noun} of {count}ct {_smart_pack_words(description)}", True

    return _smart_pack_words(segment), False


def normalize_pack_style(value: str) -> tuple[str, str | None]:
    """Normalize Pack Style into consistent human-readable output.

    Deterministic rules are used so source quantities/content are not invented.
    Unknown styles are cleaned but preserved and flagged for review.
    """
    original = str(value or "").strip()
    if not original:
        return original, None

    settings = load_settings()
    formatting = settings.get("formatting", {})
    if not formatting.get("normalize_pack_style", True):
        return original, None

    aliases = formatting.get("pack_style_aliases", {}) or {}
    alias_lookup = {_pack_alias_key(k): str(v).strip() for k, v in aliases.items() if str(k).strip()}
    direct = alias_lookup.get(_pack_alias_key(original))
    if direct:
        return direct, None

    text = original.replace("\r", "\n")
    # PDF table extraction currently joins multiline pack cells with |. Most are
    # line-wrap artifacts, so flatten first, then restore a separator before a
    # repeated "N case(s) of Nct" pack description.
    text = re.sub(r"\s*\|\s*", " ", text)
    text = re.sub(r"(?i)\blettuce\s*-\s*with\b", "Lettuce with", text)
    text = re.sub(r"\s*-\s+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()

    # Standardize explicit separators from source documents.
    separator = str(formatting.get("pack_style_separator", " / ") or " / ")
    text = re.sub(r"\s*/\s*", separator, text)

    # If a second descriptive pack starts without a delimiter, insert one.
    starts = list(re.finditer(r"(?i)\b\d+\s+cases?\s+of\s+\d+\s*ct\b", text))
    if len(starts) > 1:
        pieces = []
        for idx, match in enumerate(starts):
            start = match.start()
            end = starts[idx + 1].start() if idx + 1 < len(starts) else len(text)
            pieces.append(text[start:end].strip())
        prefix = text[:starts[0].start()].strip()
        if prefix:
            pieces[0] = f"{prefix} {pieces[0]}".strip()
    else:
        pieces = [piece.strip() for piece in text.split(separator) if piece.strip()]

    normalized_parts = []
    all_recognized = True
    for piece in pieces:
        normalized, recognized = _normalize_pack_segment(piece)
        if normalized:
            normalized_parts.append(normalized)
        all_recognized = all_recognized and recognized

    normalized_text = separator.join(normalized_parts) if normalized_parts else original

    # Alias lookup again after generic cleanup, allowing admins to map recurring
    # OCR/source variants without code changes.
    mapped = alias_lookup.get(_pack_alias_key(normalized_text))
    if mapped:
        return mapped, None

    if all_recognized:
        return normalized_text, None

    return normalized_text, f'Pack Style "{original}" was cleaned but does not match a known formatting rule. Please verify.'


def normalize_date(value: str) -> tuple[str, str | None]:
    """Normalize recognized shipment dates to 'August 3, 2026' style.

    Returns (normalized_value, warning). Unknown formats are preserved unchanged.
    Numeric dates are interpreted as month-day-year, matching the supplied templates.
    """
    original = str(value or "").strip()
    if not original:
        return original, None

    formats = (
        "%B %d, %Y", "%B %d %Y", "%b %d, %Y", "%b %d %Y",
        "%m-%d-%Y", "%m/%d/%Y", "%Y-%m-%d",
        "%m-%d-%y", "%m/%d/%y",
    )
    for fmt in formats:
        try:
            parsed = datetime.strptime(original, fmt)
            # %-d is not portable on Windows, so construct the day without a leading zero.
            return f"{parsed.strftime('%B')} {parsed.day}, {parsed.year}", None
        except ValueError:
            continue

    return original, f'Date "{original}" could not be normalized; original value was kept.'


def normalize_record(record: dict) -> list[str]:
    """Apply shared output normalization in-place and return any review warnings."""
    warnings = []

    normalized, warning = normalize_date(record.get("Date", ""))
    record["Date"] = normalized
    if warning:
        warnings.append(warning)

    normalized_pack, pack_warning = normalize_pack_style(record.get("Pack Style", ""))
    record["Pack Style"] = normalized_pack
    if pack_warning:
        warnings.append(pack_warning)

    return warnings


# ---------------------------------------------------------------------------
# 5. RECORD VALIDATION / CONFIDENCE
# ---------------------------------------------------------------------------

def validate_record(record: dict, template_name: str, ocr_source: bool | None = None) -> tuple[int, str, list[str]]:
    """
    Return (confidence_percent, status, warnings) for one extracted row.

    Template A is treated as the formatting baseline. Values that do not follow
    the same general shape are not automatically rejected, but confidence is
    reduced and the row is flagged for human verification. Records read with
    OCR receive an additional review penalty. ``ocr_source`` says whether OCR
    was used; when it is None, Template B is assumed to be an OCR'd image.
    """
    warnings = []
    score = 100
    settings = get_template_settings().get(template_name, {})

    # Required fields. Template C's supplied layout does not require a lot code.
    required = ["Date", "Sales Order Number", "Customer and Location", "Pack Style", "Number of Plants"]
    if settings.get("lot_required", template_name != "Template C"):
        required.append("Lot Code")

    missing = [field for field in required if not str(record.get(field, "")).strip()]
    if missing:
        warnings.append("Missing required field(s): " + ", ".join(missing))
        score -= 18 * len(missing)

    # ---- Template A formatting baseline ----
    # Dates should already be normalized to "August 3, 2026" style.
    date_value = str(record.get("Date", "")).strip()
    if date_value:
        try:
            datetime.strptime(date_value, "%B %d, %Y")
        except ValueError:
            warnings.append(
                f'Date "{date_value}" differs from the Template A format (Month D, YYYY). Please verify.'
            )
            score -= 10

    sales_order = str(record.get("Sales Order Number", "")).strip()
    sales_pattern = str(settings.get("sales_order_pattern", r"^\d{5,9}$") or r"^\d{5,9}$")
    if sales_order:
        try:
            valid_sales = re.fullmatch(sales_pattern, sales_order) is not None
        except re.error:
            valid_sales = re.fullmatch(r"\d{5,9}", sales_order) is not None
            warnings.append("Configured Sales Order regex is invalid; default validation was used. Please review Backend settings.")
            score -= 5
        if not valid_sales:
            warnings.append(
                f'Sales Order Number "{sales_order}" does not match the configured format. Please verify.'
            )
            score -= 12

    customer = str(record.get("Customer and Location", "")).strip()
    if customer:
        # Normalized Template A customer/address values are comma-separated and
        # include multiple address components. A shorter/non-separated value may
        # still be valid, but deserves review.
        if customer.count(",") < 2:
            warnings.append(
                "Customer and Location differs from the normalized Template A address format. Please verify."
            )
            score -= 7

    pack_style = str(record.get("Pack Style", "")).strip()
    if pack_style:
        compact_ok = re.fullmatch(r"\d+x\d+ct(?:\s+.+)?", pack_style, re.IGNORECASE) is not None
        descriptive_parts = [part.strip() for part in re.split(r"\s*/\s*", pack_style) if part.strip()]
        descriptive_ok = bool(descriptive_parts) and all(
            re.fullmatch(r"\d+\s+cases?\s+of\s+\d+ct\s+.+", part, re.IGNORECASE)
            for part in descriptive_parts
        )
        if not (compact_ok or descriptive_ok):
            warnings.append(f'Pack Style "{pack_style}" does not match a recognized normalized format. Please verify.')
            score -= 6

    plant_value = str(record.get("Number of Plants", "")).strip()
    if plant_value:
        # Template A normally includes both a total and the case breakdown.
        # Multi-pack shipments list one breakdown per pack, separated by "/":
        # "1520 heads (12 heads x 60 cases / 16 heads x 50 cases)".
        group = r"\d+\s+(?:heads?|plants?)\s*[x×]\s*\d+\s+cases?"
        baseline_plant = re.fullmatch(
            rf"[\d,]+\s+(?:heads?|plants?)\s*\(\s*{group}(?:\s*[/+,&]\s*{group})*\s*\)",
            plant_value,
            re.IGNORECASE,
        )
        if not baseline_plant:
            warnings.append(
                f'Number of Plants "{plant_value}" differs from the full Template A format with case breakdown. Please verify.'
            )
            score -= 8

        parts = _plant_count_parts(plant_value)
        if parts:
            total, groups = parts
            computed = sum(each * cases for each, cases in groups)
            if computed != total:
                formula = " + ".join(f"{each} x {cases}" for each, cases in groups)
                warnings.append(f"Plant-count check failed: {formula} = {computed}, not {total}.")
                score -= 25

            # A compact pack style states the heads per case: 12x1ct = 12,
            # 8x2ct = 16. It must agree with each case breakdown, which
            # catches OCR slips such as 12x7ct that look valid on their own.
            packs = re.findall(r"(\d+)x(\d+)ct\b", pack_style, re.IGNORECASE)
            if packs and len(packs) == len(groups):
                per_case = [int(outer) * int(inner) for outer, inner in packs]
                if per_case != [each for each, _cases in groups]:
                    warnings.append(
                        f'Pack Style "{pack_style}" does not match the heads per case in Number of Plants. Please verify.'
                    )
                    score -= 10

    lot = str(record.get("Lot Code", "")).strip()
    if lot and not re.fullmatch(r"[A-Za-z0-9\-]+", lot):
        warnings.append(f'Lot Code "{lot}" contains unusual characters compared with Template A. Please verify.')
        score -= 8

    # OCR-derived records always receive a visible review reminder rather than
    # presenting OCR as 100% certain.
    app_settings = load_settings()
    ocr_settings = app_settings.get("ocr", {})
    if ocr_source is None:
        ocr_source = template_name == "Template B"
    if template_name == "Template B" and ocr_source:
        warnings.append("Template B was read from an image using OCR. Verify the extracted values against the source image.")
        score -= int(ocr_settings.get("template_b_penalty", 10))
    elif ocr_source:
        warnings.append(SCANNED_PDF_WARNING)
        score -= int(ocr_settings.get("scanned_pdf_penalty", 10))

    # Remove duplicate warning strings while preserving order.
    warnings = list(dict.fromkeys(warnings))
    score = max(0, min(100, score))

    verified_threshold = int(ocr_settings.get("verified_threshold", 95))
    review_threshold = int(ocr_settings.get("review_threshold", 70))
    if score >= verified_threshold and not warnings:
        status = "Verified"
    elif score >= review_threshold:
        status = "Review"
    else:
        status = "Attention"
    return score, status, warnings


# ---------------------------------------------------------------------------
# 6. PDF / TOP-LEVEL ENTRY POINT
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


# PDFs with less selectable text than this are treated as scanned images.
SCANNED_PDF_TEXT_THRESHOLD = 25
SCANNED_PDF_WARNING = (
    "Read from a scanned PDF using OCR. Verify the extracted values against the source."
)


def ocr_pdf_pages(path: str, resolution: int = 300) -> str:
    """OCR every page of a scanned PDF and return the combined text.

    Pages are rendered by pdfplumber (via pypdfium2), so Poppler is not needed.
    """
    if pytesseract is None or ImageOps is None:
        raise RuntimeError(
            "Scanned PDF support requires the Python packages Pillow and pytesseract. "
            "Install them with: pip install pillow pytesseract"
        )
    if not TESSERACT_PATH:
        raise RuntimeError(
            "This PDF has no selectable text and Tesseract OCR was not found to read it. "
            "Install Tesseract OCR for Windows, or set the TESSERACT_CMD environment "
            "variable to the full path to tesseract.exe."
        )

    page_texts = []
    for image in render_pdf_pages(path, resolution):
        gray = ImageOps.autocontrast(image.convert("L"))
        page_texts.append(pytesseract.image_to_string(gray, config="--psm 6"))
    return "\n".join(page_texts)


def render_pdf_pages(path: str, resolution: int = 300) -> list:
    """Render every PDF page to a PIL image (via pypdfium2; Poppler is not needed)."""
    with pdfplumber.open(path) as pdf:
        return [page.to_image(resolution=resolution).original for page in pdf.pages]


# Characters OCR can read as each other in codes (I/1, O/0, S/5, B/8, Z/2).
_LOOKALIKES = str.maketrans("IOSBZ", "10582")


def lookalike_key(code: str) -> str:
    """Comparison key under which codes that differ only by look-alike characters match."""
    return str(code or "").strip().upper().translate(_LOOKALIKES)


def lookalike_note(code: str, other: str, where: str) -> str:
    return (
        f'Lot Code "{code}" differs from "{other}" {where} only in characters OCR can confuse '
        f'(I/1, O/0, S/5, B/8, Z/2). Please verify against the source.'
    )


def known_lot_code_note(code: str, known: set) -> str | None:
    """Note when an OCR'd lot code is new but a look-alike of a known one (e.g. from the workbook)."""
    code = str(code or "").strip()
    if not code or code in known:
        return None
    key = lookalike_key(code)
    matches = sorted(value for value in known if lookalike_key(value) == key)
    return lookalike_note(code, matches[0], "already in the workbook") if matches else None


def _flag_lookalike_lot_codes(records: list, warnings_by_record: list) -> None:
    """Flag lot codes in one file that differ only by look-alike characters (likely an OCR slip)."""
    by_key = {}
    for index, record in enumerate(records):
        code = str(record.get("Lot Code", "") or "").strip()
        if code:
            by_key.setdefault(lookalike_key(code), []).append((index, code))
    for entries in by_key.values():
        codes = {code for _index, code in entries}
        if len(codes) < 2:
            continue
        for index, code in entries:
            other = sorted(codes - {code})[0]
            warnings_by_record[index].append(lookalike_note(code, other, "in another table of this file"))


def _set_template_b_records(out: dict, records: list, warnings_by_record: list, ocr_source: bool, source: str) -> None:
    """Normalize, validate and store several Template B shipments in a process_pdf result."""
    out["template"] = "Template B"
    out["records"] = records
    out["ocr_source"] = ocr_source
    warnings_by_record = [list(warnings) for warnings in warnings_by_record]
    warnings_by_record += [[] for _ in range(len(records) - len(warnings_by_record))]
    if ocr_source:
        _flag_lookalike_lot_codes(records, warnings_by_record)
    verified_threshold = int(load_settings().get("ocr", {}).get("verified_threshold", 95))
    validated_warnings = []
    confidences = []
    statuses = []
    for idx, record in enumerate(records):
        normalization_warnings = normalize_record(record)
        confidence, status, validation_warnings = validate_record(record, "Template B", ocr_source)
        base = warnings_by_record[idx] if idx < len(warnings_by_record) else []
        if base and confidence >= verified_threshold:
            # Extraction-time warnings (e.g. a label placed by row position)
            # must still send the record through manual review.
            confidence = verified_threshold - 1
            status = "Review"
        validated_warnings.append(list(dict.fromkeys(base + normalization_warnings + validation_warnings)))
        confidences.append(confidence)
        statuses.append(status)
    out["record_warnings"] = validated_warnings
    out["record_confidence"] = confidences
    out["record_status"] = statuses
    out["fields"] = records[0] if records else {h: "" for h in HEADERS}
    if len(records) > 1:
        out["warnings"].append(f"Detected {len(records)} Template B shipment tables in this {source}.")


def _template_b_from_pdf_pages(path: str):
    """Last resort for PDFs: render the pages and look for Template B tables in them."""
    if pytesseract is None or not TESSERACT_PATH:
        return None
    try:
        records, raw_text, warnings_by_record = extract_template_b_images(render_pdf_pages(path))
    except RuntimeError:
        return None
    return records, raw_text, warnings_by_record


def process_pdf(path: str) -> dict:
    """
    Process one source file.

    The result always includes ``records``. PDFs normally produce one record;
    Template B images can produce multiple records when several shipment tables
    are stacked in the same image. ``fields`` remains as a compatibility alias
    for the first record.
    """
    out = {
        "source_file": path,
        "template": "",
        "fields": {h: "" for h in HEADERS},
        "records": [],
        "record_warnings": [],
        "record_confidence": [],
        "record_status": [],
        "warnings": [],
        "raw_text": "",
        "ocr_source": False,
        "error": None,
    }

    try:
        extension = Path(path).suffix.lower()

        if extension in IMAGE_EXTENSIONS:
            records, raw_text, warnings_by_record = extract_template_b_image(path)
            out["raw_text"] = raw_text
            _set_template_b_records(out, records, warnings_by_record, ocr_source=True, source="image")

        else:
            text, tables = read_pdf(path)
            ocr_source = False
            if len(text.strip()) < SCANNED_PDF_TEXT_THRESHOLD:
                if not load_settings().get("ocr", {}).get("ocr_scanned_pdfs", True):
                    out["error"] = "No extractable text found (PDF may be scanned/image-based)."
                    return out
                # Scanned PDF: OCR the rendered pages and rely on the text rules,
                # since OCR does not produce the table structure pdfplumber reads.
                text = ocr_pdf_pages(path)
                tables = None
                ocr_source = True
                if not text.strip():
                    out["error"] = "No extractable text found, and OCR could not read any text from this PDF."
                    return out
            out["raw_text"] = text
            out["ocr_source"] = ocr_source

            template_name = detect_template(text)
            out["template"] = template_name

            if template_name == "Template B":
                # Template B PDFs (printed/exported emails, scans) can hold any
                # number of shipment tables; read each one separately.
                if ocr_source:
                    found = _template_b_from_pdf_pages(path)
                    if found:
                        records, _raw, warnings_by_record = found
                        _set_template_b_records(out, records, warnings_by_record, ocr_source=True, source="PDF")
                        return out
                else:
                    records, warnings_by_record = extract_template_b_pdf_tables(tables)
                    if any(sum(1 for h in HEADERS if record[h]) >= 3 for record in records):
                        _set_template_b_records(out, records, warnings_by_record, ocr_source=False, source="PDF")
                        return out

            fields = extract_fields(text, template_name, tables)
            if not ocr_source and sum(1 for value in fields.values() if value) < 2:
                # Almost nothing in the PDF text: the shipment table may be a
                # picture inside the PDF (e.g. a pasted screenshot in an email).
                found = _template_b_from_pdf_pages(path)
                if found:
                    records, _raw, warnings_by_record = found
                    _set_template_b_records(out, records, warnings_by_record, ocr_source=True, source="PDF")
                    return out
            normalization_warnings = normalize_record(fields)
            out["fields"] = fields
            out["records"] = [fields]

            record_warnings = list(normalization_warnings)
            missing = [field for field, value in fields.items() if not value]
            if template_name == "Template C":
                missing = [field for field in missing if field != "Lot Code"]
            if missing:
                record_warnings.append("Could not find: " + ", ".join(missing))

            if template_name == "Template C" and not fields.get("Lot Code"):
                record_warnings.append(
                    "Template C does not require a lot code; Lot Code was left blank."
                )

            if template_name == "Generic":
                record_warnings.append(
                    "Template not recognized (using generic rules) - please review."
                )

            confidence, status, validation_warnings = validate_record(fields, template_name, ocr_source)
            record_warnings = list(dict.fromkeys(record_warnings + validation_warnings))
            out["record_warnings"] = [record_warnings]
            out["record_confidence"] = [confidence]
            out["record_status"] = [status]
            out["warnings"].extend(record_warnings)

    except Exception as e:
        out["error"] = f"Failed to process file: {e}"

    return out


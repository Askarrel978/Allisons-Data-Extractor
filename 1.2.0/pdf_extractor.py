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
import sys
import logging
from pathlib import Path
from datetime import datetime

import pdfplumber

from template_config import get_template_markers, get_template_settings
from settings_manager import load_settings

logging.getLogger("pdfminer").setLevel(logging.ERROR)
logging.getLogger("pdfminer.pdffont").setLevel(logging.ERROR)

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


def _ocr_digits_cell(image) -> str:
    """OCR a short numeric cell such as a sales order using a digits-only pass."""
    if pytesseract is None or ImageOps is None:
        return ""
    gray = image.convert("L")
    gray = ImageOps.autocontrast(gray)
    scale = 6
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
    """Detect the bright horizontal/vertical rules in Template B's table."""
    gray = image.convert("L")
    width, height = gray.size
    pix = gray.load()

    # The screenshot uses light grid lines on a dark background. Thresholds
    # are administrator-configurable through Advanced Configuration JSON.
    grid_settings = load_settings().get("extraction", {}).get("template_b_grid", {})
    bright_threshold = int(grid_settings.get("bright_threshold", 100))
    horizontal_coverage = float(grid_settings.get("horizontal_coverage", 0.75))
    vertical_coverage = float(grid_settings.get("vertical_coverage", 0.70))

    horizontal = []
    for y in range(height):
        bright = sum(1 for x in range(width) if pix[x, y] >= bright_threshold)
        if bright >= int(width * horizontal_coverage):
            horizontal.append(y)

    vertical = []
    for x in range(width):
        bright = sum(1 for y in range(height) if pix[x, y] >= bright_threshold)
        if bright >= int(height * vertical_coverage):
            vertical.append(x)

    return _group_nearby(horizontal), _group_nearby(vertical)


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


def _finish_template_b_record(record: dict, labels_seen: list[str], warnings: list[str]) -> tuple[dict, list[str]]:
    """Clean one Template B shipment and attach record-specific warnings."""
    record = dict(record)

    # Small targeted correction for a common OCR error in this exact packaging format.
    # e.g. "12xict clam" -> "12x1ct clam". This does not alter other pack styles.
    pack = record.get("Pack Style", "")
    if pack:
        pack = re.sub(r"(?i)(\d+)x[iIlL]ct\b", r"\1x1ct", pack)
        # Another recurring OCR read on the compact multi-table screenshots:
        # "12x1ct dam" is the visible "12x1ct clam" value.
        pack = re.sub(r"(?i)\b(\d+x1ct)\s+d[ai]m\b", r"\1 clam", pack)
        record["Pack Style"] = pack

    record_warnings = list(warnings)
    marker_found = any("lot#" in _clean_label(label) or "lot #" in _clean_label(label) for label in labels_seen)
    if not marker_found:
        record_warnings.append('Template B record found, but the "LOT#" marker was not read clearly by OCR.')

    missing = [field for field, value in record.items() if field in HEADERS and not value]
    if missing:
        record_warnings.append("Could not find: " + ", ".join(missing))

    return record, record_warnings


def extract_template_b_image(path: str) -> tuple[list[dict], str, list[list[str]]]:
    """
    Extract one or more Template B shipment tables from a single image.

    Every OCR'd row whose label reads "Plant Species" starts a new shipment.
    Therefore one image may contain 1, 2, 5, or more Template B tables and each
    table becomes a separate record/Excel row.
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

    width, height = image.size
    divider_candidates = [x for x in vertical if width * 0.20 < x < width * 0.75]
    divider = divider_candidates[0] if divider_candidates else vertical[0]
    right_edge = max(x for x in vertical if x > divider) if any(x > divider for x in vertical) else width - 1

    records = []
    warnings_by_record = []
    raw_lines = []

    current = None
    current_labels = []
    current_warnings = []

    for top, bottom in zip(horizontal, horizontal[1:]):
        if bottom - top < 8:
            continue

        left_cell = image.crop((1, top + 1, max(2, divider - 1), bottom - 1))
        right_cell = image.crop((divider + 1, top + 1, max(divider + 2, right_edge - 1), bottom - 1))

        label = _ocr_cell(left_cell, psm=7)
        value = _ocr_cell(right_cell, psm=6)

        if label:
            raw_lines.append(label)
        if value:
            raw_lines.append(value)

        # A repeated Plant Species row is the boundary between shipment tables.
        if label and _is_template_b_record_start(label):
            if current is not None:
                finished, record_warnings = _finish_template_b_record(
                    current, current_labels, current_warnings
                )
                records.append(finished)
                warnings_by_record.append(record_warnings)

            current = {header: "" for header in HEADERS}
            current_labels = [label]
            current_warnings = []
            continue

        # Ignore table/gap OCR until the first Plant Species row is located.
        if current is None:
            continue

        if label:
            current_labels.append(label)

        field = _template_b_field_from_label(label)

        # Sales-order cells are short numeric values. OCR them again with a
        # digits-only whitelist because small white digits on the dark Template B
        # screenshots can occasionally disappear or be misread during the normal
        # OCR pass.
        if field == "Sales Order Number":
            numeric_value = _ocr_digits_cell(right_cell)
            if numeric_value:
                value = numeric_value

        if field and value:
            current[field] = _clean_value(field, value)
            continue

        # Positional fallback for Template B: Sales order is the row immediately
        # after LOT#. If OCR mangles the left-hand "Sales order #:" label, keep
        # the row by recognizing its 5-9 digit value after a lot code has already
        # been found. Re-OCR the right cell using digits only before giving up.
        if (
            current.get("Lot Code")
            and not current.get("Sales Order Number")
            and not current.get("Customer and Location")
        ):
            numeric_value = _ocr_digits_cell(right_cell)
            if not numeric_value:
                compact_value = re.sub(r"\D", "", value or "")
                if 5 <= len(compact_value) <= 9:
                    numeric_value = compact_value
            if numeric_value:
                current["Sales Order Number"] = numeric_value

    if current is not None:
        finished, record_warnings = _finish_template_b_record(
            current, current_labels, current_warnings
        )
        records.append(finished)
        warnings_by_record.append(record_warnings)

    if not records:
        raise RuntimeError(
            'Template B table was detected, but no "Plant Species" record starts were recognized.'
        )

    raw_text = "\n".join(raw_lines)
    return records, raw_text, warnings_by_record


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

def validate_record(record: dict, template_name: str) -> tuple[int, str, list[str]]:
    """
    Return (confidence_percent, status, warnings) for one extracted row.

    Template A is treated as the formatting baseline. Values that do not follow
    the same general shape are not automatically rejected, but confidence is
    reduced and the row is flagged for human verification. Template B receives
    an additional OCR review penalty because it comes from an image rather than
    selectable PDF text.
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
        baseline_plant = re.fullmatch(
            r"[\d,]+\s+(?:heads?|plants?)\s*\(\s*\d+\s+(?:heads?|plants?)\s*[x×]\s*\d+\s+cases?\s*\)",
            plant_value,
            re.IGNORECASE,
        )
        if not baseline_plant:
            warnings.append(
                f'Number of Plants "{plant_value}" differs from the full Template A format with case breakdown. Please verify.'
            )
            score -= 8

        total_match = re.search(r"([\d,]+)\s*(?:heads?|plants?)?", plant_value, re.IGNORECASE)
        math_match = re.search(r"\((\d+)\s*(?:heads?|plants?)?\s*[x×]\s*(\d+)\s*cases?\)", plant_value, re.IGNORECASE)
        if total_match and math_match:
            total = int(total_match.group(1).replace(",", ""))
            each = int(math_match.group(1))
            cases = int(math_match.group(2))
            if each * cases != total:
                warnings.append(f"Plant-count check failed: {each} x {cases} = {each * cases}, not {total}.")
                score -= 25

    lot = str(record.get("Lot Code", "")).strip()
    if lot and not re.fullmatch(r"[A-Za-z0-9\-]+", lot):
        warnings.append(f'Lot Code "{lot}" contains unusual characters compared with Template A. Please verify.')
        score -= 8

    # Template B is OCR-derived, so even a perfectly shaped record should receive
    # a visible review reminder rather than presenting image OCR as 100% certain.
    app_settings = load_settings()
    ocr_settings = app_settings.get("ocr", {})
    if template_name == "Template B":
        warnings.append("Template B was read from an image using OCR. Verify the extracted values against the source image.")
        score -= int(ocr_settings.get("template_b_penalty", 10))

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
        "error": None,
    }

    try:
        extension = Path(path).suffix.lower()

        if extension in IMAGE_EXTENSIONS:
            records, raw_text, warnings_by_record = extract_template_b_image(path)
            out["raw_text"] = raw_text
            out["template"] = "Template B"
            out["records"] = records
            validated_warnings = []
            confidences = []
            statuses = []
            for idx, record in enumerate(records):
                normalization_warnings = normalize_record(record)
                confidence, status, validation_warnings = validate_record(record, "Template B")
                base = warnings_by_record[idx] if idx < len(warnings_by_record) else []
                validated_warnings.append(list(dict.fromkeys(base + normalization_warnings + validation_warnings)))
                confidences.append(confidence)
                statuses.append(status)
            out["record_warnings"] = validated_warnings
            out["record_confidence"] = confidences
            out["record_status"] = statuses
            out["fields"] = records[0] if records else {h: "" for h in HEADERS}

            if len(records) > 1:
                out["warnings"].append(
                    f"Detected {len(records)} Template B shipment tables in this image."
                )

        else:
            text, tables = read_pdf(path)
            out["raw_text"] = text
            if not text.strip():
                out["error"] = "No extractable text found (PDF may be scanned/image-based)."
                return out

            template_name = detect_template(text)
            out["template"] = template_name
            fields = extract_fields(text, template_name, tables)
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

            confidence, status, validation_warnings = validate_record(fields, template_name)
            record_warnings = list(dict.fromkeys(record_warnings + validation_warnings))
            out["record_warnings"] = [record_warnings]
            out["record_confidence"] = [confidence]
            out["record_status"] = [status]
            out["warnings"].extend(record_warnings)

    except Exception as e:
        out["error"] = f"Failed to process file: {e}"

    return out


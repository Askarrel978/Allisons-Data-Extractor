"""Template B: any number of shipment tables per image, in any layout."""

import pytest
from PIL import Image, ImageDraw, ImageFont

import pdf_extractor
from pdf_extractor import (
    _find_tables,
    _fix_pack_ocr,
    _TemplateBRecordBuilder,
    extract_template_b_pdf_tables,
    process_pdf,
)
from template_b_factory import DARK, EXPECTED, LIGHT, SHIPMENTS, compose, draw_table

needs_ocr = pytest.mark.skipif(
    pdf_extractor.pytesseract is None or not pdf_extractor.TESSERACT_PATH,
    reason="Tesseract OCR is not available",
)


def _double_size(image):
    """Like a screenshot taken on a high-DPI (200% scaling) display."""
    return image.resize((image.width * 2, image.height * 2), Image.Resampling.LANCZOS)


def _with_email_around(image):
    """Surround tables with email text, a separator rule and a pane border."""
    font = ImageFont.load_default(size=15)
    canvas = Image.new("RGB", (image.width + 300, image.height + 160), DARK["background"])
    draw = ImageDraw.Draw(canvas)
    draw.text((20, 15), "From: Shipping <shipping@example.com>  Sent: Friday, August 21, 2026", fill="white", font=font)
    draw.text((20, 40), "Hi, please see this weekend's phytos below. Thanks!", fill="white", font=font)
    draw.line((0, 70, canvas.width, 70), fill=(120, 120, 120))
    draw.line((image.width + 200, 0, image.width + 200, canvas.height), fill=(120, 120, 120))
    canvas.paste(image, (40, 100))
    draw.text((20, image.height + 120), "Regards, the shipping team", fill="white", font=font)
    return canvas


# --- Table finding (no OCR needed) ---------------------------------------------

def test_finds_stacked_tables_of_different_widths_and_offsets():
    image = compose(
        [draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[1], width=520), draw_table(SHIPMENTS[2], width=380)],
        offsets=[60, 0, 110],
        gap=40,
    )
    layouts = _find_tables(image)
    assert len(layouts) == 3
    assert [len(layout["rows"]) for layout in layouts] == [10, 10, 10]
    assert all(not layout["divider_estimated"] for layout in layouts)


def test_finds_side_by_side_tables_left_to_right():
    image = compose([draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[1])], horizontal=True, offsets=[0, 25])
    layouts = _find_tables(image)
    assert len(layouts) == 2
    assert layouts[0]["label_x"][0] < layouts[1]["label_x"][0]


def test_finds_tables_in_light_theme():
    image = compose([draw_table(s, theme=LIGHT) for s in SHIPMENTS[:2]], theme=LIGHT)
    assert len(_find_tables(image)) == 2


def test_ignores_email_text_and_window_rules():
    image = _with_email_around(compose([draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[1])]))
    layouts = _find_tables(image)
    assert len(layouts) == 2
    assert all(len(layout["rows"]) == 10 for layout in layouts)


def test_touching_tables_keep_all_rows():
    image = compose([draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[1])], gap=0)
    # Touching tables may merge into one ruled block; the Plant Species rows
    # still separate the shipments when they are read.
    assert sum(len(layout["rows"]) for layout in _find_tables(image)) == 20


def test_image_without_tables_finds_nothing():
    image = Image.new("RGB", (600, 400), (30, 30, 30))
    ImageDraw.Draw(image).text((20, 20), "No tables here", fill="white")
    assert _find_tables(image) == []


# --- Shipment assembly (no OCR needed) --------------------------------------------

def _rows(builder, rows):
    for label, value in rows:
        key = pdf_extractor._template_b_row_key(label)
        builder.add_row(label, key, lambda field, text=value: pdf_extractor._clean_template_b_value(field, text))


def test_plant_species_rows_split_shipments_within_one_table():
    builder = _TemplateBRecordBuilder()
    builder.new_table()
    _rows(builder, [
        ("Plant Species:", "Lettuce"), ("Sales order #:", "111111"), ("LOT#", "AB123"),
        ("Plant Species:", "Lettuce"), ("Sales order #:", "222222"), ("LOT#", "CD456"),
    ])
    records, _warnings = builder.finish()
    assert [(r["Sales Order Number"], r["Lot Code"]) for r in records] == [("111111", "AB123"), ("222222", "CD456")]


def test_each_table_is_its_own_shipment_even_without_plant_species():
    builder = _TemplateBRecordBuilder()
    for so in ("111111", "222222"):
        builder.new_table()
        _rows(builder, [("???", "Lettuce"), ("Sales order #:", so), ("Shipping Date:", "August 22, 2026")])
    records, _warnings = builder.finish()
    assert [r["Sales Order Number"] for r in records] == ["111111", "222222"]


def test_unreadable_label_is_placed_by_row_position_when_value_fits():
    builder = _TemplateBRecordBuilder()
    builder.new_table()
    _rows(builder, [
        ("Plant Species:", "Lettuce"),
        ("Country of Origin:", "Canada"),
        ("", "12x1ct clam"),                       # Packaging style label lost
        ("N#mb3r 0f P!ants", "240 heads (12 heads x 20 cases)"),  # garbled label
        ("Shipping Date:", "August 22, 2026"),
    ])
    records, warnings = builder.finish()
    assert records[0]["Pack Style"] == "12x1ct clam"
    assert records[0]["Number of Plants"] == "240 heads (12 heads x 20 cases)"
    assert sum("placed by its row position" in w for w in warnings[0]) == 2


def test_unreadable_label_is_not_guessed_when_value_does_not_fit():
    builder = _TemplateBRecordBuilder()
    builder.new_table()
    _rows(builder, [("Plant Species:", "Lettuce"), ("Shipping Date:", "August 22, 2026"), ("", "Canada"), ("Sales order #:", "111111")])
    records, _warnings = builder.finish()
    assert records[0]["Lot Code"] == ""


def test_repeated_field_starts_the_next_shipment():
    builder = _TemplateBRecordBuilder()
    builder.new_table()
    _rows(builder, [
        ("Sales order #:", "111111"), ("Customer and Address:", "A Co\n1 Main St\nTown ST 1\nUS"),
        ("Sales order #:", "222222"), ("Customer and Address:", "B Co\n2 Main St\nTown ST 2\nUS"),
    ])
    records, _warnings = builder.finish()
    assert [r["Sales Order Number"] for r in records] == ["111111", "222222"]


def test_stray_table_with_one_unrelated_value_is_dropped():
    builder = _TemplateBRecordBuilder()
    builder.new_table()
    _rows(builder, [("Inbox", "6598"), ("Shipping Date:", "August 22, 2026")])
    records, _warnings = builder.finish()
    assert records == []


def test_pdf_tables_with_split_labels_and_cut_off_second_table():
    """Shaped like an Outlook email printed to PDF (see 'Akhil - B1.pdf')."""
    first = [
        ["", None, "Pla", None, "nt Species:", "Lettuce (Lactuca sativa)"],
        [None, "", "Pla", None, None, None],
        [None, None, "\nCo", "u", "ntry of Origin:", "Canada"],
        [None, None, "", None, None, "12x1ct clam"],
        [None, None, "Number of Plants:", None, None, "360 heads (12 heads x 30 cases)"],
        [None, None, "Shipping Date:", None, None, "August 22, 2026"],
        [None, None, "LOT#", None, None, "AB123"],
        [None, None, "Sales order #:", None, None, "787149"],
        [None, None, "Customer and Address:", None, None, "Example Produce\n131 Example Rd.\nTown VA 24120\nUS"],
        [None, None, "Grown at:", None, None, "Example Farms"],
    ]
    second = [
        ["Plant Species:", "Lettuce"], ["Packaging style:", "12x1ct clam"],
        ["Number of Plants:", "240 heads (12 heads x 20 cases)"], ["Shipping Date:", "August 22, 2026"], ["LOT#", ""],
    ]
    outlook_chrome = [["", "Outlook Search"], [None, "Inbox 6598\nSent Items"]]
    records, warnings = extract_template_b_pdf_tables([outlook_chrome, first, second])
    assert len(records) == 2
    assert records[0] == {
        "Date": "August 22, 2026", "Sales Order Number": "787149",
        "Customer and Location": "Example Produce, 131 Example Rd., Town VA 24120, US",
        "Pack Style": "12x1ct clam", "Number of Plants": "360 heads (12 heads x 30 cases)", "Lot Code": "AB123",
    }
    assert records[1]["Sales Order Number"] == ""
    assert any("Could not find" in w for w in warnings[1])


def test_lookalike_lot_codes_in_one_file_are_flagged():
    records = [{"Lot Code": "MUKFI232"}, {"Lot Code": "MUKF1232"}, {"Lot Code": "AB999"}]
    warnings = [[], [], []]
    pdf_extractor._flag_lookalike_lot_codes(records, warnings)
    assert "MUKF1232" in warnings[0][0] and "MUKFI232" in warnings[1][0]
    assert warnings[2] == []


def test_lot_code_compared_with_workbook():
    known = {"MUKFI232", "MU01233"}
    assert pdf_extractor.known_lot_code_note("MUKFI232", known) is None  # already known
    assert pdf_extractor.known_lot_code_note("ZZ123", known) is None  # genuinely new
    note = pdf_extractor.known_lot_code_note("MUKF1232", known)
    assert note and '"MUKFI232"' in note and "already in the workbook" in note
    assert "MU01233" in pdf_extractor.known_lot_code_note("MUO1233", known)


@pytest.mark.parametrize("raw, fixed", [
    ("12xict dam", "12x1ct clam"),
    ("12x1ict dam / 8x2ct cam", "12x1ct clam / 8x2ct clam"),
    ("12xlct clam", "12x1ct clam"),
    ("8x2ict clam", "8x2ct clam"),
    ("1 case of 24ct Green Lettuce", "1 case of 24ct Green Lettuce"),
])
def test_pack_style_ocr_corrections(raw, fixed):
    assert _fix_pack_ocr(raw) == fixed


# --- End to end with OCR ------------------------------------------------------------

VOTED_FIELDS = ("Sales Order Number", "Lot Code", "Date", "Customer and Location")


def _check(path, expected):
    """Every table is found, every value is right, and any OCR doubt is flagged.

    Identity fields may occasionally be misread by OCR (e.g. I vs 1); that is
    acceptable only when the record carries a warning about that field.
    """
    result = process_pdf(str(path))
    assert result["error"] is None
    assert result["template"] == "Template B"
    assert result["ocr_source"] is True
    assert len(result["records"]) == len(expected)
    for record, warnings, want in zip(result["records"], result["record_warnings"], expected):
        for field, value in want.items():
            if record[field] == value:
                continue
            flagged = any(field in warning for warning in warnings)
            assert field in VOTED_FIELDS and flagged, (field, record[field], value, warnings)
    assert all(status != "Verified" for status in result["record_status"])


@needs_ocr
@pytest.mark.parametrize("name, build, expected", [
    ("single", lambda: compose([draw_table(SHIPMENTS[0])]), EXPECTED[:1]),
    ("three stacked, mixed widths", lambda: compose(
        [draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[1], width=520), draw_table(SHIPMENTS[2], width=380)],
        offsets=[60, 0, 110], gap=40), EXPECTED),
    ("side by side", lambda: compose(
        [draw_table(SHIPMENTS[2]), draw_table(SHIPMENTS[0])], horizontal=True, offsets=[0, 25]),
        [EXPECTED[2], EXPECTED[0]]),
    ("light theme", lambda: compose([draw_table(s, theme=LIGHT) for s in SHIPMENTS[:2]], theme=LIGHT), EXPECTED[:2]),
    ("Plant Species label unreadable", lambda: compose(
        [draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[1], label="")]), EXPECTED[:2]),
    ("high resolution", lambda: _double_size(compose([draw_table(SHIPMENTS[1]), draw_table(SHIPMENTS[2])])), EXPECTED[1:]),
    ("inside an email screenshot", lambda: _with_email_around(
        compose([draw_table(SHIPMENTS[1]), draw_table(SHIPMENTS[0])])), [EXPECTED[1], EXPECTED[0]]),
])
def test_reads_every_table(tmp_path, name, build, expected):
    path = tmp_path / "shipments.png"
    build().save(path)
    _check(path, expected)


@needs_ocr
def test_transparent_png_is_flattened(tmp_path):
    image = compose([draw_table(SHIPMENTS[0], theme=LIGHT)], theme=LIGHT).convert("RGBA")
    # Make the white background fully transparent.
    alpha = image.convert("L").point(lambda p: 0 if p == 255 else 255)
    image.putalpha(alpha)
    path = tmp_path / "transparent.png"
    image.save(path)
    _check(path, EXPECTED[:1])


@needs_ocr
def test_image_without_tables_reports_clear_error(tmp_path):
    path = tmp_path / "blank.png"
    Image.new("RGB", (500, 300), "white").save(path)
    result = process_pdf(str(path))
    assert "No Template B shipment tables could be read" in result["error"]


@needs_ocr
def test_scanned_pdf_with_two_template_b_tables(tmp_path):
    """A 300 dpi scan: the text is large and crisp, as on a real scanner."""
    tables = compose([draw_table(SHIPMENTS[0], zoom=3), draw_table(SHIPMENTS[2], zoom=3)], gap=90, margin=150)
    page = Image.new("RGB", (2550, 3300), DARK["background"])
    page.paste(tables, (0, 0))
    path = tmp_path / "scan.pdf"
    page.save(path, "PDF", resolution=300)
    _check(path, [EXPECTED[0], EXPECTED[2]])


@needs_ocr
def test_degraded_image_is_never_silently_wrong(tmp_path):
    """Heavily blurred input: every identity value must be right or flagged for review."""
    small = compose([draw_table(SHIPMENTS[0]), draw_table(SHIPMENTS[2])])
    blurred = small.resize((small.width * 17 // 10, small.height * 17 // 10), Image.Resampling.BILINEAR)
    blurred = blurred.resize((blurred.width * 3, blurred.height * 3), Image.Resampling.BILINEAR)
    path = tmp_path / "blurred.png"
    blurred.save(path)
    result = process_pdf(str(path))
    assert len(result["records"]) == 2
    for record, warnings, expected in zip(result["records"], result["record_warnings"], [EXPECTED[0], EXPECTED[2]]):
        for field in ("Sales Order Number", "Lot Code", "Date"):
            flagged = any(f"OCR readings of {field} disagreed" in warning for warning in warnings)
            assert record[field] == expected[field] or flagged, (field, record[field], warnings)

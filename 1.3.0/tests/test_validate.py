from pdf_extractor import SCANNED_PDF_WARNING, validate_record


def test_clean_template_a_record_is_verified(sample_record):
    assert validate_record(sample_record, "Template A") == (100, "Verified", [])


def test_plant_count_arithmetic_mismatch(sample_record):
    sample_record["Number of Plants"] = "250 heads (12 heads x 20 cases)"
    score, status, warnings = validate_record(sample_record, "Template A")
    assert any("Plant-count check failed: 12 x 20 = 240, not 250" in w for w in warnings)
    assert score == 75
    assert status == "Review"


def test_multi_pack_plant_count_is_checked(sample_record):
    sample_record["Pack Style"] = "12x1ct Clam / 8x2ct Clam"
    sample_record["Number of Plants"] = "1520 heads (12 heads x 60 cases / 16 heads x 50 cases)"
    assert validate_record(sample_record, "Template A") == (100, "Verified", [])

    sample_record["Number of Plants"] = "1530 heads (12 heads x 60 cases / 16 heads x 50 cases)"
    _score, _status, warnings = validate_record(sample_record, "Template A")
    assert "Plant-count check failed: 12 x 60 + 16 x 50 = 1520, not 1530." in warnings


def test_template_b_from_text_pdf_has_no_ocr_penalty(sample_record):
    assert validate_record(sample_record, "Template B", ocr_source=False) == (100, "Verified", [])


def test_missing_fields_reduce_confidence(sample_record):
    sample_record["Sales Order Number"] = ""
    sample_record["Lot Code"] = ""
    score, status, warnings = validate_record(sample_record, "Template A")
    assert "Missing required field(s): Sales Order Number, Lot Code" in warnings
    assert score == 64
    assert status == "Attention"


def test_template_c_does_not_require_lot_code(sample_record):
    sample_record["Lot Code"] = ""
    assert validate_record(sample_record, "Template C") == (100, "Verified", [])


def test_scanned_pdf_penalty_prevents_verified(sample_record):
    score, status, warnings = validate_record(sample_record, "Template A", ocr_source=True)
    assert score == 90
    assert status == "Review"
    assert warnings == [SCANNED_PDF_WARNING]


def test_template_b_penalty_is_not_applied_twice(sample_record):
    score, _status, warnings = validate_record(sample_record, "Template B", ocr_source=True)
    assert score == 90
    assert SCANNED_PDF_WARNING not in warnings
    assert any("Template B was read from an image" in w for w in warnings)


def test_invalid_sales_order_format(sample_record):
    sample_record["Sales Order Number"] = "SO-12"
    score, _status, warnings = validate_record(sample_record, "Template A")
    assert score == 88
    assert any("does not match the configured format" in w for w in warnings)

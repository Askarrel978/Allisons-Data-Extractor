import pytest

from pdf_extractor import normalize_date, normalize_pack_style


@pytest.mark.parametrize("raw", ["08/03/2026", "8/3/2026", "08-03-2026", "2026-08-03", "Aug 3, 2026", "August 03 2026", "08/03/26"])
def test_dates_normalize_to_month_day_year(raw):
    assert normalize_date(raw) == ("August 3, 2026", None)


def test_unrecognized_date_is_kept_with_warning():
    value, warning = normalize_date("next tuesday")
    assert value == "next tuesday"
    assert warning and "could not be normalized" in warning


def test_blank_date_has_no_warning():
    assert normalize_date("") == ("", None)


@pytest.mark.parametrize("raw", ["12x1ctClam", "12 X 1 CT clam", "12x1ct clam", "12x1ct Clam"])
def test_compact_pack_styles(raw):
    assert normalize_pack_style(raw) == ("12x1ct Clam", None)


def test_descriptive_pack_style():
    value, warning = normalize_pack_style("1 case of 24ct green butterhead lettuce with roots")
    assert value == "1 case of 24ct Green Butterhead Lettuce with roots"
    assert warning is None


def test_multiple_descriptive_packs_are_separated():
    value, warning = normalize_pack_style("1 case of 24ct Green Lettuce 2 cases of 12ct Red Lettuce")
    assert value == "1 case of 24ct Green Lettuce / 2 cases of 12ct Red Lettuce"
    assert warning is None


def test_multiline_pdf_pack_style_is_flattened():
    value, warning = normalize_pack_style("1 case of 24ct Green Butterhead | Lettuce - with roots")
    assert value == "1 case of 24ct Green Butterhead Lettuce with roots"
    assert warning is None


def test_unknown_pack_style_is_flagged():
    value, warning = normalize_pack_style("mystery box")
    assert value == "Mystery Box"
    assert warning and "Please verify" in warning

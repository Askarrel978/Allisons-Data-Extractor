from pdf_extractor import detect_template, extract_fields, extract_fields_from_tables, extract_fields_regex

TEMPLATE_A_TEXT = """ATTENTION US
Shipping Date: 08/03/2026
Sales Order #: SO# 123456
Customer and Address: Acme Farms
1 Main St
Salinas, CA
Lot ID: L-77
Packaging Style: 12x1ct clam
Number of Plants: 240 heads (12 heads x 20 cases)
"""


def test_detect_template_by_marker():
    assert detect_template(TEMPLATE_A_TEXT) == "Template A"
    assert detect_template("LOT#: 55") == "Template B"
    assert detect_template("Green *Lettuce order") == "Template C"
    assert detect_template("nothing recognizable") == "Generic"


def test_fields_from_two_column_tables():
    tables = [[
        ["Shipping Date:", "08/03/2026"],
        ["Sales Order #:", "SO# 123456"],
        ["Customer and Address:", "Acme Farms\n1 Main St\nSalinas, CA"],
        ["Packaging Style:", "12x1ct clam"],
        ["Number of Plants:", "240 heads (12 heads x 20 cases)"],
        ["Lot ID:", "L-77"],
        ["Unrelated:", "ignored"],
    ]]
    assert extract_fields_from_tables(tables) == {
        "Date": "08/03/2026",
        "Sales Order Number": "123456",
        "Customer and Location": "Acme Farms, 1 Main St, Salinas, CA",
        "Pack Style": "12x1ct clam",
        "Number of Plants": "240 heads (12 heads x 20 cases)",
        "Lot Code": "L-77",
    }


def test_regex_fallback_on_plain_text():
    fields = extract_fields_regex(TEMPLATE_A_TEXT, "Template A")
    assert fields == {
        "Date": "08/03/2026",
        "Sales Order Number": "123456",
        "Customer and Location": "Acme Farms, 1 Main St, Salinas, CA",
        "Pack Style": "12x1ct clam",
        "Number of Plants": "240 heads (12 heads x 20 cases)",
        "Lot Code": "L-77",
    }


def test_table_values_take_priority_over_regex():
    tables = [[["Sales Order #:", "999999"]]]
    fields = extract_fields(TEMPLATE_A_TEXT, "Template A", tables)
    assert fields["Sales Order Number"] == "999999"
    assert fields["Lot Code"] == "L-77"

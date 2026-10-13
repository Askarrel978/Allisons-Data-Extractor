"""Draw Template B-style shipment tables for tests (made-up data, no real customers)."""

import os

from PIL import Image, ImageDraw, ImageFont

DARK = {"background": (37, 37, 37), "rule": (205, 205, 205), "text": (255, 255, 255)}
LIGHT = {"background": (255, 255, 255), "rule": (150, 150, 150), "text": (0, 0, 0)}

SHIPMENTS = [
    {
        "Date": "August 22, 2026",
        "Sales Order Number": "786674",
        "Customer and Location": "Harbor Greens Inc.\n430 N. Example Rd.\nGreensboro NC 27410\nUS",
        "Pack Style": "12x1ct clam",
        "Number of Plants": "840 heads (12 heads x 70 cases)",
        "Lot Code": "MUKFI232",
    },
    {
        "Date": "August 23, 2026",
        "Sales Order Number": "786024",
        "Customer and Location": "Sample Foods LLC\n2085 Test Rd\nSalisbury NC 28147\nUS",
        "Pack Style": "12x1ct clam / 8x2ct clam",
        "Number of Plants": "1520 heads (12 heads x 60 cases /\n16 heads x 50 cases)",
        "Lot Code": "MUKFI233",
    },
    {
        "Date": "August 24, 2026",
        "Sales Order Number": "787188",
        "Customer and Location": "Demo Markets Inc.\n3 Quarry Road\nReading PA 19605\nUS",
        "Pack Style": "8x2ct clam",
        "Number of Plants": "320 heads (16 heads x 20 cases)",
        "Lot Code": "MU01233",
    },
]

# What the extractor should return for each shipment above.
EXPECTED = [
    {
        "Date": "August 22, 2026",
        "Sales Order Number": "786674",
        "Customer and Location": "Harbor Greens Inc., 430 N. Example Rd., Greensboro NC 27410, US",
        "Pack Style": "12x1ct Clam",
        "Number of Plants": "840 heads (12 heads x 70 cases)",
        "Lot Code": "MUKFI232",
    },
    {
        "Date": "August 23, 2026",
        "Sales Order Number": "786024",
        "Customer and Location": "Sample Foods LLC, 2085 Test Rd, Salisbury NC 28147, US",
        "Pack Style": "12x1ct Clam / 8x2ct Clam",
        "Number of Plants": "1520 heads (12 heads x 60 cases / 16 heads x 50 cases)",
        "Lot Code": "MUKFI233",
    },
    {
        "Date": "August 24, 2026",
        "Sales Order Number": "787188",
        "Customer and Location": "Demo Markets Inc., 3 Quarry Road, Reading PA 19605, US",
        "Pack Style": "8x2ct Clam",
        "Number of Plants": "320 heads (16 heads x 20 cases)",
        "Lot Code": "MU01233",
    },
]


def _font(names, size):
    folder = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    for name in names:
        path = os.path.join(folder, name)
        if os.path.exists(path):
            return ImageFont.truetype(path, size)
    return ImageFont.load_default(size=size)


def draw_table(shipment, width=430, theme=DARK, label="Plant Species:", zoom=1):
    """One Template B table, row by row, like the screenshots Allison receives.

    ``zoom`` draws everything natively larger, like a high-resolution scan.
    """
    label_font = _font(["arialbi.ttf", "arial.ttf"], 15 * zoom)
    value_font = _font(["arialbd.ttf", "arial.ttf"], 13 * zoom)
    width *= zoom
    rows = [
        (label, "Lettuce (Lactuca sativa)"),
        ("Country of Origin:", "Canada"),
        ("Packaging style:", shipment["Pack Style"]),
        ("Number of Plants:", shipment["Number of Plants"]),
        ("Shipping Date:", shipment["Date"]),
        ("LOT#", shipment["Lot Code"]),
        ("Sales order #:", shipment["Sales Order Number"]),
        ("Customer and Address:", shipment["Customer and Location"]),
        ("Grown at:", "Example Farms Inc.\n838 Road 3 East, Kingsville, ON"),
        ("Exported By:", "Example Marketing Inc.,\nKingsville, ON, N9Y 2N1."),
    ]
    line_height = 17 * zoom
    pad = 3 * zoom
    rule = zoom
    heights = [max(len(value.split("\n")), 1) * line_height + 2 * pad for _label, value in rows]
    height = sum(heights) + (len(rows) + 1) * rule
    image = Image.new("RGB", (width, height), theme["background"])
    draw = ImageDraw.Draw(image)
    # Like real tables, the label column is always wide enough for its labels.
    widest_label = max(draw.textlength(label_text, font=label_font) for label_text, _value in rows)
    divider = max(int(width * 0.41), int(widest_label) + 14 * zoom)
    y = 0
    draw.line((0, y, width - 1, y), fill=theme["rule"], width=rule)
    for (label_text, value), row_height in zip(rows, heights):
        draw.text((2 * pad, y + rule + pad), label_text, fill=theme["text"], font=label_font)
        for index, line in enumerate(value.split("\n")):
            text_width = draw.textlength(line, font=value_font)
            x = divider + (width - divider - text_width) / 2
            draw.text((x, y + rule + pad + index * line_height), line, fill=theme["text"], font=value_font)
        y += row_height + rule
        draw.line((0, y, width - 1, y), fill=theme["rule"], width=rule)
    for x in (0, divider, width - rule):
        draw.line((x, 0, x, height - 1), fill=theme["rule"], width=rule)
    return image


def compose(tables, gap=30, offsets=None, horizontal=False, theme=DARK, margin=12):
    """Place several tables on one canvas, stacked or side by side."""
    offsets = offsets or [0] * len(tables)
    if horizontal:
        width = sum(t.width for t in tables) + gap * (len(tables) - 1) + 2 * margin
        height = max(t.height + o for t, o in zip(tables, offsets)) + 2 * margin
    else:
        width = max(t.width + o for t, o in zip(tables, offsets)) + 2 * margin
        height = sum(t.height for t in tables) + gap * (len(tables) - 1) + 2 * margin
    canvas = Image.new("RGB", (width, height), theme["background"])
    position = margin
    for table, offset in zip(tables, offsets):
        if horizontal:
            canvas.paste(table, (position, margin + offset))
            position += table.width + gap
        else:
            canvas.paste(table, (margin + offset, position))
            position += table.height + gap
    return canvas

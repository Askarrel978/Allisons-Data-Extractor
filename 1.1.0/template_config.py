"""Central template configuration for Allison's PDF Extractor."""

TEMPLATE_MARKERS = {
    "Template A": ["ATTENTION US"],
    "Template B": ["LOT#"],
    "Template C": ["*Lettuce"],
}

TEMPLATE_SETTINGS = {
    "Template A": {
        "source_type": "pdf",
        "lot_required": True,
        "sales_order_pattern": r"^\d{5,9}$",
    },
    "Template B": {
        "source_type": "image",
        "lot_required": True,
        "sales_order_pattern": r"^\d{5,9}$",
    },
    "Template C": {
        "source_type": "pdf",
        "lot_required": False,
        "sales_order_pattern": r"^\d{5,9}$",
    },
}

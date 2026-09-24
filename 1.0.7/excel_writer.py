"""
excel_writer.py

Writes extracted rows into an Excel file with the fixed header set:
Date, Sales Order Number, Customer and Location, Pack Style,
Number of Plants, Lot Code

If the target file already exists, new rows are appended after the last
row of data (existing content is never overwritten). If it doesn't exist,
a new workbook is created with a bolded header row.
"""

import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from pdf_extractor import HEADERS


def _new_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(HEADERS)
    for col_idx in range(1, len(HEADERS) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)
    return wb, ws


def _autosize_columns(ws):
    for col_idx, header in enumerate(HEADERS, start=1):
        max_len = len(str(header))
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def append_rows(filepath: str, rows: list):
    """
    rows: list of dicts, each with keys matching HEADERS (missing keys become "")
    Creates the file if it doesn't exist, otherwise appends to the first sheet.
    """
    if os.path.exists(filepath):
        wb = load_workbook(filepath)
        ws = wb.active
        # If the sheet is totally empty, add headers first.
        if ws.max_row == 1 and all(c.value is None for c in ws[1]):
            ws.append(HEADERS)
            for col_idx in range(1, len(HEADERS) + 1):
                ws.cell(row=1, column=col_idx).font = Font(bold=True)
    else:
        wb, ws = _new_workbook()

    for row in rows:
        ws.append([row.get(h, "") for h in HEADERS])

    _autosize_columns(ws)
    wb.save(filepath)

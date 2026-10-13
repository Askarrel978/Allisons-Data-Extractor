# Allison's Data Extractor

A Windows desktop application for logging plant shipments. It reads shipment documents (PDFs, scanned PDFs and screenshots of emailed shipment tables), pulls out the fields that matter, and appends them to an existing Excel workbook.

Without it, each shipment is typed into the spreadsheet by hand from the source document. With it, a batch of documents is added in one step, every value is checked, and the results are shown in a table for review before anything is written to the workbook.

## Fields extracted

Each shipment produces one row with:

- Date
- Sales Order Number
- Customer and Location
- Pack Style
- Number of Plants
- Lot Code

## Supported documents

| Source | Format | Method |
| --- | --- | --- |
| Template A | PDF | Text and table extraction |
| Template B | Image (PNG, JPG, TIFF) or PDF printed from email | Tesseract OCR, cell by cell. One file can contain several shipment tables. |
| Template C | PDF | Text and table extraction |
| Scanned PDFs | PDF with no selectable text | Tesseract OCR of each page, then the normal template rules |

The template is detected automatically from the document content.

## How it works

1. Add files to the queue and run the extraction.
2. Each record is validated and given a confidence score. The checks cover:
   - Sales Order format
   - Dates
   - Required fields
   - Lot codes
   - Plant-count arithmetic (for example 12 x 20 = 240)
   - Duplicates within the batch or already in the workbook

   Anything read by OCR is always marked for review.
3. Records are reviewed and corrected in an editable table. The original document can be opened from a selected record, and the record can be reprocessed or approved there.
4. The reviewed records are exported to the selected workbook.

## Protecting the workbook

The workbook is the record of truth, so exporting is deliberately cautious:

- A timestamped backup is made before anything is written.
- Values are matched to columns by header name. Workbooks with a different column order or extra columns still line up correctly.
- The file is written to a temporary copy and then swapped into place, so an interrupted export can't corrupt the original.
- Export is refused while the workbook is open in Excel.
- In Production Safety Mode, export is blocked if required fields are missing or low-confidence records haven't been approved.

Every extraction, edit, approval and export is recorded in a local SQLite history. The history can be viewed in the app and exported as CSV.

## Administration

Template rules, validation patterns and Pack Style aliases are set in Settings. None of these changes needs a rebuild. Administrator settings are protected by a passcode, stored only as a salted PBKDF2 hash.

Settings also includes:

- A diagnostics page
- Configuration import and export
- Installation of updates from a USB package. Packages are checked against the current version and SHA-256 verified before they're installed.

## Repository layout

Each top-level folder is a snapshot of the source code for one release:

| Folder | Highlights |
| --- | --- |
| [`1.3.0`](1.3.0) | **Current release.** Reads any number of Template B tables per image or PDF. Scanned PDFs are supported. Extraction runs in the background. Excel export is safer. |
| [`1.2.1`](1.2.1) | Expanded settings, Production Safety Mode, diagnostics and audit options |
| [`1.2.0`](1.2.0) | Renamed to Allison's Data Extractor. Added the administrator settings, configurable templates and USB updates. |
| [`1.1.0`](1.1.0) | Confidence scoring, validation, duplicate detection, backups and processing history |
| [`1.0.7`](1.0.7) | Templates A, B and C with automatic detection, and Template B OCR |
| [`1.0.6`](1.0.6) | First version (Allison's PDF Extractor) |

The full release notes are in [`1.3.0/CHANGELOG.txt`](1.3.0/CHANGELOG.txt).

## Running and building

The 1.3.0 instructions are in [`1.3.0/README.md`](1.3.0/README.md). In short, from the `1.3.0` folder:

```
uv sync --group dev
uv run main.py      # run the app
uv run pytest -q    # run the tests
build.bat           # build the portable Windows app and an update package
```

Requirements:

- Windows
- Python 3.14
- [uv](https://docs.astral.sh/uv/)
- Tesseract OCR, for images and scanned PDFs

The bundled `Tesseract-OCR` folder is too large for GitHub, so it isn't in this repository. Copy it into the build folder before running `build.bat`.

Built with Tkinter, pdfplumber, pytesseract, Pillow and openpyxl.

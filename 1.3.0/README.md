# Allison's Data Extractor

A Windows desktop app that reads shipment documents and adds their details to an Excel workbook.

For each shipment it extracts:

- Date
- Sales Order Number
- Customer and Location
- Pack Style
- Number of Plants
- Lot Code

Results appear in an editable review table, so they can be checked and corrected before export.

| Source | Format | How it is read |
| --- | --- | --- |
| Template A | PDF | PDF text and tables |
| Template B | PNG/JPG/... image | Tesseract OCR, cell by cell; one image can hold several shipments |
| Template C | PDF | PDF text and tables |
| Scanned PDFs | PDF without selectable text | Tesseract OCR of each page, then the normal text rules |

## Running from source

Requires [uv](https://docs.astral.sh/uv/) and Python 3.14.

```
uv sync --group dev
uv run main.py
```

OCR (Template B and scanned PDFs) needs Tesseract. The app finds it in these places, in order:

1. The path set in Settings > OCR
2. The `TESSERACT_CMD` environment variable
3. A `Tesseract-OCR` folder beside the app
4. A normal Windows install

## Tests

```
uv run pytest -q
```

The tests use a temporary data folder and never touch the real settings or history. The scanned-PDF test is skipped if Tesseract isn't available.

## Building the portable app

The `Tesseract-OCR/` folder is too large for GitHub and isn't in the repository. Before building, copy it into this folder from a previous build or a Tesseract install. It must contain `tesseract.exe` and `tessdata/`.

Then run `build.bat`. The steps are:

1. Build the app into `dist\Allisons Data Extractor\`.
2. Build the updater helper.
3. Create a USB update package in `release\`.

## Where data is stored

Settings, configuration backups and the processing history (`history.db`) are kept per Windows user in `%LOCALAPPDATA%\AllisonsDataExtractor`. This means replacing or updating the app doesn't lose them.

Before each export into an existing workbook, a timestamped backup is saved to an `Allisons Data Extractor Backups` folder next to it.

## Code layout

| File | Purpose |
| --- | --- |
| `main.py` | Tkinter GUI: file queue, review table, export flow |
| `pdf_extractor.py` | Template detection, field extraction, OCR, normalization, validation |
| `excel_writer.py` | Header-matched appending, duplicate check, backups, safe saving |
| `settings_manager.py` | Settings file, defaults, admin passcode |
| `settings_dialog.py` | Settings window |
| `history_db.py` | SQLite processing history |
| `update_manager.py`, `updater.py`, `package_update.py` | USB update packages |

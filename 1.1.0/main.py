"""
main.py - PDF -> Excel Plant Data Extractor

A portable desktop GUI that:
  1. Lets you pick individual PDF/image files or a whole folder.
  2. Extracts Date, Sales Order Number, Customer and Location, Pack Style,
     Number of Plants, and Lot Code from each one.
  3. Shows the results in an editable preview table so you can fix anything
     before it's saved.
  4. Appends the confirmed rows into your chosen Excel file.

Run with:  python main.py
Build to .exe with: build_exe.bat  (see README.md)
"""

import os
import re
import sys
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pdf_extractor import process_pdf, HEADERS, validate_record
from excel_writer import append_rows, find_duplicates
from history_db import log_event, recent_events, clear_history

# Internal provenance marker. Not displayed in the GUI.
# This is a lightweight authorship breadcrumb, not copy protection.
__author__ = "Alexander Hostine"
__copyright__ = "Copyright (c) 2026 Alexander Hostine"
__provenance_id__ = "AH-ALLISON-EXTRACTOR-2026-8F2C7A19"

DISPLAY_COLUMNS = HEADERS + ["Confidence", "Status", "Template", "Warnings", "Source File"]

# ---------------------------------------------------------------------------
# First Pop Up Warning Window.
# ---------------------------------------------------------------------------
APP_AUTHOR = "Alexander Hostine"
APP_REVISION = "1.1.0"
APP_DESCRIPTION = (
    "This tool reads shipment files (Templates A, B, and C) and "
    "extracts the following fields: Date, Sales Order Number, Customer and "
    "Location, Pack Style, Number of Plants, and Lot Code. Extracted data "
    "is reviewed in an editable table and then saved into an Excel "
    "spreadsheet of your choosing."
)
APP_COPYRIGHT = (
    "Made with love for Allison by Alexander Hostine (2026)."
)

CHANGELOG_TEXT = """Revision 1.1.0

Major additions
• Confidence scoring and validation for every extracted shipment.
• Plant-count arithmetic validation, for example 12 x 20 = 240.
• Sales-order, date, required-field and lot-code validation.
• Duplicate detection against the selected Excel workbook and within the current batch.
• Automatic timestamped Excel backups before appending to an existing workbook.
• SQLite processing history for extraction, edits, approvals and exports.
• History window in the GUI.
• Selected-record detail and review panel.
• Open Source, Reprocess and Approve actions.
• Template configuration moved into template_config.py for easier future templates.
• Template B supports multiple tables in one image, with one output row per table.
• Template B uses field/cell-level OCR for improved accuracy.
• Added a Help window to the main GUI.
• Added a Changelog button to the startup window.
• Added clickable column sorting, including chronological Date sorting.
• Confidence scoring now compares extracted formatting against Template A conventions.
• Template B OCR records receive an automatic verification warning and confidence reduction.
• Added Clear All to Processing History with confirmation.

Revision 1.0.7

• Added support for Templates A, B, and C.
• Added automatic template detection.
• Added Template B image OCR using Tesseract.
• Added portable Tesseract-OCR support for compiled builds.
• Added multi-table Template B extraction with one output row per table.
• Improved Template B Sales Order and field recognition.
• Preserved complete Number of Plants values, including units and case breakdowns.
• Added PDF and image file selection support.
• Improved GUI column sizing and portable PyInstaller compatibility.
"""

HELP_TEXT = """Allison's PDF Extractor - Help

ADDING FILES
Use Add Files to select PDF or image shipment files, or Add Folder to load all supported files in a folder.

SUPPORTED INPUTS
Template A: PDF
Template B: image files such as PNG/JPG. A single image may contain multiple shipment tables; each detected table becomes its own row.
Template C: PDF

EXTRACTING DATA
After adding files, click Extract Data. Review the Confidence, Status and Warnings columns before exporting. Click any column heading to sort; click it again to reverse the order.

EDITING
Double-click any extracted data cell to correct it. The record is revalidated automatically after an edit.

SELECTED RECORD
Open Source opens the original file. Reprocess runs extraction again. Approve marks the selected row as manually approved.

HISTORY
The History button shows recent extraction, editing, approval and export events. Use Clear All in that window to permanently erase the local history after confirmation.

EXPORTING
Choose an existing Excel file or create a new one. The program checks for possible duplicate sales orders before export and creates a timestamped backup before modifying an existing workbook.

TEMPLATE B / OCR
Template B requires Tesseract OCR. Portable builds should include the Tesseract-OCR folder beside the application. If OCR produces questionable data, review any warning or low-confidence record before exporting.

TROUBLESHOOTING
If a field is missing, open the source and compare it with the extracted row. Correct the field manually or use Reprocess. For persistent template problems, keep the source file/image as a sample for updating the extractor.
"""

EXPORT_WARNING_TEXT = (
    "Before exporting, please note:\n\n"
    "\u2022 Excel files can become corrupted if this program is closed, the "
    "computer loses power, or the file is edited elsewhere while exporting.\n\n"
    "\u2022 Templates A, B, and C are not identical, and the field layout in your "
    "spreadsheet may differ from what was extracted \u2014 please check the "
    "preview table carefully for accuracy and consistency before proceeding.\n\n"
    "\u2022 Always keep a separate backup of your Excel file before exporting, "
    "in case the data needs to be recovered or the export needs to be undone."
)


class TextInfoDialog(tk.Toplevel):
    """Simple scrollable informational dialog used for Help and Changelog."""

    def __init__(self, parent, title, text, width=680, height=520):
        super().__init__(parent)
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.minsize(520, 360)
        self.transient(parent)

        container = ttk.Frame(self, padding=10)
        container.pack(fill="both", expand=True)

        text_frame = ttk.Frame(container)
        text_frame.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")

        viewer = tk.Text(
            text_frame,
            wrap="word",
            yscrollcommand=scrollbar.set,
            padx=10,
            pady=10,
        )
        viewer.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=viewer.yview)
        viewer.insert("1.0", text)
        viewer.configure(state="disabled")

        ttk.Button(container, text="Close", command=self.destroy).pack(anchor="e", pady=(8, 0))


class StartupDialog(tk.Toplevel):
    """Modal dialog shown on launch. Blocks the app until Accept/Decline."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Welcome")
        self.resizable(False, False)
        self.accepted = False

        # Prevent interaction with the (hidden) main window until this closes.
        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_decline)

        container = ttk.Frame(self, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="Allison's PDF Extractor",
            font=("TkDefaultFont", 13, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            container,
            text=f"Developed by {APP_AUTHOR}    |    Revision {APP_REVISION}",
            font=("TkDefaultFont", 9, "bold"),
        ).pack(anchor="w", pady=(2, 12))

        ttk.Label(
            container,
            text=APP_DESCRIPTION,
            wraplength=420,
            justify="left",
        ).pack(anchor="w")

        ttk.Label(
             container,
            text=APP_COPYRIGHT,
            font=("TkDefaultFont", 9, "italic"),
            wraplength=420,
            justify="left",
        ).pack(anchor="w")

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(20, 0))

        ttk.Button(button_row, text="Changelog", command=self._show_changelog).pack(side="left")
        ttk.Button(button_row, text="Decline", command=self._on_decline).pack(side="right", padx=(6, 0))
        ttk.Button(button_row, text="Accept", command=self._on_accept).pack(side="right")

        self.update_idletasks()
        self._center_on_parent(parent)

    def _center_on_parent(self, parent):
        parent.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        x = px + (pw - w) // 2 if pw > 1 else px + 200
        y = py + (ph - h) // 2 if ph > 1 else py + 150
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _show_changelog(self):
        TextInfoDialog(self, f"Changelog - Revision {APP_REVISION}", CHANGELOG_TEXT)

    def _on_accept(self):
        self.accepted = True
        self.destroy()

    def _on_decline(self):
        self.accepted = False
        self.destroy()


class ExportWarningDialog(tk.Toplevel):
    """Modal warning shown before writing to Excel. Blocks until a choice is made."""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("Before You Export")
        self.resizable(False, False)
        self.confirmed = False

        self.transient(parent)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._on_go_back)

        container = ttk.Frame(self, padding=20)
        container.pack(fill="both", expand=True)

        ttk.Label(
            container,
            text="\u26A0 Data Consistency Warning",
            font=("TkDefaultFont", 12, "bold"),
        ).pack(anchor="w")

        ttk.Label(
            container,
            text=EXPORT_WARNING_TEXT,
            wraplength=440,
            justify="left",
        ).pack(anchor="w", pady=(10, 0))

        button_row = ttk.Frame(container)
        button_row.pack(fill="x", pady=(20, 0))

        ttk.Button(button_row, text="Go Back", command=self._on_go_back).pack(side="right", padx=(6, 0))
        ttk.Button(button_row, text="I Understand", command=self._on_understand).pack(side="right")

        self.update_idletasks()
        self._center_on_parent(parent)

    def _center_on_parent(self, parent):
        parent.update_idletasks()
        px, py = parent.winfo_x(), parent.winfo_y()
        pw, ph = parent.winfo_width(), parent.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        x = px + (pw - w) // 2 if pw > 1 else px + 200
        y = py + (ph - h) // 2 if ph > 1 else py + 150
        self.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _on_understand(self):
        self.confirmed = True
        self.destroy()

    def _on_go_back(self):
        self.confirmed = False
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Allison's PDF Extractor")
        self.geometry("1380x760")
        self.minsize(1050, 620)

        self.pdf_paths = []
        self.rows_data = []  # list of dicts: HEADERS + Template/Warnings/Source File
        self.output_path = tk.StringVar(value="")
        self._sort_reverse = {}

        # Show the startup dialog on top of the main window (do not hide the
        # main window first - on some Windows/Tk builds, a dialog that is
        # "transient" to a withdrawn window stays invisible too).
        self._build_top_bar()
        self._build_table()
        self._build_detail_panel()
        self._build_bottom_bar()
        self._build_status_bar()

        self.after(0, self._show_startup_dialog)

    def _show_startup_dialog(self):
        dialog = StartupDialog(self)
        self.wait_window(dialog)
        if not dialog.accepted:
            self.destroy()
            sys.exit(0)

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------
    def _build_top_bar(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")

        ttk.Button(bar, text="Add Files...", command=self.add_files).pack(side="left", padx=4)
        ttk.Button(bar, text="Add Folder...", command=self.add_folder).pack(side="left", padx=4)
        ttk.Button(bar, text="Clear List", command=self.clear_files).pack(side="left", padx=4)
        ttk.Button(bar, text="History", command=self.show_history).pack(side="left", padx=4)
        ttk.Button(bar, text="Help", command=self.show_help).pack(side="left", padx=4)

        self.file_count_label = ttk.Label(bar, text="0 file(s) queued")
        self.file_count_label.pack(side="left", padx=16)

        ttk.Button(bar, text="Extract Data", command=self.run_extraction).pack(side="right", padx=4)

    def _build_table(self):
        frame = ttk.Frame(self, padding=(8, 0))
        frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(frame, columns=DISPLAY_COLUMNS, show="headings", selectmode="browse")
        column_widths = {
            "Date": 140,
            "Sales Order Number": 150,
            "Customer and Location": 430,
            "Pack Style": 300,
            "Number of Plants": 280,
            "Lot Code": 140,
            "Confidence": 95,
            "Status": 95,
            "Template": 110,
            "Warnings": 320,
            "Source File": 200,
        }
        for col in DISPLAY_COLUMNS:
            self.tree.heading(col, text=col, command=lambda c=col: self.sort_by_column(c))
            self.tree.column(col, width=column_widths.get(col, 120), anchor="w", stretch=True)

        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self.on_cell_double_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)

        hint = ttk.Label(
            self,
            text="Tip: click a column heading to sort; click again to reverse. Double-click a cell to edit.",
            foreground="#555555",
        )
        hint.pack(anchor="w", padx=10, pady=(2, 0))

    def sort_by_column(self, column):
        """Sort the backing rows and redraw the Treeview, like File Explorer headings."""
        reverse = self._sort_reverse.get(column, False)

        def key_for(row):
            value = row.get(column, "")
            text = str(value or "").strip()

            if column == "Date":
                try:
                    return (0, datetime.strptime(text, "%B %d, %Y"))
                except ValueError:
                    return (1, datetime.min)

            if column == "Confidence":
                try:
                    return (0, float(text.rstrip("%")))
                except ValueError:
                    return (1, -1.0)

            if column == "Sales Order Number":
                digits = re.sub(r"\D", "", text)
                if digits:
                    return (0, int(digits))
                return (1, 0)

            return text.casefold()

        # Python cannot compare unlike tuple/string key types, so use a wrapper
        # that keeps text columns separate from the numeric/date special cases.
        if column in ("Date", "Confidence", "Sales Order Number"):
            self.rows_data.sort(key=key_for, reverse=reverse)
        else:
            self.rows_data.sort(key=lambda row: str(row.get(column, "") or "").casefold(), reverse=reverse)

        self._sort_reverse[column] = not reverse
        self._populate_table()

        # Show the direction in the active heading and restore the others.
        for col in DISPLAY_COLUMNS:
            label = col
            if col == column:
                label += " ▼" if reverse else " ▲"
            self.tree.heading(col, text=label, command=lambda c=col: self.sort_by_column(c))

    def _build_detail_panel(self):
        panel = ttk.LabelFrame(self, text="Selected Record", padding=8)
        panel.pack(fill="x", padx=8, pady=(6, 0))

        self.detail_text = tk.Text(panel, height=5, wrap="word")
        self.detail_text.pack(side="left", fill="both", expand=True)
        self.detail_text.configure(state="disabled")

        buttons = ttk.Frame(panel)
        buttons.pack(side="right", fill="y", padx=(8, 0))
        ttk.Button(buttons, text="Open Source", command=self.open_selected_source).pack(fill="x", pady=2)
        ttk.Button(buttons, text="Reprocess", command=self.reprocess_selected).pack(fill="x", pady=2)
        ttk.Button(buttons, text="Approve", command=self.approve_selected).pack(fill="x", pady=2)

    def _on_row_selected(self, event=None):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        if idx >= len(self.rows_data):
            return
        row = self.rows_data[idx]
        lines = [f"{h}: {row.get(h, '')}" for h in HEADERS]
        lines += [
            f"Confidence: {row.get('Confidence', '')}",
            f"Status: {row.get('Status', '')}",
            f"Template: {row.get('Template', '')}",
            f"Warnings: {row.get('Warnings', '') or 'None'}",
        ]
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", "\n".join(lines))
        self.detail_text.configure(state="disabled")

    def _build_bottom_bar(self):
        bar = ttk.Frame(self, padding=8)
        bar.pack(fill="x")

        ttk.Label(bar, text="Output Excel file:").pack(side="left")
        entry = ttk.Entry(bar, textvariable=self.output_path, width=70)
        entry.pack(side="left", padx=6)
        ttk.Button(bar, text="Browse Existing...", command=self.browse_existing_excel).pack(side="left", padx=4)
        ttk.Button(bar, text="Save As New...", command=self.browse_new_excel).pack(side="left", padx=4)

        ttk.Button(bar, text="Export to Excel", command=self.export_to_excel).pack(side="right", padx=4)
        ttk.Button(bar, text="Remove Selected Row", command=self.remove_selected_row).pack(side="right", padx=4)

    def _build_status_bar(self):
        self.status_var = tk.StringVar(value="Ready.")
        status = ttk.Label(self, textvariable=self.status_var, relief="sunken", anchor="w", padding=4)
        status.pack(fill="x", side="bottom")

    # ------------------------------------------------------------------
    # File selection
    # ------------------------------------------------------------------
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select shipment files",
            filetypes=[
                ("Supported files", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("All files", "*.*"),
            ],
        )
        if paths:
            for p in paths:
                if p not in self.pdf_paths:
                    self.pdf_paths.append(p)
            self._refresh_file_count()

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select a folder of shipment files")
        if folder:
            supported_extensions = (".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp")
            found = [
                os.path.join(folder, f)
                for f in sorted(os.listdir(folder))
                if f.lower().endswith(supported_extensions)
            ]
            for p in found:
                if p not in self.pdf_paths:
                    self.pdf_paths.append(p)
            self._refresh_file_count()
            if not found:
                messagebox.showinfo(
                    "No supported files found",
                    "That folder didn't contain any supported PDF or image files.",
                )

    def clear_files(self):
        self.pdf_paths = []
        self._refresh_file_count()

    def _refresh_file_count(self):
        self.file_count_label.config(text=f"{len(self.pdf_paths)} file(s) queued")

    # ------------------------------------------------------------------
    # Extraction
    # ------------------------------------------------------------------
    def run_extraction(self):
        if not self.pdf_paths:
            messagebox.showwarning("No files", "Add some PDF/image files or a folder first.")
            return
        threading.Thread(target=self._extract_worker, daemon=True).start()

    def _extract_worker(self):
        self._set_status(f"Processing 0/{len(self.pdf_paths)}...")
        new_rows = []
        for i, path in enumerate(self.pdf_paths, start=1):
            result = process_pdf(path)

            if result["error"]:
                # Keep one visible row for a failed file so the user can review the error.
                row = {header: "" for header in HEADERS}
                row["Template"] = result.get("template", "")
                row["Source File"] = os.path.basename(path)
                row["_Source Path"] = path
                row["Confidence"] = "0%"
                row["Status"] = "Attention"
                row["Warnings"] = result["error"]
                new_rows.append(row)
            else:
                records = result.get("records") or [result.get("fields", {})]
                warnings_by_record = result.get("record_warnings") or []

                confidences = result.get("record_confidence") or []
                statuses = result.get("record_status") or []
                for record_index, fields in enumerate(records):
                    row = dict(fields)
                    row["Template"] = result["template"]
                    row["Source File"] = os.path.basename(path)
                    row["_Source Path"] = path

                    record_warnings = (
                        warnings_by_record[record_index]
                        if record_index < len(warnings_by_record)
                        else []
                    )
                    confidence = confidences[record_index] if record_index < len(confidences) else None
                    status = statuses[record_index] if record_index < len(statuses) else None
                    if confidence is None or status is None:
                        confidence, status, extra_warnings = validate_record(fields, result["template"])
                        record_warnings = list(dict.fromkeys(record_warnings + extra_warnings))
                    row["Confidence"] = f"{confidence}%"
                    row["Status"] = status
                    row["Warnings"] = "; ".join(record_warnings)
                    new_rows.append(row)
                    log_event(
                        "Extracted", os.path.basename(path), result["template"],
                        fields.get("Sales Order Number", ""), status,
                        details={"confidence": confidence, "warnings": record_warnings},
                    )

            self._set_status(f"Processing {i}/{len(self.pdf_paths)}...")

        self.rows_data = new_rows
        self.after(0, self._populate_table)
        self._set_status(f"Done. Extracted {len(new_rows)} row(s). Review and edit below, then export.")

    def _populate_table(self):
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.rows_data):
            values = [row.get(col, "") for col in DISPLAY_COLUMNS]
            self.tree.insert("", "end", iid=str(i), values=values)
        if self.rows_data:
            self.tree.selection_set("0")
            self._on_row_selected()

    def _set_status(self, text):
        self.status_var.set(text)

    # ------------------------------------------------------------------
    # Inline editing
    # ------------------------------------------------------------------
    def on_cell_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        col_name = DISPLAY_COLUMNS[col_index]
        if col_name not in HEADERS:
            return
        current_value = self.tree.set(row_id, col_name)

        x, y, width, height = self.tree.bbox(row_id, col_id)
        edit_var = tk.StringVar(value=current_value)
        editor = ttk.Entry(self.tree, textvariable=edit_var)
        editor.place(x=x, y=y, width=width, height=height)
        editor.focus()
        editor.select_range(0, "end")

        def commit(event=None):
            new_value = edit_var.get()
            self.tree.set(row_id, col_name, new_value)
            self.rows_data[int(row_id)][col_name] = new_value
            if col_name in HEADERS:
                self._revalidate_row(int(row_id), edited_field=col_name)
            log_event("Edited", self.rows_data[int(row_id)].get("Source File", ""), self.rows_data[int(row_id)].get("Template", ""), self.rows_data[int(row_id)].get("Sales Order Number", ""), self.rows_data[int(row_id)].get("Status", ""), details={"field": col_name})
            editor.destroy()
            self._on_row_selected()

        def cancel(event=None):
            editor.destroy()

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", cancel)

    def _revalidate_row(self, idx, edited_field=""):
        row = self.rows_data[idx]
        confidence, status, warnings = validate_record(row, row.get("Template", ""))
        row["Confidence"] = f"{confidence}%"
        if row.get("Status") != "Approved":
            row["Status"] = status
        row["Warnings"] = "; ".join(warnings)
        if self.tree.exists(str(idx)):
            for col in DISPLAY_COLUMNS:
                self.tree.set(str(idx), col, row.get(col, ""))

    def open_selected_source(self):
        selected = self.tree.selection()
        if not selected:
            return
        row = self.rows_data[int(selected[0])]
        path = row.get("_Source Path", "")
        if not path or not os.path.exists(path):
            messagebox.showwarning("Source unavailable", "The original source file could not be found.")
            return
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Open failed", str(e))

    def approve_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        row = self.rows_data[idx]
        row["Status"] = "Approved"
        self.tree.set(str(idx), "Status", "Approved")
        log_event("Approved", row.get("Source File", ""), row.get("Template", ""), row.get("Sales Order Number", ""), "Approved")
        self._on_row_selected()

    def reprocess_selected(self):
        selected = self.tree.selection()
        if not selected:
            return
        idx = int(selected[0])
        path = self.rows_data[idx].get("_Source Path", "")
        if not path or not os.path.exists(path):
            messagebox.showwarning("Source unavailable", "The original source file could not be found.")
            return
        result = process_pdf(path)
        if result.get("error"):
            messagebox.showerror("Reprocess failed", result["error"])
            return
        replacements = []
        for i, fields in enumerate(result.get("records") or [result.get("fields", {})]):
            row = dict(fields)
            row["Template"] = result.get("template", "")
            row["Source File"] = os.path.basename(path)
            row["_Source Path"] = path
            warnings = (result.get("record_warnings") or [[]])[i] if i < len(result.get("record_warnings") or []) else []
            confidence = (result.get("record_confidence") or [None])[i] if i < len(result.get("record_confidence") or []) else None
            status = (result.get("record_status") or [None])[i] if i < len(result.get("record_status") or []) else None
            if confidence is None:
                confidence, status, extra = validate_record(fields, row["Template"])
                warnings = list(dict.fromkeys(warnings + extra))
            row["Confidence"] = f"{confidence}%"
            row["Status"] = status
            row["Warnings"] = "; ".join(warnings)
            replacements.append(row)
        self.rows_data[idx:idx + 1] = replacements
        self._populate_table()
        self._set_status(f"Reprocessed {os.path.basename(path)} into {len(replacements)} row(s).")

    def show_help(self):
        TextInfoDialog(self, f"Help - Allison's PDF Extractor {APP_REVISION}", HELP_TEXT, width=760, height=600)

    def show_history(self):
        win = tk.Toplevel(self)
        win.title("Processing History")
        win.geometry("1050x460")

        toolbar = ttk.Frame(win, padding=(8, 8, 8, 0))
        toolbar.pack(fill="x")
        count_var = tk.StringVar(value="")
        ttk.Label(toolbar, textvariable=count_var).pack(side="left")

        cols = ("Time", "Action", "Source", "Template", "Sales Order", "Status", "Output")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=130 if col not in ("Source", "Output") else 220, anchor="w")
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        def refresh():
            tree.delete(*tree.get_children())
            records = recent_events(250)
            for record in records:
                tree.insert("", "end", values=record[:7])
            count_var.set(f"Showing {len(records)} most recent history event(s)")

        def clear_all_history():
            confirmed = messagebox.askyesno(
                "Clear processing history",
                "Permanently delete all local processing history?\n\nThis cannot be undone.",
                parent=win,
            )
            if not confirmed:
                return
            try:
                removed = clear_history()
            except Exception as exc:
                messagebox.showerror("History error", str(exc), parent=win)
                return
            refresh()
            messagebox.showinfo("History cleared", f"Removed {removed} history event(s).", parent=win)

        ttk.Button(toolbar, text="Clear All", command=clear_all_history).pack(side="right")
        ttk.Button(toolbar, text="Refresh", command=refresh).pack(side="right", padx=(0, 6))
        refresh()

    def remove_selected_row(self):
        selected = self.tree.selection()
        if not selected:
            return
        row_id = selected[0]
        idx = int(row_id)
        del self.rows_data[idx]
        self._populate_table()

    # ------------------------------------------------------------------
    # Excel output
    # ------------------------------------------------------------------
    def browse_existing_excel(self):
        path = filedialog.askopenfilename(title="Choose existing Excel file", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.output_path.set(path)

    def browse_new_excel(self):
        path = filedialog.asksaveasfilename(
            title="Save new Excel file as",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def export_to_excel(self):
        if not self.rows_data:
            messagebox.showwarning("Nothing to export", "Extract some data first.")
            return
        out_path = self.output_path.get().strip()
        if not out_path:
            messagebox.showwarning("No output file", "Choose or create an output Excel file first.")
            return

        warning = ExportWarningDialog(self)
        self.wait_window(warning)
        if not warning.confirmed:
            self._set_status("Export cancelled.")
            return

        duplicates = find_duplicates(out_path, self.rows_data)
        if duplicates:
            preview = "\n".join(
                f"SO {d['sales_order']} already exists on Excel row {d['excel_row']} ({d['match_type']})"
                for d in duplicates[:8]
            )
            if len(duplicates) > 8:
                preview += f"\n...and {len(duplicates) - 8} more."
            proceed = messagebox.askyesno(
                "Possible duplicates",
                "Possible duplicate shipment(s) were found:\n\n" + preview +
                "\n\nExport them anyway?",
            )
            if not proceed:
                self._set_status("Export cancelled because of duplicate review.")
                return

        try:
            backup_path = append_rows(out_path, self.rows_data, create_backup_first=True)
        except PermissionError:
            messagebox.showerror(
                "File in use",
                "Couldn't save - close the Excel file if it's currently open, then try again.",
            )
            return
        except Exception as e:
            messagebox.showerror("Export failed", str(e))
            return

        for row in self.rows_data:
            log_event("Exported", row.get("Source File", ""), row.get("Template", ""), row.get("Sales Order Number", ""), row.get("Status", ""), out_path, {"backup": backup_path})
        backup_note = f"\n\nBackup created at:\n{backup_path}" if backup_path else ""
        messagebox.showinfo("Done", f"Exported {len(self.rows_data)} row(s) to:\n{out_path}{backup_note}")
        self._set_status(f"Exported {len(self.rows_data)} row(s) to {out_path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()

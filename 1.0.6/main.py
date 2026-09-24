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
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pdf_extractor import process_pdf, HEADERS
from excel_writer import append_rows

DISPLAY_COLUMNS = HEADERS + ["Template", "Warnings", "Source File"]

# ---------------------------------------------------------------------------
# First Pop Up Warning Window.
# ---------------------------------------------------------------------------
APP_AUTHOR = "Alexander Hostine"
APP_REVISION = "1.0.6"
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
        self.geometry("1150x600")
        self.minsize(900, 500)

        self.pdf_paths = []
        self.rows_data = []  # list of dicts: HEADERS + Template/Warnings/Source File
        self.output_path = tk.StringVar(value="")

        # Show the startup dialog on top of the main window (do not hide the
        # main window first - on some Windows/Tk builds, a dialog that is
        # "transient" to a withdrawn window stays invisible too).
        self._build_top_bar()
        self._build_table()
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
            "Template": 110,
            "Warnings": 320,
            "Source File": 200,
        }
        for col in DISPLAY_COLUMNS:
            self.tree.heading(col, text=col)
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

        hint = ttk.Label(
            self,
            text="Tip: double-click any cell to edit it before exporting.",
            foreground="#555555",
        )
        hint.pack(anchor="w", padx=10, pady=(2, 0))

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
            row = dict(result["fields"])
            row["Template"] = result["template"]
            row["Source File"] = os.path.basename(path)
            if result["error"]:
                row["Warnings"] = result["error"]
            else:
                row["Warnings"] = "; ".join(result["warnings"])
            new_rows.append(row)
            self._set_status(f"Processing {i}/{len(self.pdf_paths)}...")

        self.rows_data = new_rows
        self.after(0, self._populate_table)
        self._set_status(f"Done. Extracted {len(new_rows)} row(s). Review and edit below, then export.")

    def _populate_table(self):
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(self.rows_data):
            values = [row.get(col, "") for col in DISPLAY_COLUMNS]
            self.tree.insert("", "end", iid=str(i), values=values)

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
            editor.destroy()

        def cancel(event=None):
            editor.destroy()

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", cancel)

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

        try:
            append_rows(out_path, self.rows_data)
        except PermissionError:
            messagebox.showerror(
                "File in use",
                "Couldn't save - close the Excel file if it's currently open, then try again.",
            )
            return
        except Exception as e:
            messagebox.showerror("Export failed", str(e))
            return

        messagebox.showinfo("Done", f"Exported {len(self.rows_data)} row(s) to:\n{out_path}")
        self._set_status(f"Exported {len(self.rows_data)} row(s) to {out_path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()

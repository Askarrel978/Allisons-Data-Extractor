"""Settings UI for Allison's Data Extractor."""

from __future__ import annotations

import copy
import csv
import json
import os
import platform
import sys
import tempfile
import tkinter as tk
import zipfile
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

from settings_manager import (
    admin_is_configured,
    app_data_dir,
    export_configuration,
    import_configuration,
    load_settings,
    reset_to_defaults,
    save_settings,
    set_admin_passcode,
    verify_admin_passcode,
)
from update_manager import UpdateError, inspect_update_package, launch_staged_update, stage_update_package, updater_executable_path
from history_db import recent_events, clear_history


class SettingsDialog(tk.Toplevel):
    def __init__(
        self, parent, app_revision: str, tesseract_test_callback=None,
        on_settings_saved=None, show_changelog_callback=None, show_help_callback=None
    ):
        super().__init__(parent)
        self.parent = parent
        self.app_revision = app_revision
        self.tesseract_test_callback = tesseract_test_callback
        self.on_settings_saved = on_settings_saved
        self.show_changelog_callback = show_changelog_callback
        self.show_help_callback = show_help_callback
        self.admin_unlocked = False
        self._admin_timeout_id = None

        self.title(f"Settings - Allison's Data Extractor {app_revision}")
        self.geometry("1120x760")
        self.minsize(940, 650)
        self.transient(parent)

        self.settings = load_settings()

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        search_row = ttk.Frame(outer)
        search_row.pack(fill="x", pady=(0, 8))
        ttk.Label(search_row, text="Search settings:").pack(side="left")
        self.settings_search_var = tk.StringVar()
        search_entry = ttk.Entry(search_row, textvariable=self.settings_search_var, width=34)
        search_entry.pack(side="left", padx=(8, 6))
        search_entry.bind("<Return>", lambda _event: self._search_settings())
        ttk.Button(search_row, text="Find", command=self._search_settings).pack(side="left")
        self.search_result_var = tk.StringVar(value="")
        ttk.Label(search_row, textvariable=self.search_result_var, foreground="#555555").pack(side="left", padx=12)

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)

        self.general_tab = ttk.Frame(self.notebook, padding=14)
        self.review_tab = ttk.Frame(self.notebook, padding=14)
        self.excel_tab = ttk.Frame(self.notebook, padding=14)
        self.format_tab = ttk.Frame(self.notebook, padding=14)
        self.ocr_tab = ttk.Frame(self.notebook, padding=14)
        self.templates_tab = ttk.Frame(self.notebook, padding=14)
        self.history_tab = ttk.Frame(self.notebook, padding=14)
        self.diagnostics_tab = ttk.Frame(self.notebook, padding=14)
        self.about_tab = ttk.Frame(self.notebook, padding=14)
        self.backend_tab = ttk.Frame(self.notebook, padding=14)

        self.notebook.add(self.general_tab, text="General")
        self.notebook.add(self.review_tab, text="Review & Accuracy")
        self.notebook.add(self.excel_tab, text="Excel & Export")
        self.notebook.add(self.format_tab, text="Data Formatting")
        self.notebook.add(self.ocr_tab, text="OCR")
        self.notebook.add(self.templates_tab, text="Templates")
        self.notebook.add(self.history_tab, text="History & Audit")
        self.notebook.add(self.diagnostics_tab, text="Diagnostics")
        self.notebook.add(self.about_tab, text="About")
        self.notebook.add(self.backend_tab, text="Backend / Administrator")

        self._build_general()
        self._build_review()
        self._build_excel()
        self._build_formatting()
        self._build_ocr()
        self._build_templates()
        self._build_history()
        self._build_diagnostics()
        self._build_about()
        self._build_backend_locked()

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Close", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Save Settings", command=self.save_normal_settings).pack(side="right", padx=(0, 6))

    # ------------------------------------------------------------------
    # Normal settings
    # ------------------------------------------------------------------
    def _build_general(self):
        ttk.Label(self.general_tab, text="General", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.general_tab,
            text="Preferences stored here apply to the current Windows user and survive portable application updates.",
            wraplength=720,
        ).pack(anchor="w", pady=(4, 14))

        general = self.settings.get("general", {})
        self.remember_output_var = tk.BooleanVar(value=bool(general.get("remember_output_file", True)))
        ttk.Checkbutton(
            self.general_tab,
            text="Remember the last selected Excel output file",
            variable=self.remember_output_var,
        ).pack(anchor="w", pady=4)

        location = ttk.LabelFrame(self.general_tab, text="Settings storage", padding=10)
        location.pack(fill="x", pady=(18, 0))
        ttk.Label(location, text=str(app_data_dir()), wraplength=700).pack(anchor="w")
        ttk.Label(
            location,
            text="History, settings, configuration backups and staged updates are stored here rather than inside the EXE.",
            foreground="#555555",
            wraplength=700,
        ).pack(anchor="w", pady=(4, 0))

    def _search_settings(self):
        query = self.settings_search_var.get().strip().casefold()
        if not query:
            self.search_result_var.set("Enter a setting or feature name.")
            return
        index = {
            "General": "remember output workbook startup close preferences",
            "Review & Accuracy": "review accuracy confidence verified approval production safety missing required warnings",
            "Excel & Export": "excel export backup duplicate confirm output folder workbook",
            "Data Formatting": "format formatting pack style normalization alias separator",
            "OCR": "ocr tesseract template b threshold penalty image",
            "Templates": "template markers a b c d e f detection source type",
            "History & Audit": "history audit log edits approvals exports clear csv retention",
            "Diagnostics": "diagnostics health check tesseract database settings updater permissions support package",
            "About": "about version revision changelog help developer application",
            "Backend / Administrator": "backend administrator passcode security timeout update import export configuration reset advanced",
        }
        words = [w for w in query.replace("/", " ").split() if w]
        best = None
        best_score = 0
        for tab_name, keywords in index.items():
            haystack = (tab_name + " " + keywords).casefold()
            score = sum(1 for word in words if word in haystack)
            if score > best_score:
                best_score = score
                best = tab_name
        if not best:
            self.search_result_var.set("No matching settings category found.")
            return
        for tab_id in self.notebook.tabs():
            if self.notebook.tab(tab_id, "text") == best:
                self.notebook.select(tab_id)
                self.search_result_var.set(f"Showing: {best}")
                return

    def _build_review(self):
        ttk.Label(self.review_tab, text="Review & Accuracy", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.review_tab,
            text="Safety controls determine when a record must be reviewed or approved before it can be exported.",
            wraplength=800,
        ).pack(anchor="w", pady=(4, 14))

        review = self.settings.get("review", {})
        self.production_safety_var = tk.BooleanVar(value=bool(review.get("production_safety_mode", True)))
        self.require_approval_var = tk.BooleanVar(value=bool(review.get("require_approval_below_verified", True)))
        self.block_missing_var = tk.BooleanVar(value=bool(review.get("block_missing_required_fields", True)))
        self.warn_warnings_var = tk.BooleanVar(value=bool(review.get("warn_before_export_with_warnings", True)))

        safety = ttk.LabelFrame(self.review_tab, text="Production Safety Mode", padding=12)
        safety.pack(fill="x")
        ttk.Checkbutton(
            safety,
            text="Enable Production Safety Mode",
            variable=self.production_safety_var,
        ).pack(anchor="w")
        ttk.Label(
            safety,
            text="Recommended for production use. When enabled, required-field checks and approval gates are enforced before Excel export.",
            foreground="#555555",
            wraplength=780,
        ).pack(anchor="w", pady=(4, 8))
        ttk.Checkbutton(
            safety,
            text="Require manual approval for records below the Verified confidence threshold",
            variable=self.require_approval_var,
        ).pack(anchor="w", pady=3)
        ttk.Checkbutton(
            safety,
            text="Block export when required shipment fields are missing",
            variable=self.block_missing_var,
        ).pack(anchor="w", pady=3)
        ttk.Checkbutton(
            safety,
            text="Warn before exporting records that still contain validation warnings",
            variable=self.warn_warnings_var,
        ).pack(anchor="w", pady=3)

        note = ttk.LabelFrame(self.review_tab, text="Accuracy principle", padding=12)
        note.pack(fill="x", pady=(16, 0))
        ttk.Label(
            note,
            text="These controls never invent or silently correct shipment data. They decide when human review is required before data can leave the review queue.",
            wraplength=780,
        ).pack(anchor="w")

    def _build_excel(self):
        ttk.Label(self.excel_tab, text="Excel & Export", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.excel_tab,
            text="Configure safeguards around the final write into the selected Excel workbook.",
            wraplength=800,
        ).pack(anchor="w", pady=(4, 14))

        general = self.settings.get("general", {})
        excel = self.settings.get("excel", {})
        self.confirm_export_var = tk.BooleanVar(value=bool(general.get("confirm_before_export", True)))
        self.open_output_folder_var = tk.BooleanVar(value=bool(general.get("open_output_folder_after_export", False)))
        self.automatic_backup_var = tk.BooleanVar(value=bool(excel.get("automatic_backup", True)))
        self.duplicate_check_var = tk.BooleanVar(value=bool(excel.get("duplicate_check", True)))

        box = ttk.LabelFrame(self.excel_tab, text="Export safeguards", padding=12)
        box.pack(fill="x")
        ttk.Checkbutton(box, text="Show confirmation/warning before export", variable=self.confirm_export_var).pack(anchor="w", pady=3)
        ttk.Checkbutton(box, text="Create automatic timestamped backup before modifying an existing workbook", variable=self.automatic_backup_var).pack(anchor="w", pady=3)
        ttk.Checkbutton(box, text="Check Sales Order Numbers for possible duplicates before export", variable=self.duplicate_check_var).pack(anchor="w", pady=3)
        ttk.Checkbutton(box, text="Open the output folder after a successful export", variable=self.open_output_folder_var).pack(anchor="w", pady=3)

        ttk.Label(
            self.excel_tab,
            text="Production Safety Mode can override unsafe export choices, such as disabling backups.",
            foreground="#555555",
            wraplength=800,
        ).pack(anchor="w", pady=(12, 0))

    def _build_templates(self):
        ttk.Label(self.templates_tab, text="Templates", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.templates_tab,
            text="This page is a read-only overview. Unlock Backend / Administrator to change how documents are detected or read.",
            wraplength=800,
        ).pack(anchor="w", pady=(4, 12))

        cols = ("Template", "Source", "Marker(s)", "Lot Required", "Status")
        tree = ttk.Treeview(self.templates_tab, columns=cols, show="headings", height=10)
        for col, width in (("Template", 140), ("Source", 100), ("Marker(s)", 300), ("Lot Required", 100), ("Status", 100)):
            tree.heading(col, text=col)
            tree.column(col, width=width, anchor="w")
        tree.pack(fill="x")
        for name, cfg in sorted((self.settings.get("templates", {}) or {}).items()):
            tree.insert("", "end", values=(
                name,
                str(cfg.get("source_type", "")),
                ", ".join(cfg.get("markers", []) or []),
                "Yes" if cfg.get("lot_required", False) else "No",
                "Active",
            ))

        row = ttk.Frame(self.templates_tab)
        row.pack(fill="x", pady=(12, 0))
        ttk.Button(row, text="Open Backend / Administrator", command=lambda: self.notebook.select(self.backend_tab)).pack(side="left")
        ttk.Label(
            row,
            text="Future templates D, E, F, etc. can be added through the configuration framework as the extractor expands.",
            foreground="#555555",
            wraplength=620,
        ).pack(side="left", padx=12)

    def _build_history(self):
        ttk.Label(self.history_tab, text="History & Audit", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.history_tab,
            text="Choose which user actions are written to the local audit history. Extraction events remain recorded for traceability.",
            wraplength=800,
        ).pack(anchor="w", pady=(4, 12))
        history = self.settings.get("history", {})
        self.log_edits_var = tk.BooleanVar(value=bool(history.get("log_edits", True)))
        self.log_approvals_var = tk.BooleanVar(value=bool(history.get("log_approvals", True)))
        self.log_exports_var = tk.BooleanVar(value=bool(history.get("log_exports", True)))
        box = ttk.LabelFrame(self.history_tab, text="Audit events", padding=12)
        box.pack(fill="x")
        ttk.Checkbutton(box, text="Log manual field edits", variable=self.log_edits_var).pack(anchor="w", pady=3)
        ttk.Checkbutton(box, text="Log manual approvals", variable=self.log_approvals_var).pack(anchor="w", pady=3)
        ttk.Checkbutton(box, text="Log Excel exports", variable=self.log_exports_var).pack(anchor="w", pady=3)
        row = ttk.Frame(self.history_tab)
        row.pack(fill="x", pady=(16, 0))
        ttk.Button(row, text="Export History CSV...", command=self._export_history_csv).pack(side="left")
        ttk.Button(row, text="Clear Processing History...", command=self._clear_history_from_settings).pack(side="left", padx=6)
        ttk.Label(row, text="Clearing history is permanent and requires confirmation.", foreground="#555555").pack(side="left", padx=8)

    def _build_diagnostics(self):
        ttk.Label(self.diagnostics_tab, text="Diagnostics", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.diagnostics_tab,
            text="Run a local health check before deployment or when extraction, OCR, history, Excel, or updates behave unexpectedly.",
            wraplength=800,
        ).pack(anchor="w", pady=(4, 12))

        self.diagnostics_text = tk.Text(self.diagnostics_tab, height=18, wrap="word", font=("Segoe UI", 10), state="disabled")
        self.diagnostics_text.pack(fill="both", expand=True)
        row = ttk.Frame(self.diagnostics_tab)
        row.pack(fill="x", pady=(10, 0))
        ttk.Button(row, text="Run Full Diagnostic", command=self._run_diagnostics).pack(side="left")
        ttk.Button(row, text="Create Diagnostic Package...", command=self._create_diagnostic_package).pack(side="left", padx=6)
        ttk.Label(row, text="Diagnostic packages do not include shipment source files or extracted shipment data.", foreground="#555555").pack(side="left", padx=8)

    def _build_about(self):
        ttk.Label(self.about_tab, text="About", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(self.about_tab, text="Allison's Data Extractor", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(14, 2))
        ttk.Label(self.about_tab, text=f"Revision {self.app_revision}", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        ttk.Label(self.about_tab, text="Developed by Alexander Hostine", font=("Segoe UI", 10)).pack(anchor="w", pady=(8, 0))
        ttk.Label(
            self.about_tab,
            text="Document and image intake, structured extraction, validation, review, and Excel export.",
            wraplength=760,
        ).pack(anchor="w", pady=(14, 18))

        buttons = ttk.Frame(self.about_tab)
        buttons.pack(anchor="w")
        ttk.Button(buttons, text="Changelog", command=self._show_changelog).pack(side="left")
        ttk.Button(buttons, text="Help", command=self._show_help).pack(side="left", padx=6)
        ttk.Button(buttons, text="Open Application Data Folder", command=self._open_app_data_folder).pack(side="left", padx=6)

        ttk.Separator(self.about_tab).pack(fill="x", pady=18)
        ttk.Label(self.about_tab, text=f"Settings location: {app_data_dir()}", foreground="#555555", wraplength=820).pack(anchor="w")

    def _build_formatting(self):
        ttk.Label(self.format_tab, text="Data Formatting", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.format_tab,
            text="Pack Style normalization keeps equivalent packaging descriptions consistent before validation and Excel export.",
            wraplength=720,
        ).pack(anchor="w", pady=(4, 12))

        formatting = self.settings.get("formatting", {})
        self.normalize_pack_var = tk.BooleanVar(value=bool(formatting.get("normalize_pack_style", True)))
        ttk.Checkbutton(
            self.format_tab,
            text="Normalize Pack Style values",
            variable=self.normalize_pack_var,
        ).pack(anchor="w", pady=4)

        sep_frame = ttk.Frame(self.format_tab)
        sep_frame.pack(fill="x", pady=(8, 6))
        ttk.Label(sep_frame, text="Multiple-pack separator:").pack(side="left")
        self.pack_separator_var = tk.StringVar(value=str(formatting.get("pack_style_separator", " / ")))
        ttk.Entry(sep_frame, textvariable=self.pack_separator_var, width=12).pack(side="left", padx=8)

        ttk.Label(
            self.format_tab,
            text="Pack Style aliases (one per line: raw value = preferred value)",
            font=("TkDefaultFont", 9, "bold"),
        ).pack(anchor="w", pady=(12, 4))

        alias_frame = ttk.Frame(self.format_tab)
        alias_frame.pack(fill="both", expand=True)
        scrollbar = ttk.Scrollbar(alias_frame, orient="vertical")
        scrollbar.pack(side="right", fill="y")
        self.alias_text = tk.Text(alias_frame, height=15, wrap="none", yscrollcommand=scrollbar.set)
        self.alias_text.pack(side="left", fill="both", expand=True)
        scrollbar.configure(command=self.alias_text.yview)

        aliases = formatting.get("pack_style_aliases", {}) or {}
        self.alias_text.insert("1.0", "\n".join(f"{key} = {value}" for key, value in aliases.items()))

    def _build_ocr(self):
        ttk.Label(self.ocr_tab, text="OCR / Template B", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.ocr_tab,
            text="Template B uses Tesseract OCR. These settings control its review penalty and validation thresholds.",
            wraplength=720,
        ).pack(anchor="w", pady=(4, 12))

        ocr = self.settings.get("ocr", {})

        grid = ttk.Frame(self.ocr_tab)
        grid.pack(fill="x")

        ttk.Label(grid, text="Template B OCR confidence penalty:").grid(row=0, column=0, sticky="w", pady=5)
        self.ocr_penalty_var = tk.IntVar(value=int(ocr.get("template_b_penalty", 10)))
        ttk.Spinbox(grid, from_=0, to=50, textvariable=self.ocr_penalty_var, width=8).grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(grid, text="Verified threshold:").grid(row=1, column=0, sticky="w", pady=5)
        self.verified_threshold_var = tk.IntVar(value=int(ocr.get("verified_threshold", 95)))
        ttk.Spinbox(grid, from_=50, to=100, textvariable=self.verified_threshold_var, width=8).grid(row=1, column=1, sticky="w", padx=8)

        ttk.Label(grid, text="Review threshold:").grid(row=2, column=0, sticky="w", pady=5)
        self.review_threshold_var = tk.IntVar(value=int(ocr.get("review_threshold", 70)))
        ttk.Spinbox(grid, from_=0, to=99, textvariable=self.review_threshold_var, width=8).grid(row=2, column=1, sticky="w", padx=8)

        override = ttk.LabelFrame(self.ocr_tab, text="Tesseract executable override (optional)", padding=10)
        override.pack(fill="x", pady=(16, 0))
        self.tesseract_override_var = tk.StringVar(value=str(ocr.get("tesseract_cmd_override", "")))
        ttk.Entry(override, textvariable=self.tesseract_override_var).pack(side="left", fill="x", expand=True)
        ttk.Button(override, text="Browse...", command=self._browse_tesseract).pack(side="left", padx=(6, 0))
        ttk.Button(override, text="Test Tesseract", command=self._test_tesseract).pack(side="left", padx=(6, 0))

        ttk.Label(
            self.ocr_tab,
            text="Leave the override blank to use portable Tesseract beside/inside the application or a normal Windows installation.",
            foreground="#555555",
            wraplength=720,
        ).pack(anchor="w", pady=(6, 0))

    def _parse_aliases(self) -> dict:
        aliases = {}
        for line_number, line in enumerate(self.alias_text.get("1.0", "end").splitlines(), start=1):
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "=" not in stripped:
                raise ValueError(f"Pack Style alias line {line_number} must contain '='.")
            raw, preferred = stripped.split("=", 1)
            raw = raw.strip()
            preferred = preferred.strip()
            if not raw or not preferred:
                raise ValueError(f"Pack Style alias line {line_number} is incomplete.")
            aliases[raw] = preferred
        return aliases

    def save_normal_settings(self):
        try:
            aliases = self._parse_aliases()
            verified = int(self.verified_threshold_var.get())
            review = int(self.review_threshold_var.get())
            if review >= verified:
                raise ValueError("Review threshold must be lower than the Verified threshold.")

            settings = load_settings()
            general = settings.setdefault("general", {})
            general["remember_output_file"] = bool(self.remember_output_var.get())
            general["confirm_before_export"] = bool(self.confirm_export_var.get())
            general["open_output_folder_after_export"] = bool(self.open_output_folder_var.get())

            review_settings = settings.setdefault("review", {})
            review_settings["production_safety_mode"] = bool(self.production_safety_var.get())
            review_settings["require_approval_below_verified"] = bool(self.require_approval_var.get())
            review_settings["block_missing_required_fields"] = bool(self.block_missing_var.get())
            review_settings["warn_before_export_with_warnings"] = bool(self.warn_warnings_var.get())

            excel_settings = settings.setdefault("excel", {})
            excel_settings["automatic_backup"] = bool(self.automatic_backup_var.get())
            excel_settings["duplicate_check"] = bool(self.duplicate_check_var.get())

            history_settings = settings.setdefault("history", {})
            history_settings["log_edits"] = bool(self.log_edits_var.get())
            history_settings["log_approvals"] = bool(self.log_approvals_var.get())
            history_settings["log_exports"] = bool(self.log_exports_var.get())

            settings.setdefault("formatting", {})["normalize_pack_style"] = bool(self.normalize_pack_var.get())
            settings["formatting"]["pack_style_separator"] = self.pack_separator_var.get() or " / "
            settings["formatting"]["pack_style_aliases"] = aliases
            settings.setdefault("ocr", {})["template_b_penalty"] = int(self.ocr_penalty_var.get())
            settings["ocr"]["verified_threshold"] = verified
            settings["ocr"]["review_threshold"] = review
            settings["ocr"]["tesseract_cmd_override"] = self.tesseract_override_var.get().strip()
            save_settings(settings, create_backup=False)
            self.settings = settings
            if self.on_settings_saved:
                self.on_settings_saved()
            messagebox.showinfo("Settings saved", "Settings were saved successfully.", parent=self)
        except Exception as exc:
            messagebox.showerror("Settings error", str(exc), parent=self)

    def _browse_tesseract(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select tesseract.exe",
            filetypes=[("Tesseract executable", "tesseract.exe"), ("Executable files", "*.exe"), ("All files", "*.*")],
        )
        if path:
            self.tesseract_override_var.set(path)

    def _test_tesseract(self):
        # Save the override temporarily so the callback tests exactly what is displayed.
        try:
            settings = load_settings()
            settings.setdefault("ocr", {})["tesseract_cmd_override"] = self.tesseract_override_var.get().strip()
            save_settings(settings, create_backup=False)
            if self.tesseract_test_callback:
                ok, detail = self.tesseract_test_callback()
                if ok:
                    messagebox.showinfo("Tesseract test", detail, parent=self)
                else:
                    messagebox.showwarning("Tesseract test", detail, parent=self)
            else:
                messagebox.showinfo("Tesseract test", "Tesseract test callback is unavailable.", parent=self)
        except Exception as exc:
            messagebox.showerror("Tesseract test", str(exc), parent=self)

    def _show_changelog(self):
        if self.show_changelog_callback:
            self.show_changelog_callback()
        else:
            messagebox.showinfo("Changelog", "Changelog viewer is unavailable in this build.", parent=self)

    def _show_help(self):
        if self.show_help_callback:
            self.show_help_callback()
        else:
            messagebox.showinfo("Help", "Help viewer is unavailable in this build.", parent=self)

    def _open_app_data_folder(self):
        path = str(app_data_dir())
        try:
            if os.name == "nt":
                os.startfile(path)
            else:
                import subprocess
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            messagebox.showerror("Open folder", str(exc), parent=self)

    def _export_history_csv(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Export processing history",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"Allisons_Data_Extractor_History_{datetime.now().strftime('%Y-%m-%d')}.csv",
        )
        if not path:
            return
        try:
            rows = recent_events(100000)
            with open(path, "w", newline="", encoding="utf-8-sig") as handle:
                writer = csv.writer(handle)
                writer.writerow(["Time", "Action", "Source", "Template", "Sales Order", "Status", "Output", "Details"])
                writer.writerows(rows)
            messagebox.showinfo("History exported", f"Exported {len(rows)} history event(s) to:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("History export", str(exc), parent=self)

    def _clear_history_from_settings(self):
        if not messagebox.askyesno(
            "Clear processing history",
            "Permanently delete all local processing history?\n\nThis cannot be undone.",
            parent=self,
        ):
            return
        try:
            removed = clear_history()
            messagebox.showinfo("History cleared", f"Removed {removed} history event(s).", parent=self)
        except Exception as exc:
            messagebox.showerror("History error", str(exc), parent=self)

    def _diagnostic_results(self):
        results = []
        data_dir = app_data_dir()

        try:
            data_dir.mkdir(parents=True, exist_ok=True)
            test_file = data_dir / ".write_test.tmp"
            test_file.write_text("ok", encoding="utf-8")
            test_file.unlink(missing_ok=True)
            results.append(("Application data write access", True, str(data_dir)))
        except Exception as exc:
            results.append(("Application data write access", False, str(exc)))

        try:
            settings = load_settings()
            results.append(("Settings configuration", isinstance(settings, dict), f"Schema {settings.get('schema_version', '?')}"))
        except Exception as exc:
            results.append(("Settings configuration", False, str(exc)))

        try:
            import openpyxl
            results.append(("Excel engine (openpyxl)", True, getattr(openpyxl, "__version__", "installed")))
        except Exception as exc:
            results.append(("Excel engine (openpyxl)", False, str(exc)))

        try:
            rows = recent_events(1)
            results.append(("History database", True, f"Accessible; recent rows read: {len(rows)}"))
        except Exception as exc:
            results.append(("History database", False, str(exc)))

        if self.tesseract_test_callback:
            try:
                ok, detail = self.tesseract_test_callback()
                results.append(("Tesseract OCR", bool(ok), str(detail).replace("\n", " | ")))
            except Exception as exc:
                results.append(("Tesseract OCR", False, str(exc)))
        else:
            results.append(("Tesseract OCR", False, "Test callback unavailable"))

        updater = updater_executable_path()
        if getattr(sys, "frozen", False):
            results.append(("Updater helper", bool(updater), updater or "Updater executable not found"))
        else:
            results.append(("Updater helper", True, "Development mode; compiled updater check skipped"))

        return results

    def _run_diagnostics(self):
        results = self._diagnostic_results()
        lines = [
            f"Allison's Data Extractor {self.app_revision}",
            f"Diagnostic run: {datetime.now().strftime('%B %d, %Y %I:%M:%S %p')}",
            "",
        ]
        for name, ok, detail in results:
            lines.append(f"{'PASS' if ok else 'FAIL'}  {name}")
            lines.append(f"      {detail}")
        lines += ["", f"Overall: {'PASS' if all(item[1] for item in results) else 'ATTENTION REQUIRED'}"]
        self.diagnostics_text.configure(state="normal")
        self.diagnostics_text.delete("1.0", "end")
        self.diagnostics_text.insert("1.0", "\n".join(lines))
        self.diagnostics_text.configure(state="disabled")
        return lines

    def _create_diagnostic_package(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Create diagnostic package",
            defaultextension=".zip",
            filetypes=[("ZIP files", "*.zip")],
            initialfile=f"Allisons_Data_Extractor_Diagnostics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        )
        if not path:
            return
        try:
            lines = self._run_diagnostics()
            settings = load_settings()
            settings["admin"] = {"note": "Administrator credential material intentionally omitted."}
            system_info = {
                "application_revision": self.app_revision,
                "python": sys.version,
                "platform": platform.platform(),
                "frozen": bool(getattr(sys, "frozen", False)),
                "created_at": datetime.now().isoformat(timespec="seconds"),
            }
            with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.writestr("diagnostic_report.txt", "\n".join(lines))
                archive.writestr("system_info.json", json.dumps(system_info, indent=2))
                archive.writestr("settings_sanitized.json", json.dumps(settings, indent=2, ensure_ascii=False))
                archive.writestr(
                    "README.txt",
                    "This package intentionally excludes shipment source files, extracted shipment records, processing history, and administrator credentials.\n",
                )
            messagebox.showinfo("Diagnostic package", f"Diagnostic package created:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Diagnostic package", str(exc), parent=self)

    # ------------------------------------------------------------------
    # Backend / administrator
    # ------------------------------------------------------------------
    def _clear_backend_tab(self):
        for child in self.backend_tab.winfo_children():
            child.destroy()

    def _build_backend_locked(self):
        self._clear_backend_tab()
        ttk.Label(self.backend_tab, text="Backend / Administrator", font=("TkDefaultFont", 12, "bold")).pack(anchor="w")
        ttk.Label(
            self.backend_tab,
            text="Administrator mode controls document-detection rules, validation settings, configuration import/export and software updates.",
            wraplength=720,
        ).pack(anchor="w", pady=(4, 16))

        box = ttk.LabelFrame(self.backend_tab, text="Administrator access", padding=14)
        box.pack(fill="x")

        if not admin_is_configured():
            ttk.Label(box, text="No administrator passcode has been created yet. Create one now.").pack(anchor="w")
            self.new_pass_var = tk.StringVar()
            self.confirm_pass_var = tk.StringVar()
            row1 = ttk.Frame(box); row1.pack(fill="x", pady=(10, 4))
            ttk.Label(row1, text="New passcode:", width=18).pack(side="left")
            ttk.Entry(row1, textvariable=self.new_pass_var, show="•").pack(side="left", fill="x", expand=True)
            row2 = ttk.Frame(box); row2.pack(fill="x", pady=4)
            ttk.Label(row2, text="Confirm passcode:", width=18).pack(side="left")
            ttk.Entry(row2, textvariable=self.confirm_pass_var, show="•").pack(side="left", fill="x", expand=True)
            ttk.Button(box, text="Create Administrator Passcode", command=self._create_admin_passcode).pack(anchor="e", pady=(10, 0))
        else:
            ttk.Label(box, text="Enter the administrator passcode to unlock backend configuration.").pack(anchor="w")
            self.unlock_pass_var = tk.StringVar()
            row = ttk.Frame(box); row.pack(fill="x", pady=(10, 4))
            ttk.Label(row, text="Passcode:", width=18).pack(side="left")
            entry = ttk.Entry(row, textvariable=self.unlock_pass_var, show="•")
            entry.pack(side="left", fill="x", expand=True)
            entry.bind("<Return>", lambda _event: self._unlock_backend())
            ttk.Button(box, text="Unlock", command=self._unlock_backend).pack(anchor="e", pady=(8, 0))

    def _create_admin_passcode(self):
        if self.new_pass_var.get() != self.confirm_pass_var.get():
            messagebox.showerror("Passcode", "The passcodes do not match.", parent=self)
            return
        try:
            set_admin_passcode(self.new_pass_var.get())
        except Exception as exc:
            messagebox.showerror("Passcode", str(exc), parent=self)
            return
        self.admin_unlocked = True
        self.settings = load_settings()
        self._start_admin_timeout()
        self._build_backend_unlocked()

    def _unlock_backend(self):
        if not verify_admin_passcode(self.unlock_pass_var.get()):
            messagebox.showerror("Administrator access", "Incorrect administrator passcode.", parent=self)
            return
        self.admin_unlocked = True
        self.settings = load_settings()
        self._start_admin_timeout()
        self._build_backend_unlocked()

    def _start_admin_timeout(self):
        if self._admin_timeout_id is not None:
            try:
                self.after_cancel(self._admin_timeout_id)
            except Exception:
                pass
        timeout_minutes = int(load_settings().get("security", {}).get("admin_timeout_minutes", 15) or 15)
        timeout_minutes = max(1, min(timeout_minutes, 240))
        self._admin_timeout_id = self.after(timeout_minutes * 60 * 1000, self._expire_admin_session)

    def _expire_admin_session(self):
        self._admin_timeout_id = None
        if self.admin_unlocked:
            self.admin_unlocked = False
            self._build_backend_locked()
            timeout_minutes = int(load_settings().get("security", {}).get("admin_timeout_minutes", 15) or 15)
            messagebox.showinfo(
                "Administrator session locked",
                f"Administrator mode was automatically locked after {timeout_minutes} minute(s).",
                parent=self,
            )

    def _build_backend_unlocked(self):
        self._clear_backend_tab()
        top = ttk.Frame(self.backend_tab)
        top.pack(fill="x")
        ttk.Label(top, text="Backend / Administrator", font=("TkDefaultFont", 12, "bold")).pack(side="left")
        ttk.Label(top, text="🔓 Administrator Mode").pack(side="left", padx=14)
        ttk.Button(top, text="Lock", command=self._lock_backend).pack(side="right")

        canvas = tk.Canvas(self.backend_tab, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.backend_tab, orient="vertical", command=canvas.yview)
        content = ttk.Frame(canvas, padding=(0, 8, 8, 8))
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        version_box = ttk.LabelFrame(content, text="Application", padding=10)
        version_box.pack(fill="x", pady=(0, 8))
        ttk.Label(version_box, text=f"Current revision: {self.app_revision}", font=("TkDefaultFont", 9, "bold")).pack(anchor="w")

        templates_box = ttk.LabelFrame(content, text="Template detection / validation", padding=10)
        templates_box.pack(fill="x", pady=8)

        self.backend_vars = {}
        templates = self.settings.get("templates", {})
        template_names = sorted(templates) or ["Template A", "Template B", "Template C"]
        for row_index, name in enumerate(template_names):
            cfg = templates.get(name, {})
            frame = ttk.LabelFrame(templates_box, text=name, padding=8)
            frame.pack(fill="x", pady=4)
            marker_var = tk.StringVar(value=", ".join(cfg.get("markers", [])))
            lot_var = tk.BooleanVar(value=bool(cfg.get("lot_required", name != "Template C")))
            sales_var = tk.StringVar(value=str(cfg.get("sales_order_pattern", r"^\d{5,9}$")))
            ttk.Label(frame, text="Marker(s):", width=18).grid(row=0, column=0, sticky="w", pady=2)
            ttk.Entry(frame, textvariable=marker_var).grid(row=0, column=1, sticky="ew", pady=2)
            ttk.Label(frame, text="Sales order regex:", width=18).grid(row=1, column=0, sticky="w", pady=2)
            ttk.Entry(frame, textvariable=sales_var).grid(row=1, column=1, sticky="ew", pady=2)
            ttk.Checkbutton(frame, text="Lot Code required", variable=lot_var).grid(row=2, column=1, sticky="w", pady=2)
            frame.grid_columnconfigure(1, weight=1)
            self.backend_vars[name] = {"markers": marker_var, "lot_required": lot_var, "sales_order_pattern": sales_var}

        b_cfg = templates.get("Template B", {})
        self.record_start_var = tk.StringVar(value=", ".join(b_cfg.get("record_start_terms", ["Plant Species"])))
        row = ttk.Frame(templates_box); row.pack(fill="x", pady=(6, 0))
        ttk.Label(row, text="Template B record start term(s):").pack(side="left")
        ttk.Entry(row, textvariable=self.record_start_var).pack(side="left", fill="x", expand=True, padx=(8, 0))

        template_buttons = ttk.Frame(templates_box)
        template_buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(template_buttons, text="Test Configuration with Sample...", command=self._test_configuration).pack(side="left")
        ttk.Button(template_buttons, text="Save Backend Configuration", command=self._save_backend).pack(side="right")

        config_box = ttk.LabelFrame(content, text="Configuration tools", padding=10)
        config_box.pack(fill="x", pady=8)
        ttk.Button(config_box, text="Advanced JSON Editor...", command=self._advanced_json_editor).pack(side="left", padx=(0, 6))
        ttk.Button(config_box, text="Export Configuration...", command=self._export_config).pack(side="left", padx=6)
        ttk.Button(config_box, text="Import Configuration...", command=self._import_config).pack(side="left", padx=6)
        ttk.Button(config_box, text="Reset Defaults...", command=self._reset_defaults).pack(side="right")

        security_box = ttk.LabelFrame(content, text="Security", padding=10)
        security_box.pack(fill="x", pady=8)
        security = self.settings.get("security", {})
        self.admin_timeout_var = tk.IntVar(value=int(security.get("admin_timeout_minutes", 15) or 15))
        timeout_row = ttk.Frame(security_box)
        timeout_row.pack(fill="x")
        ttk.Label(timeout_row, text="Auto-lock administrator mode after:").pack(side="left")
        ttk.Spinbox(timeout_row, from_=1, to=240, textvariable=self.admin_timeout_var, width=8).pack(side="left", padx=8)
        ttk.Label(timeout_row, text="minutes").pack(side="left")
        ttk.Button(security_box, text="Change Administrator Passcode...", command=self._change_passcode).pack(anchor="w", pady=(10, 0))

        update_box = ttk.LabelFrame(content, text="Software Update", padding=10)
        update_box.pack(fill="x", pady=8)
        ttk.Label(
            update_box,
            text="Install a newer update package from a USB drive or local folder. The application validates the package, closes, installs it, and restarts automatically.",
            wraplength=690,
        ).pack(anchor="w")
        ttk.Button(update_box, text="Load Update Package...", command=self._install_update).pack(anchor="e", pady=(10, 0))

    def _lock_backend(self):
        if self._admin_timeout_id is not None:
            try:
                self.after_cancel(self._admin_timeout_id)
            except Exception:
                pass
            self._admin_timeout_id = None
        self.admin_unlocked = False
        self._build_backend_locked()

    def _test_configuration(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select a sample shipment file",
            filetypes=[
                ("Supported shipment files", "*.pdf *.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            # Save the visible backend fields first so the sample tests exactly
            # what the administrator is about to deploy. A backup is created.
            settings = load_settings()
            templates = settings.setdefault("templates", {})
            for name, variables in self.backend_vars.items():
                cfg = templates.setdefault(name, {})
                cfg["markers"] = [part.strip() for part in variables["markers"].get().split(",") if part.strip()]
                cfg["lot_required"] = bool(variables["lot_required"].get())
                cfg["sales_order_pattern"] = variables["sales_order_pattern"].get().strip() or r"^\d{5,9}$"
            templates.setdefault("Template B", {})["record_start_terms"] = [
                part.strip() for part in self.record_start_var.get().split(",") if part.strip()
            ] or ["Plant Species"]
            save_settings(settings, create_backup=True)

            import pdf_extractor
            pdf_extractor.refresh_tesseract_configuration()
            result = pdf_extractor.process_pdf(path)
            if result.get("error"):
                raise RuntimeError(result["error"])

            records = result.get("records") or [result.get("fields", {})]
            lines = [
                f"Detected template: {result.get('template', '')}",
                f"Records detected: {len(records)}",
                "",
            ]
            for index, record in enumerate(records, start=1):
                lines.append(f"Record {index}")
                for field in pdf_extractor.HEADERS:
                    lines.append(f"  {field}: {record.get(field, '')}")
                warnings = (result.get("record_warnings") or [[]])[index - 1] if index - 1 < len(result.get("record_warnings") or []) else []
                confidence = (result.get("record_confidence") or [""])[index - 1] if index - 1 < len(result.get("record_confidence") or []) else ""
                lines.append(f"  Confidence: {confidence}%")
                lines.append(f"  Warnings: {'; '.join(warnings) if warnings else 'None'}")
                lines.append("")

            win = tk.Toplevel(self)
            win.title("Configuration Test Results")
            win.geometry("820x620")
            viewer = tk.Text(win, wrap="word", padx=10, pady=10)
            viewer.pack(fill="both", expand=True)
            viewer.insert("1.0", "\n".join(lines))
            viewer.configure(state="disabled")
            ttk.Button(win, text="Close", command=win.destroy).pack(anchor="e", padx=10, pady=10)
        except Exception as exc:
            messagebox.showerror("Configuration test failed", str(exc), parent=self)

    def _save_backend(self):
        try:
            settings = load_settings()
            templates = settings.setdefault("templates", {})
            for name, variables in self.backend_vars.items():
                cfg = templates.setdefault(name, {})
                markers = [part.strip() for part in variables["markers"].get().split(",") if part.strip()]
                if not markers:
                    raise ValueError(f"{name} must have at least one marker.")
                cfg["markers"] = markers
                cfg["lot_required"] = bool(variables["lot_required"].get())
                cfg["sales_order_pattern"] = variables["sales_order_pattern"].get().strip() or r"^\d{5,9}$"
            templates.setdefault("Template B", {})["record_start_terms"] = [
                part.strip() for part in self.record_start_var.get().split(",") if part.strip()
            ] or ["Plant Species"]
            if hasattr(self, "admin_timeout_var"):
                timeout = int(self.admin_timeout_var.get())
                if not 1 <= timeout <= 240:
                    raise ValueError("Administrator auto-lock timeout must be between 1 and 240 minutes.")
                settings.setdefault("security", {})["admin_timeout_minutes"] = timeout
            backup = save_settings(settings, create_backup=True)
            self.settings = load_settings()
            if self.on_settings_saved:
                self.on_settings_saved()
            message = "Backend configuration saved."
            if backup:
                message += f"\n\nPrevious configuration backed up to:\n{backup}"
            messagebox.showinfo("Backend saved", message, parent=self)
        except Exception as exc:
            messagebox.showerror("Backend error", str(exc), parent=self)

    def _advanced_json_editor(self):
        settings = load_settings()
        editable = copy.deepcopy(settings)
        editable["admin"] = {"note": "Administrator credentials are protected and are not editable here."}

        win = tk.Toplevel(self)
        win.title("Advanced Configuration JSON")
        win.geometry("820x650")
        text = tk.Text(win, wrap="none")
        text.pack(fill="both", expand=True, padx=8, pady=8)
        text.insert("1.0", json.dumps(editable, indent=2, ensure_ascii=False))

        def save_json():
            try:
                incoming = json.loads(text.get("1.0", "end"))
                if not isinstance(incoming, dict):
                    raise ValueError("JSON root must be an object.")
                current = load_settings()
                incoming["admin"] = current.get("admin", {})
                save_settings(incoming, create_backup=True)
                self.settings = load_settings()
                if self.on_settings_saved:
                    self.on_settings_saved()
                messagebox.showinfo("Configuration", "Advanced configuration saved.", parent=win)
                win.destroy()
                self._build_backend_unlocked()
            except Exception as exc:
                messagebox.showerror("Configuration error", str(exc), parent=win)

        buttons = ttk.Frame(win, padding=8); buttons.pack(fill="x")
        ttk.Button(buttons, text="Cancel", command=win.destroy).pack(side="right")
        ttk.Button(buttons, text="Validate & Save", command=save_json).pack(side="right", padx=(0, 6))

    def _export_config(self):
        path = filedialog.asksaveasfilename(
            parent=self,
            title="Export extractor configuration",
            defaultextension=".json",
            filetypes=[("JSON configuration", "*.json")],
            initialfile="Allisons_Data_Extractor_Config.json",
        )
        if not path:
            return
        try:
            export_configuration(path)
            messagebox.showinfo("Configuration exported", f"Configuration saved to:\n{path}", parent=self)
        except Exception as exc:
            messagebox.showerror("Export failed", str(exc), parent=self)

    def _import_config(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Import extractor configuration",
            filetypes=[("JSON configuration", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        if not messagebox.askyesno("Import configuration", "Import this configuration and replace the current non-password settings?", parent=self):
            return
        try:
            backup = import_configuration(path)
            self.settings = load_settings()
            if self.on_settings_saved:
                self.on_settings_saved()
            messagebox.showinfo("Configuration imported", f"Configuration imported successfully.\nBackup: {backup or 'None'}", parent=self)
            self._build_backend_unlocked()
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc), parent=self)

    def _change_passcode(self):
        current = simpledialog.askstring("Change passcode", "Current administrator passcode:", show="•", parent=self)
        if current is None:
            return
        if not verify_admin_passcode(current):
            messagebox.showerror("Passcode", "Current passcode is incorrect.", parent=self)
            return
        new = simpledialog.askstring("Change passcode", "New administrator passcode (minimum 6 characters):", show="•", parent=self)
        if new is None:
            return
        confirm = simpledialog.askstring("Change passcode", "Confirm new administrator passcode:", show="•", parent=self)
        if new != confirm:
            messagebox.showerror("Passcode", "New passcodes do not match.", parent=self)
            return
        try:
            set_admin_passcode(new)
            messagebox.showinfo("Passcode", "Administrator passcode changed.", parent=self)
        except Exception as exc:
            messagebox.showerror("Passcode", str(exc), parent=self)

    def _reset_defaults(self):
        if not messagebox.askyesno(
            "Reset settings",
            "Reset all extractor/formatting/OCR settings to defaults?\n\nThe administrator passcode will be preserved.",
            parent=self,
        ):
            return
        backup = reset_to_defaults(preserve_admin=True)
        self.settings = load_settings()
        if self.on_settings_saved:
            self.on_settings_saved()
        messagebox.showinfo("Settings reset", f"Defaults restored.\nBackup: {backup or 'None'}", parent=self)
        self.destroy()

    def _install_update(self):
        path = filedialog.askopenfilename(
            parent=self,
            title="Select Allison's Data Extractor update package",
            filetypes=[("Allison's Data Extractor update package", "*.zip"), ("ZIP files", "*.zip")],
        )
        if not path:
            return
        try:
            manifest = inspect_update_package(path, self.app_revision)
        except UpdateError as exc:
            messagebox.showerror("Update package", str(exc), parent=self)
            return

        version = manifest.get("version", "")
        proceed = messagebox.askyesno(
            "Install update",
            f"Current revision: {self.app_revision}\nUpdate revision: {version}\n\n"
            "The package will be verified, the application will close, update itself, and restart. Continue?",
            parent=self,
        )
        if not proceed:
            return

        try:
            manifest, payload = stage_update_package(path, self.app_revision)
            launch_staged_update(manifest, payload)
        except Exception as exc:
            messagebox.showerror("Update failed", str(exc), parent=self)
            return

        messagebox.showinfo(
            "Update ready",
            f"Revision {version} has been verified and staged.\n\nThe application will now restart to complete the update.",
            parent=self,
        )
        self.parent.after(150, self.parent.destroy)

"""Settings UI for Allison's Data Extractor."""

from __future__ import annotations

import copy
import json
import os
import tkinter as tk
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
from update_manager import UpdateError, inspect_update_package, launch_staged_update, stage_update_package


class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, app_revision: str, tesseract_test_callback=None, on_settings_saved=None):
        super().__init__(parent)
        self.parent = parent
        self.app_revision = app_revision
        self.tesseract_test_callback = tesseract_test_callback
        self.on_settings_saved = on_settings_saved
        self.admin_unlocked = False
        self._admin_timeout_id = None

        self.title(f"Settings - Allison's Data Extractor {app_revision}")
        self.geometry("850x680")
        self.minsize(760, 590)
        self.transient(parent)

        self.settings = load_settings()

        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)

        self.general_tab = ttk.Frame(self.notebook, padding=14)
        self.format_tab = ttk.Frame(self.notebook, padding=14)
        self.ocr_tab = ttk.Frame(self.notebook, padding=14)
        self.backend_tab = ttk.Frame(self.notebook, padding=14)

        self.notebook.add(self.general_tab, text="General")
        self.notebook.add(self.format_tab, text="Data Formatting")
        self.notebook.add(self.ocr_tab, text="OCR")
        self.notebook.add(self.backend_tab, text="Backend / Administrator")

        self._build_general()
        self._build_formatting()
        self._build_ocr()
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
            settings.setdefault("general", {})["remember_output_file"] = bool(self.remember_output_var.get())
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
        self._admin_timeout_id = self.after(15 * 60 * 1000, self._expire_admin_session)

    def _expire_admin_session(self):
        self._admin_timeout_id = None
        if self.admin_unlocked:
            self.admin_unlocked = False
            self._build_backend_locked()
            messagebox.showinfo(
                "Administrator session locked",
                "Administrator mode was automatically locked after 15 minutes.",
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
        for row_index, name in enumerate(("Template A", "Template B", "Template C")):
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
        ttk.Button(config_box, text="Change Passcode...", command=self._change_passcode).pack(side="left", padx=6)
        ttk.Button(config_box, text="Reset Defaults...", command=self._reset_defaults).pack(side="right")

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

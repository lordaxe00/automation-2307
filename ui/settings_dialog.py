"""
settings_dialog.py - Modal dialog for payor and income settings.
"""
import logging
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from src.config_loader import AppConfig
logger = logging.getLogger(__name__)


class SettingsDialog(ctk.CTkToplevel):
    """Two-tab settings dialog: Payor + Period & Income."""

    def __init__(self, parent, config: AppConfig):
        super().__init__(parent)
        self._config = config
        self.title("Settings")
        self.geometry("560x520")
        self.resizable(False, False)
        self.grab_set()  # modal
        self.lift()
        self.focus_force()

        self._build_ui()
        self._load_values()

    # ── Build UI ─────────────────────────────────────────────────────────

    def _build_ui(self):
        # Create a tabbed dialog for all settings sections.
        self.tabview = ctk.CTkTabview(self, width=540, height=440)
        self.tabview.pack(padx=10, pady=(10, 5), fill="both", expand=True)
        self.tabview.add("Payor Info")
        self.tabview.add("Period & Income")

        self._build_payor_tab(self.tabview.tab("Payor Info"))
        self._build_period_income_tab(self.tabview.tab("Period & Income"))

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)

        ctk.CTkButton(
            btn_frame, text="Save & Close", width=140,
            command=self._save_all
        ).pack(side="right", padx=5)
        ctk.CTkButton(
            btn_frame, text="Cancel", width=100,
            fg_color="gray40", hover_color="gray30",
            command=self.destroy
        ).pack(side="right")

    def _row(self, parent, label: str, row: int) -> ctk.CTkEntry:
        ctk.CTkLabel(parent, text=label, anchor="w").grid(
            row=row, column=0, padx=10, pady=5, sticky="w"
        )
        entry = ctk.CTkEntry(parent, width=300)
        entry.grid(row=row, column=1, padx=10, pady=5, sticky="ew")
        return entry

    def _build_payor_tab(self, tab):
        tab.columnconfigure(1, weight=1)
        self._payor_name = self._row(tab, "Company Name:", 0)
        self._payor_tin = self._row(tab, "TIN:", 1)
        self._payor_address = self._row(tab, "Address:", 2)
        self._payor_zip = self._row(tab, "ZIP Code:", 3)
        self._payor_signatory = self._row(tab, "Signatory:", 4)
        ctk.CTkLabel(
            tab,
            text="Signatory format: Name / Designation / TIN",
            text_color="gray60", font=("", 11)
        ).grid(row=5, column=0, columnspan=2, padx=10, sticky="w")

    def _build_period_income_tab(self, tab):
        tab.columnconfigure(1, weight=1)
        self._period_from = self._row(tab, "Period From (MM/DD/YYYY):", 0)
        self._period_to = self._row(tab, "Period To (MM/DD/YYYY):", 1)

        ctk.CTkLabel(tab, text="Income Description:", anchor="w").grid(
            row=2, column=0, padx=10, pady=5, sticky="nw"
        )
        self._income_desc = ctk.CTkTextbox(tab, width=300, height=100)
        self._income_desc.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        self._atc_code = self._row(tab, "ATC Code:", 3)
        self._tax_rate = self._row(tab, "Tax Rate (e.g. 0.02):", 4)
        ctk.CTkLabel(
            tab,
            text="These settings override default record values when present.",
            text_color="gray60", font=("", 10)
        ).grid(row=5, column=0, columnspan=2, padx=10, sticky="w")

    # ── Load values ───────────────────────────────────────────────────────

    def _load_values(self):
        # Populate UI fields from the current config values.
        p = self._config.payor
        self._payor_name.insert(0, p.get("name", ""))
        self._payor_tin.insert(0, p.get("tin", ""))
        self._payor_address.insert(0, p.get("address", ""))
        self._payor_zip.insert(0, p.get("zip_code", ""))
        self._payor_signatory.insert(0, p.get("signatory", ""))

        d = self._config.defaults
        self._period_from.insert(0, d.get("period_from", ""))
        self._period_to.insert(0, d.get("period_to", ""))
        self._income_desc.delete("0.0", "end")
        self._income_desc.insert("0.0", d.get("income_description", ""))
        self._atc_code.insert(0, d.get("atc_code", ""))
        self._tax_rate.insert(0, str(d.get("tax_rate", 0.02)))

    # ── Actions ───────────────────────────────────────────────────────────

    def _save_all(self):
        try:
            self._config.save_payor_settings(
                name=self._payor_name.get().strip(),
                tin=self._payor_tin.get().strip(),
                address=self._payor_address.get().strip(),
                zip_code=self._payor_zip.get().strip(),
                signatory=self._payor_signatory.get().strip(),
            )
            # Save custom defaults for income/period so they are reused on next run.
            try:
                tax_rate = float(self._tax_rate.get())
            except ValueError:
                tax_rate = 0.02
            self._config.save_defaults(
                atc_code=self._atc_code.get().strip(),
                income_description=self._income_desc.get("0.0", "end").strip(),
                tax_rate=tax_rate,
                period_from=self._period_from.get().strip(),
                period_to=self._period_to.get().strip(),
            )

            messagebox.showinfo("Settings", "Settings saved successfully.", parent=self)
            self.destroy()
        except Exception as exc:
            messagebox.showerror("Error", str(exc), parent=self)


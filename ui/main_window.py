"""
main_window.py - Primary CustomTkinter UI for BIR Form 2307 Generator.
"""
import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import List, Optional

import customtkinter as ctk

from src.batch_processor import BatchProcessor, BatchSummary
from src.config_loader import AppConfig
from src.logger_setup import UILogHandler
from ui.settings_dialog import SettingsDialog

logger = logging.getLogger(__name__)

# ── Theme ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

_ACCENT = "#1F6AA5"
_SUCCESS = "#2ECC71"
_WARNING = "#F39C12"
_ERROR = "#E74C3C"
_BG_DARK = "#1C1C2E"
_BG_PANEL = "#2A2A3E"
_TEXT = "#ECEFF4"


class MainWindow(ctk.CTk):
    """Root application window."""

    def __init__(self, config: AppConfig):
        super().__init__()
        self._config = config
        self._processor: Optional[BatchProcessor] = None
        self._payees: List[str] = []
        self._stop_event = threading.Event()
        self._running = False

        self.title(f"{config.app['title']}  v{config.app['version']}")
        self.geometry("1100x750")
        self.minsize(900, 620)
        self.configure(fg_color=_BG_DARK)

        # Attach UI log handler
        self._ui_log_handler = UILogHandler(self._append_log)
        self._ui_log_handler.setLevel(logging.INFO)
        logging.getLogger().addHandler(self._ui_log_handler)

        self._build_ui()
        self._refresh_payees_async()

    # ── Build UI ─────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── Top toolbar ──────────────────────────────────────────────────
        toolbar = ctk.CTkFrame(self, height=50, fg_color=_BG_PANEL,
                               corner_radius=0)
        toolbar.pack(fill="x", side="top")
        toolbar.pack_propagate(False)

        ctk.CTkLabel(
            toolbar,
            text="🇵🇭  BIR Form 2307 Generator",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=_TEXT,
        ).pack(side="left", padx=16, pady=10)

        ctk.CTkButton(
            toolbar, text="⚙ Settings", width=110,
            command=self._open_settings,
            fg_color="#3A3A5C", hover_color="#4A4A6C",
        ).pack(side="right", padx=8, pady=8)

        ctk.CTkButton(
            toolbar, text="↻ Refresh", width=100,
            command=self._refresh_payees_async,
            fg_color="#3A3A5C", hover_color="#4A4A6C",
        ).pack(side="right", padx=(0, 4), pady=8)

        # ── Main layout ──────────────────────────────────────────────────
        main = ctk.CTkFrame(self, fg_color=_BG_DARK)
        main.pack(fill="both", expand=True, padx=10, pady=8)
        main.columnconfigure(0, weight=1, minsize=310)
        main.columnconfigure(1, weight=2)
        main.rowconfigure(0, weight=1)

        self._build_left_panel(main)
        self._build_right_panel(main)

    # ── Left panel: payee list ────────────────────────────────────────────

    def _build_left_panel(self, parent):
        panel = ctk.CTkFrame(parent, fg_color=_BG_PANEL, corner_radius=10)
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=0)
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(2, weight=1)

        ctk.CTkLabel(
            panel, text="Payees", font=ctk.CTkFont(size=14, weight="bold"),
            text_color=_TEXT,
        ).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")

        # Search bar
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._filter_payees)
        search_frame = ctk.CTkFrame(panel, fg_color="transparent")
        search_frame.grid(row=1, column=0, padx=10, pady=(0, 6), sticky="ew")
        search_frame.columnconfigure(0, weight=1)

        ctk.CTkEntry(
            search_frame,
            textvariable=self._search_var,
            placeholder_text="🔍  Search payee…",
        ).grid(row=0, column=0, sticky="ew")

        # Payee listbox
        list_frame = ctk.CTkFrame(panel, fg_color="transparent")
        list_frame.grid(row=2, column=0, padx=10, pady=(0, 8), sticky="nsew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        self._payee_listbox = tk.Listbox(
            list_frame,
            selectmode="extended",
            activestyle="none",
            bg="#1E1E30", fg=_TEXT,
            selectbackground=_ACCENT,
            selectforeground=_TEXT,
            borderwidth=0, highlightthickness=0,
            font=("Segoe UI", 11),
            relief="flat",
        )
        scrollbar = ctk.CTkScrollbar(list_frame, command=self._payee_listbox.yview)
        self._payee_listbox.configure(yscrollcommand=scrollbar.set)
        self._payee_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self._payee_count_lbl = ctk.CTkLabel(
            panel, text="0 payees", text_color="gray60",
            font=ctk.CTkFont(size=11)
        )
        self._payee_count_lbl.grid(row=3, column=0, padx=12, pady=(0, 4), sticky="w")

        # Select helpers
        btn_row = ctk.CTkFrame(panel, fg_color="transparent")
        btn_row.grid(row=4, column=0, padx=10, pady=(0, 8), sticky="ew")
        btn_row.columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            btn_row, text="Select All", height=28,
            fg_color="#3A3A5C", hover_color="#4A4A6C",
            command=lambda: self._payee_listbox.select_set(0, "end"),
        ).grid(row=0, column=0, padx=(0, 3), sticky="ew")

        ctk.CTkButton(
            btn_row, text="Clear", height=28,
            fg_color="#3A3A5C", hover_color="#4A4A6C",
            command=lambda: self._payee_listbox.selection_clear(0, "end"),
        ).grid(row=0, column=1, padx=(3, 0), sticky="ew")

    # ── Right panel: actions + log ────────────────────────────────────────

    def _build_right_panel(self, parent):
        panel = ctk.CTkFrame(parent, fg_color=_BG_DARK)
        panel.grid(row=0, column=1, sticky="nsew")
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(2, weight=1)

        # ── Output folder ────────────────────────────────────────────────
        folder_frame = ctk.CTkFrame(panel, fg_color=_BG_PANEL, corner_radius=10)
        folder_frame.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        folder_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(
            folder_frame, text="Output Folder:", font=ctk.CTkFont(weight="bold"),
            text_color=_TEXT,
        ).grid(row=0, column=0, padx=12, pady=10, sticky="w")

        self._output_var = tk.StringVar(value=str(self._config.output_base.resolve()))
        ctk.CTkEntry(
            folder_frame, textvariable=self._output_var, state="readonly"
        ).grid(row=0, column=1, padx=6, pady=10, sticky="ew")

        ctk.CTkButton(
            folder_frame, text="Browse", width=80,
            command=self._browse_output,
        ).grid(row=0, column=2, padx=(0, 12), pady=10)

        ctk.CTkButton(
            folder_frame, text="Open Folder", width=100,
            fg_color="#3A3A5C", hover_color="#4A4A6C",
            command=self._open_output_folder,
        ).grid(row=0, column=3, padx=(0, 12), pady=10)

        # ── Action buttons ───────────────────────────────────────────────
        action_frame = ctk.CTkFrame(panel, fg_color=_BG_PANEL, corner_radius=10)
        action_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        action_frame.columnconfigure((0, 1, 2), weight=1)

        self._btn_selected = ctk.CTkButton(
            action_frame,
            text="▶  Generate Selected",
            height=48,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=_ACCENT,
            command=self._generate_selected,
        )
        self._btn_selected.grid(row=0, column=0, padx=10, pady=12, sticky="ew")

        self._btn_all = ctk.CTkButton(
            action_frame,
            text="▶▶  Generate All",
            height=48,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#217A4B",
            hover_color="#1A5E39",
            command=self._generate_all,
        )
        self._btn_all.grid(row=0, column=1, padx=10, pady=12, sticky="ew")

        self._btn_stop = ctk.CTkButton(
            action_frame,
            text="⏹  Stop",
            height=48,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#8B1A1A",
            hover_color="#6B1010",
            command=self._stop_processing,
            state="disabled",
        )
        self._btn_stop.grid(row=0, column=2, padx=10, pady=12, sticky="ew")

        # ── Status + progress ────────────────────────────────────────────
        status_frame = ctk.CTkFrame(panel, fg_color=_BG_PANEL, corner_radius=10)
        status_frame.grid(row=2, column=0, sticky="nsew")
        status_frame.columnconfigure(0, weight=1)
        status_frame.rowconfigure(2, weight=1)

        self._status_lbl = ctk.CTkLabel(
            status_frame,
            text="Ready — select payees and click Generate.",
            text_color="gray70",
            anchor="w",
            font=ctk.CTkFont(size=12),
        )
        self._status_lbl.grid(row=0, column=0, padx=12, pady=(12, 4), sticky="ew")

        self._progress = ctk.CTkProgressBar(status_frame, mode="determinate")
        self._progress.grid(row=1, column=0, padx=12, pady=(0, 8), sticky="ew")
        self._progress.set(0)

        # Log area
        log_header = ctk.CTkFrame(status_frame, fg_color="transparent")
        log_header.grid(row=2, column=0, sticky="new", padx=12)
        log_header.columnconfigure(0, weight=1)
        ctk.CTkLabel(
            log_header, text="Event Log",
            font=ctk.CTkFont(size=13, weight="bold"), text_color=_TEXT
        ).grid(row=0, column=0, sticky="w", pady=(0, 4))
        ctk.CTkButton(
            log_header, text="Clear", width=60, height=24,
            fg_color="#3A3A5C", hover_color="#4A4A6C",
            command=self._clear_log,
        ).grid(row=0, column=1, sticky="e")

        self._log_box = tk.Text(
            status_frame,
            state="disabled",
            bg="#121220", fg=_TEXT,
            font=("Consolas", 10),
            borderwidth=0, highlightthickness=0,
            relief="flat", wrap="word",
        )
        self._log_box.tag_configure("ERROR", foreground=_ERROR)
        self._log_box.tag_configure("WARNING", foreground=_WARNING)
        self._log_box.tag_configure("INFO", foreground="#A8D8EA")
        self._log_box.tag_configure("SUCCESS", foreground=_SUCCESS)
        log_scrollbar = ctk.CTkScrollbar(status_frame, command=self._log_box.yview)
        self._log_box.configure(yscrollcommand=log_scrollbar.set)
        self._log_box.grid(row=3, column=0, padx=(12, 0), pady=(0, 12), sticky="nsew")
        log_scrollbar.grid(row=3, column=1, padx=(0, 6), pady=(0, 12), sticky="ns")
        status_frame.rowconfigure(3, weight=1)

    # ── Payee list management ─────────────────────────────────────────────

    def _refresh_payees_async(self):
        self._set_status("Connecting to database…")
        threading.Thread(target=self._load_payees, daemon=True).start()

    def _load_payees(self):
        try:
            processor = self._get_processor()
            payees = processor.get_payee_list()
            self._payees = payees
            self.after(0, lambda: self._populate_listbox(payees))
            self.after(0, lambda: self._set_status(
                f"Loaded {len(payees)} payees from database."
            ))
        except Exception as exc:
            err_msg = str(exc)
            self.after(0, lambda: self._set_status(
                f"Database error: {err_msg}", error=True
            ))
            self.after(0, lambda: self._append_log(
                f"Failed to load payees: {err_msg}", level="ERROR"
            ))

    def _populate_listbox(self, payees: List[str]):
        self._payee_listbox.delete(0, "end")
        for p in payees:
            self._payee_listbox.insert("end", p)
        self._payee_count_lbl.configure(text=f"{len(payees)} payee(s)")

    def _filter_payees(self, *_):
        term = self._search_var.get().lower()
        filtered = [p for p in self._payees if term in p.lower()]
        self._populate_listbox(filtered)

    def _get_selected_payees(self) -> List[str]:
        indices = self._payee_listbox.curselection()
        return [self._payee_listbox.get(i) for i in indices]

    # ── Actions ───────────────────────────────────────────────────────────

    def _generate_selected(self):
        selected = self._get_selected_payees()
        if not selected:
            messagebox.showwarning(
                "No Selection", "Please select at least one payee.", parent=self
            )
            return
        self._start_generation(selected)

    def _generate_all(self):
        if not self._payees:
            messagebox.showwarning(
                "No Payees", "No payees loaded. Refresh first.", parent=self
            )
            return
        if not messagebox.askyesno(
            "Generate All",
            f"Generate BIR 2307 for all {len(self._payees)} payees?",
            parent=self,
        ):
            return
        self._start_generation(self._payees)

    def _start_generation(self, payees: List[str]):
        if self._running:
            messagebox.showwarning("Busy", "Processing already in progress.", parent=self)
            return
        self._running = True
        self._stop_event.clear()
        self._set_buttons_state("disabled")
        self._progress.set(0)
        self._set_status(f"Starting generation for {len(payees)} payee(s)…")
        self._append_log(f"▶ Starting batch: {len(payees)} payee(s).", level="INFO")

        threading.Thread(
            target=self._run_batch,
            args=(payees,),
            daemon=True,
        ).start()

    def _run_batch(self, payees: List[str]):
        try:
            processor = self._get_processor()
            summary = processor.process_payees(
                payees,
                progress_cb=self._on_progress,
                stop_event=self._stop_event,
            )
            self.after(0, lambda: self._on_batch_done(summary))
        except Exception as exc:
            logger.exception("Batch run error")
            self.after(0, lambda: self._set_status(f"Error: {exc}", error=True))
        finally:
            self.after(0, self._on_batch_finish)

    def _on_progress(self, current: int, total: int, name: str, status: str):
        pct = current / total if total else 0
        self.after(0, lambda: self._progress.set(pct))
        self.after(0, lambda: self._set_status(
            f"[{current}/{total}] {name} — {status}"
        ))
        level = "ERROR" if "✗" in status or "failed" in status.lower() else "INFO"
        self.after(0, lambda: self._append_log(
            f"[{current}/{total}] {name}: {status}", level=level
        ))

    def _on_batch_done(self, summary: BatchSummary):
        self._progress.set(1.0)
        msg = (
            f"✓ Complete — {summary.succeeded} succeeded, "
            f"{summary.failed} failed out of {summary.total}."
        )
        self._set_status(msg)
        self._append_log(msg, level="SUCCESS" if summary.failed == 0 else "WARNING")

        for r in summary.failed_results:
            self._append_log(
                f"  ✗ {r.payee_name}: {'; '.join(r.errors)}", level="ERROR"
            )
        for r in summary.succeeded_results:
            if r.warnings:
                self._append_log(
                    f"  ⚠ {r.payee_name}: {'; '.join(r.warnings)}", level="WARNING"
                )

    def _on_batch_finish(self):
        self._running = False
        self._set_buttons_state("normal")

    def _stop_processing(self):
        self._stop_event.set()
        self._set_status("Stopping after current payee…", error=False)
        self._append_log("⏹ Stop requested by user.", level="WARNING")

    # ── Settings / output ─────────────────────────────────────────────────

    def _open_settings(self):
        SettingsDialog(self, self._config)

    def _browse_output(self):
        folder = filedialog.askdirectory(parent=self, title="Select Output Folder")
        if folder:
            self._output_var.set(folder)
            self._config._app_cfg["output"]["base_folder"] = folder
            self._append_log(f"Output folder changed to: {folder}", level="INFO")

    def _open_output_folder(self):
        folder = Path(self._output_var.get())
        folder.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(str(folder))
        elif sys.platform == "darwin":
            subprocess.run(["open", str(folder)])
        else:
            subprocess.run(["xdg-open", str(folder)])

    # ── Log ───────────────────────────────────────────────────────────────

    def _append_log(self, message: str, level: str = "INFO"):
        def _write():
            self._log_box.configure(state="normal")
            tag = level.upper() if level.upper() in ("ERROR", "WARNING", "INFO") else "INFO"
            if "✓" in message or "succeeded" in message.lower():
                tag = "SUCCESS"
            self._log_box.insert("end", message + "\n", tag)
            self._log_box.see("end")
            self._log_box.configure(state="disabled")
        self.after(0, _write)

    def _clear_log(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    # ── Status bar ────────────────────────────────────────────────────────

    def _set_status(self, text: str, error: bool = False):
        color = _ERROR if error else "gray70"
        self._status_lbl.configure(text=text, text_color=color)

    def _set_buttons_state(self, state: str):
        self._btn_selected.configure(state=state)
        self._btn_all.configure(state=state)
        stop_state = "normal" if state == "disabled" else "disabled"
        self._btn_stop.configure(state=stop_state)

    # ── Processor factory ─────────────────────────────────────────────────

    def _get_processor(self) -> BatchProcessor:
        if self._processor is None:
            self._processor = BatchProcessor(self._config)
        return self._processor

    def on_close(self):
        """Call before destroying the window."""
        logging.getLogger().removeHandler(self._ui_log_handler)
        self.destroy()

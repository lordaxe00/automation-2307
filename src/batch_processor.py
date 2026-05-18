"""
batch_processor.py - Orchestrates end-to-end BIR 2307 generation.
Coordinates: DB query → validation → Excel fill → PDF export.
"""
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Dict, List, Optional

import pandas as pd

from src.config_loader import AppConfig
from src.database import DatabaseManager, DatabaseError
from src.excel_automation import BIR2307ExcelFiller, ExcelAutomationError
from src.pdf_export import ExcelPDFExporter, PDFExportError
from src.validators import (
    PayeeRecord,
    ValidationResult,
    validate_and_build_record,
)

logger = logging.getLogger(__name__)


@dataclass
class ProcessResult:
    """Result of processing a single payee."""
    payee_name: str
    success: bool = False
    excel_path: Optional[Path] = None
    pdf_path: Optional[Path] = None
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)

    def __str__(self) -> str:
        status = "✓" if self.success else "✗"
        return f"{status} {self.payee_name}"


@dataclass
class BatchSummary:
    """Aggregated result of a batch run."""
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    results: List[ProcessResult] = field(default_factory=list)

    @property
    def failed_results(self) -> List[ProcessResult]:
        return [r for r in self.results if not r.success]

    @property
    def succeeded_results(self) -> List[ProcessResult]:
        return [r for r in self.results if r.success]


# Type alias for the progress callback
ProgressCallback = Callable[[int, int, str, str], None]
# signature: (current, total, payee_name, status_message) -> None


class BatchProcessor:
    """
    High-level processor that glues together all subsystems.

    Usage:
        processor = BatchProcessor(config)
        summary = processor.process_payees(["Juan Dela Cruz", "Maria Santos"])
    """

    def __init__(self, config: AppConfig):
        self._config = config
        self._db = DatabaseManager(config)
        self._filler = BIR2307ExcelFiller(config)

    # ── Public API ────────────────────────────────────────────────────────

    def get_payee_list(self) -> List[str]:
        """Retrieve all payee names from the database."""
        try:
            self._db.connect()
            return self._db.get_all_payees()
        except DatabaseError as exc:
            logger.error("Failed to load payee list: %s", exc)
            raise
        finally:
            self._db.disconnect()

    def process_payees(
        self,
        payee_names: List[str],
        progress_cb: Optional[ProgressCallback] = None,
        stop_event=None,
    ) -> BatchSummary:
        """
        Process a list of payees and return a BatchSummary.

        Args:
            payee_names: List of payee names to process.
            progress_cb: Optional callback(current, total, name, status).
            stop_event: Optional threading.Event; processing stops if set.
        """
        summary = BatchSummary(total=len(payee_names))
        excel_paths: List[Path] = []

        # ── Phase 1: Fetch data + fill Excel ─────────────────────────────
        try:
            self._db.connect()

            for i, name in enumerate(payee_names, start=1):
                if stop_event and stop_event.is_set():
                    logger.info("Processing stopped by user.")
                    break

                self._emit(progress_cb, i, len(payee_names), name,
                           f"Querying database for {name}…")

                result = self._process_single(name)
                summary.results.append(result)

                if result.success:
                    summary.succeeded += 1
                    if result.excel_path:
                        excel_paths.append(result.excel_path)
                else:
                    summary.failed += 1

                self._emit(
                    progress_cb, i, len(payee_names), name,
                    "Done ✓" if result.success else f"Failed ✗ — {result.errors}",
                )
        finally:
            self._db.disconnect()

        # ── Phase 2: PDF export (bulk, single Excel instance) ─────────────
        if excel_paths:
            self._export_pdfs(excel_paths, summary, progress_cb, stop_event)

        logger.info(
            "Batch complete — %d succeeded, %d failed.",
            summary.succeeded, summary.failed,
        )
        return summary

    def process_all(
        self,
        progress_cb: Optional[ProgressCallback] = None,
        stop_event=None,
    ) -> BatchSummary:
        """Convenience: process every payee in the database."""
        payees = self.get_payee_list()
        return self.process_payees(payees, progress_cb, stop_event)

    # ── Internal ──────────────────────────────────────────────────────────

    def _process_single(self, payee_name: str) -> ProcessResult:
        result = ProcessResult(payee_name=payee_name)

        try:
            # 1. Fetch data from the database for the current payee
            df = self._db.get_payee_data(payee_name)

            # 2. Validate the query results and build a record for Excel filling
            q_cfg = self._config.query
            d_cfg = self._config.defaults
            record: Optional[PayeeRecord] = validate_and_build_record(
                df=df,
                payee_name=payee_name,
                atc_code_default=d_cfg["atc_code"],
                description_default=d_cfg["income_description"],
                tax_rate_default=d_cfg["tax_rate"],
                period_from=d_cfg.get("period_from"),
                period_to=d_cfg.get("period_to"),
                date_format=d_cfg["date_format"],
                payee_tin_col=q_cfg.get("payee_tin_col"),
                atc_col=q_cfg.get("atc_col"),
                description_col=q_cfg.get("description_col"),
                tax_withheld_col=q_cfg.get("tax_withheld_col"),
            )

            if record is None:
                result.errors.append("Validation failed — check logs for details.")
                return result

            result.warnings = record.validation.warnings

            # 3. Fill Excel
            xlsx_path = self._filler.fill_and_save(record)
            result.excel_path = xlsx_path
            result.success = True

        except DatabaseError as exc:
            result.errors.append(f"Database error: {exc}")
            logger.error("[%s] DB error: %s", payee_name, exc)
        except ExcelAutomationError as exc:
            result.errors.append(f"Excel error: {exc}")
            logger.error("[%s] Excel error: %s", payee_name, exc)
        except Exception as exc:
            result.errors.append(f"Unexpected error: {exc}")
            logger.exception("[%s] Unexpected error", payee_name)

        return result

    def _export_pdfs(
        self,
        excel_paths: List[Path],
        summary: BatchSummary,
        progress_cb: Optional[ProgressCallback],
        stop_event,
    ) -> None:
        total = len(excel_paths)
        path_to_result: Dict[Path, ProcessResult] = {
            r.excel_path: r for r in summary.results if r.excel_path
        }

        with ExcelPDFExporter() as exporter:
            for i, xlsx_path in enumerate(excel_paths, start=1):
                if stop_event and stop_event.is_set():
                    break

                name = xlsx_path.stem
                self._emit(progress_cb, i, total, name,
                           f"Exporting PDF for {name}…")
                try:
                    pdf = exporter.export_to_pdf(xlsx_path)
                    pr = path_to_result.get(xlsx_path)
                    if pr:
                        pr.pdf_path = pdf
                        if not pdf.exists():
                            warning = (
                                "PDF export skipped or failed; "
                                "check Excel output and COM availability."
                            )
                            pr.warnings.append(warning)
                            logger.warning("%s: %s", name, warning)
                except PDFExportError as exc:
                    logger.error("PDF failed for '%s': %s", name, exc)
                    pr = path_to_result.get(xlsx_path)
                    if pr:
                        pr.warnings.append(f"PDF export failed: {exc}")

    @staticmethod
    def _emit(
        cb: Optional[ProgressCallback],
        current: int,
        total: int,
        name: str,
        status: str,
    ) -> None:
        if cb:
            try:
                cb(current, total, name, status)
            except Exception:
                pass

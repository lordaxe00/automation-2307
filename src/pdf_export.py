"""
pdf_export.py - Export a completed BIR 2307 Excel file to PDF.
Uses win32com.client (Microsoft Excel COM) to preserve exact BIR formatting.
"""
import logging
import os
import time
from pathlib import Path

logger = logging.getLogger(__name__)

# win32com is Windows-only; import is deferred so the module loads on all platforms
_win32_available: bool = False


def _ensure_win32_available() -> bool:
    global _win32_available
    if _win32_available:
        return True
    try:
        import win32com.client  # type: ignore
        import pywintypes  # type: ignore
        _win32_available = True
        return True
    except ImportError:
        logger.warning(
            "win32com not available. PDF export will be skipped. "
            "Install pywin32 on Windows to enable PDF generation."
        )
        return False


if _ensure_win32_available():
    import win32com.client  # type: ignore
    import pywintypes  # type: ignore


class PDFExportError(Exception):
    """Raised when PDF export fails."""


class ExcelPDFExporter:
    """
    Exports an Excel workbook to PDF using the Excel COM object.
    Keeps one Excel instance alive per exporter lifecycle for performance.
    """

    # Excel COM constants
    _XL_TYPE_PDF = 0          # xlTypePDF
    _XL_QUALITY_STD = 0       # xlQualityStandard
    _PAGE1_ONLY = True

    def __init__(self):
        self._excel = None

    # ── Lifecycle ─────────────────────────────────────────────────────────

    def open_excel(self) -> None:
        """Launch a hidden Excel COM instance."""
        if not _ensure_win32_available():
            return
        try:
            self._excel = win32com.client.Dispatch("Excel.Application")
            self._excel.Visible = False
            self._excel.DisplayAlerts = False
            self._excel.ScreenUpdating = False
            logger.info("Excel COM instance started.")
        except Exception as exc:
            raise PDFExportError(f"Failed to start Excel: {exc}") from exc

    def close_excel(self) -> None:
        """Quit the Excel COM instance."""
        if self._excel:
            try:
                self._excel.Quit()
            except Exception:
                pass
            self._excel = None
            logger.info("Excel COM instance closed.")

    def __enter__(self):
        self.open_excel()
        return self

    def __exit__(self, *_):
        self.close_excel()

    # ── Export ────────────────────────────────────────────────────────────

    def export_to_pdf(self, xlsx_path: Path) -> Path:
        """
        Export ``xlsx_path`` to a PDF in the same directory.
        Returns the PDF path.

        Falls back gracefully if Excel COM is unavailable.
        """
        pdf_path = xlsx_path.with_suffix(".pdf")

        if not _ensure_win32_available():
            logger.warning(
                "PDF export skipped (win32com unavailable): %s", xlsx_path.name
            )
            return pdf_path

        if not xlsx_path.exists():
            raise PDFExportError(f"Excel file not found: {xlsx_path}")

        if self._excel is None:
            self.open_excel()

        wb = None
        try:
            abs_xlsx = str(xlsx_path.resolve())
            abs_pdf = str(pdf_path.resolve())

            wb = self._excel.Workbooks.Open(abs_xlsx, ReadOnly=True)

            # ExportAsFixedFormat(Type, Filename, Quality,
            #   IncludeDocProperties, IgnorePrintAreas, From, To,
            #   OpenAfterPublish, FixedFormatExtClassPtr)
            wb.ExportAsFixedFormat(
                self._XL_TYPE_PDF,
                abs_pdf,
                self._XL_QUALITY_STD,
                True,   # IncludeDocProperties
                False,  # IgnorePrintAreas
                OpenAfterPublish=False,
            )
            logger.info("PDF exported → %s", pdf_path)
            return pdf_path

        except Exception as exc:
            logger.error("PDF export error for '%s': %s", xlsx_path.name, exc)
            raise PDFExportError(
                f"Failed to export PDF for '{xlsx_path.name}': {exc}"
            ) from exc

        finally:
            if wb:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass

    def batch_export(self, xlsx_paths: list, progress_callback=None) -> dict:
        """
        Export multiple Excel files to PDF.

        Args:
            xlsx_paths: list of Path objects.
            progress_callback: optional callable(current, total, path).

        Returns:
            dict {xlsx_path: pdf_path | None}
        """
        results = {}
        total = len(xlsx_paths)

        for i, path in enumerate(xlsx_paths, start=1):
            path = Path(path)
            try:
                pdf = self.export_to_pdf(path)
                results[path] = pdf
            except PDFExportError as exc:
                logger.error("Batch PDF export failed for '%s': %s", path.name, exc)
                results[path] = None

            if progress_callback:
                try:
                    progress_callback(i, total, path)
                except Exception:
                    pass

            # Brief pause between exports to avoid COM timing issues
            time.sleep(0.3)

        return results

"""
logger_setup.py - Centralized logging configuration
"""
import logging
import logging.handlers
from pathlib import Path


def setup_logging(log_file: str = "logs/bir2307.log",
                  level: str = "INFO") -> None:
    """
    Configure root logger with rotating file handler and console handler.
    Call once at application startup.
    """
    log_path = Path(log_file)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    numeric_level = getattr(logging, level.upper(), logging.INFO)

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(numeric_level)

    # Rotating file handler — max 5 MB, keep 5 backups
    fh = logging.handlers.RotatingFileHandler(
        log_path, maxBytes=5 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    fh.setFormatter(fmt)
    fh.setLevel(numeric_level)

    # Console handler
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    ch.setLevel(logging.WARNING)

    root.addHandler(fh)
    root.addHandler(ch)
    logging.info("Logging initialized → %s", log_path.resolve())


class UILogHandler(logging.Handler):
    """
    Custom handler that forwards log records to a UI callback
    so log messages appear in the CustomTkinter log widget.
    """

    def __init__(self, callback):
        super().__init__()
        self._callback = callback
        self.setFormatter(logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%H:%M:%S",
        ))

    def emit(self, record: logging.LogRecord) -> None:
        msg = self.format(record)
        try:
            self._callback(msg, level=record.levelname)
        except Exception:
            # UI may not be ready or may be on a worker thread.
            # Avoid crashing the application on log output failure.
            pass

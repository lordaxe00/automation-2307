"""
main.py - BIR Form 2307 Generator — Application Entry Point
"""
import sys
import os

# Ensure project root is on the path when running as a frozen exe
if getattr(sys, "frozen", False):
    _BASE = sys._MEIPASS  # type: ignore[attr-defined]
    os.chdir(os.path.dirname(sys.executable))
else:
    _BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _BASE)

from src.config_loader import AppConfig
from src.logger_setup import setup_logging


def main():
    # 1. Bootstrap config & logging
    config = AppConfig()
    setup_logging(
        log_file=config.app.get("log_file", "logs/bir2307.log"),
        level=config.app.get("log_level", "INFO"),
    )

    import logging
    log = logging.getLogger(__name__)
    log.info("=" * 60)
    log.info("BIR 2307 Generator  v%s  starting…", config.app["version"])
    log.info("=" * 60)

    # 2. Validate template exists
    if not config.template_path.exists():
        import tkinter.messagebox as mb
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        mb.showerror(
            "Missing Template",
            f"BIR 2307 template not found:\n{config.template_path}\n\n"
            "Please place 'BIR2307_template.xlsx' in the 'assets' folder.",
        )
        root.destroy()
        sys.exit(1)

    # 3. Launch UI
    from ui.main_window import MainWindow
    app = MainWindow(config)
    app.protocol("WM_DELETE_WINDOW", app.on_close)

    log.info("UI ready.")
    app.mainloop()
    log.info("Application exited.")


if __name__ == "__main__":
    main()

"""
config_loader.py - Centralized configuration management
"""
import json
import logging
import shutil
import sys
from pathlib import Path
from typing import Any, Dict

logger = logging.getLogger(__name__)

if getattr(sys, "frozen", False):
    _STATIC_BASE_DIR = Path(sys._MEIPASS)
    _USER_BASE_DIR = Path(sys.executable).resolve().parent
else:
    _STATIC_BASE_DIR = Path(__file__).resolve().parent.parent
    _USER_BASE_DIR = _STATIC_BASE_DIR

_STATIC_CONFIG_DIR = _STATIC_BASE_DIR / "config"
_USER_CONFIG_DIR = _USER_BASE_DIR / "config"
_STATIC_ASSETS_DIR = _STATIC_BASE_DIR / "assets"
_USER_ASSETS_DIR = _USER_BASE_DIR / "assets"


def _ensure_user_config() -> None:
    _USER_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    for file_name in ("app_config.json", "cell_mapping.json"):
        user_file = _USER_CONFIG_DIR / file_name
        if not user_file.exists():
            shutil.copy2(_STATIC_CONFIG_DIR / file_name, user_file)


def _load_json(filepath: Path) -> Dict[str, Any]:
    """Read JSON from disk and return it as a Python dict."""
    if not filepath.exists():
        raise FileNotFoundError(f"Config file not found: {filepath}")
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_json(filepath: Path, data: Dict[str, Any]) -> None:
    """Write the Python dict back to JSON on disk."""
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


class AppConfig:
    """Singleton-style config holder loaded once at startup."""

    _instance = None
    _app_cfg: Dict = {}
    _cell_map: Dict = {}

    def __new__(cls):
        # Use a simple singleton pattern so config is loaded once and reused.
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._load()
        return cls._instance

    def _load(self):
        if getattr(sys, "frozen", False):
            _ensure_user_config()
        self._app_cfg = _load_json(_USER_CONFIG_DIR / "app_config.json")
        self._cell_map = _load_json(_USER_CONFIG_DIR / "cell_mapping.json")
        logger.info("Configuration loaded successfully.")

    def reload(self):
        self._load()

    # ── Database ──────────────────────────────────────────────────────────
    @property
    def db(self) -> Dict:
        return self._app_cfg["database"]

    @property
    def db_connection_string(self) -> str:
        """Build the ODBC connection string from current database settings."""
        db = self.db
        if db.get("trusted_connection"):
            return (
                f"DRIVER={{{db['driver']}}};"
                f"SERVER={db['server']};"
                f"DATABASE={db['database']};"
                "Trusted_Connection=yes;"
            )
        return (
            f"DRIVER={{{db['driver']}}};"
            f"SERVER={db['server']};"
            f"DATABASE={db['database']};"
            f"UID={db['uid']};"
            f"PWD={db['pwd']};"
        )

    # ── Query ─────────────────────────────────────────────────────────────
    @property
    def query(self) -> Dict:
        return self._app_cfg["query"]

    # ── Payor (static company info) ───────────────────────────────────────
    @property
    def payor(self) -> Dict:
        return self._app_cfg["payor"]
    
    @property
    def signatory(self) -> Dict:
        return self._app_cfg["signatory"]

    # ── Defaults ──────────────────────────────────────────────────────────
    @property
    def defaults(self) -> Dict:
        return self._app_cfg["defaults"]

    # ── Output ────────────────────────────────────────────────────────────
    @property
    def output(self) -> Dict:
        return self._app_cfg["output"]

    @property
    def output_base(self) -> Path:
        return _USER_BASE_DIR / self._app_cfg["output"]["base_folder"]

    # ── Cell mapping ──────────────────────────────────────────────────────
    @property
    def cell_map(self) -> Dict:
        return self._cell_map

    # ── App meta ──────────────────────────────────────────────────────────
    @property
    def app(self) -> Dict:
        return self._app_cfg["app"]

    @property
    def template_path(self) -> Path:
        external = _USER_ASSETS_DIR / "REFORMATTED-2307.xlsx"
        if external.exists():
            return external
        return _STATIC_ASSETS_DIR / "REFORMATTED-2307.xlsx"

    # ── Persistence ───────────────────────────────────────────────────────
    def save_db_settings(self, server: str, database: str, uid: str, pwd: str,
                          trusted: bool = False) -> None:
        self._app_cfg["database"].update({
            "server": server,
            "database": database,
            "uid": uid,
            "pwd": pwd,
            "trusted_connection": trusted,
        })
        _save_json(_USER_CONFIG_DIR / "app_config.json", self._app_cfg)
        logger.info("Database settings saved.")

    def save_payor_settings(self, name: str, tin: str, address: str,
                             zip_code: str, signatory: str) -> None:
        self._app_cfg["payor"].update({
            "name": name,
            "tin": tin,
            "address": address,
            "zip_code": zip_code,
            "signatory": signatory,
        })
        _save_json(_USER_CONFIG_DIR / "app_config.json", self._app_cfg)
        logger.info("Payor settings saved.")

    def save_defaults(
        self,
        atc_code: str,
        income_description: str,
        tax_rate: float,
        period_from: str,
        period_to: str,
    ) -> None:
        """Persist the period and income defaults in the config file."""
        self._app_cfg["defaults"].update({
            "atc_code": atc_code,
            "income_description": income_description,
            "tax_rate": tax_rate,
            "period_from": period_from,
            "period_to": period_to,
        })
        _save_json(_USER_CONFIG_DIR / "app_config.json", self._app_cfg)
        logger.info("Default values saved.")

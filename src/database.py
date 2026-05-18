"""
database.py - SQL Server connection manager and BIR 2307 queries
"""
import logging
from contextlib import contextmanager
from typing import Dict, Generator, List, Optional

import pandas as pd
import pyodbc

from src.config_loader import AppConfig

logger = logging.getLogger(__name__)


class DatabaseError(Exception):
    """Raised for database-related failures."""


class DatabaseManager:
    """
    Manages SQL Server connections and all BIR 2307 queries.
    Uses a single shared connection per instance with auto-reconnect.
    """

    def __init__(self, config: AppConfig):
        self._config = config
        self._conn: Optional[pyodbc.Connection] = None

    # ── Connection management ─────────────────────────────────────────────

    def connect(self) -> None:
        """Open a connection to SQL Server."""
        try:
            conn_str = self._config.db_connection_string
            self._conn = pyodbc.connect(conn_str, timeout=30)
            self._conn.autocommit = False
            logger.info(
                "Connected to SQL Server: %s / %s",
                self._config.db["server"],
                self._config.db["database"],
            )
        except pyodbc.Error as exc:
            raise DatabaseError(f"Connection failed: {exc}") from exc

    def disconnect(self) -> None:
        """Close the connection if open."""
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
            self._conn = None
            logger.info("Disconnected from SQL Server.")

    def test_connection(self) -> bool:
        """Return True if server is reachable, False otherwise."""
        try:
            self.connect()
            with self._conn.cursor() as cur:
                cur.execute("SELECT 1")
            self.disconnect()
            return True
        except Exception as exc:
            logger.error("Connection test failed: %s", exc)
            return False

    def _ensure_connected(self) -> None:
        if self._conn is None:
            self.connect()

    @contextmanager
    def _cursor(self) -> Generator[pyodbc.Cursor, None, None]:
        self._ensure_connected()
        cur = self._conn.cursor()
        try:
            yield cur
        except pyodbc.Error as exc:
            logger.error("SQL error: %s", exc)
            raise DatabaseError(str(exc)) from exc
        finally:
            cur.close()

    # ── Queries ───────────────────────────────────────────────────────────

    def get_all_payees(self) -> List[str]:
        """Return distinct payee names, sorted alphabetically."""
        q = self._config.query
        sql = f"""
            SELECT DISTINCT [PARTICULARS]
            FROM [{q['TBL_BIR']}]
            WHERE [PARTICULARS] IS NOT NULL
              AND LTRIM(RTRIM([PARTICULARS])) <> ''
            ORDER BY [PARTICULARS]
        """
        try:
            with self._cursor() as cur:
                cur.execute(sql)
                rows = cur.fetchall()
            payees = [r[0].strip() for r in rows]
            logger.info("Retrieved %d payees.", len(payees))
            return payees
        except DatabaseError:
            raise
        except Exception as exc:
            raise DatabaseError(f"Failed to fetch payees: {exc}") from exc

    def get_payee_data(self, payee_name: str) -> pd.DataFrame:
        """
        Fetch all monthly records for a single payee.

        Returns a DataFrame with columns:
            PARTICULAR, MONTH, AMOUNT,
            ATC_CODE (optional), DESCRIPTION (optional),
            TAX_WITHHELD (optional)
        """
        q = self._config.query
        optional_cols = self._get_optional_columns(q)
        select_optional = (
            ", ".join(f"[{c}]" for c in optional_cols)
            if optional_cols else ""
        )
        select_optional = f", {select_optional}" if select_optional else ""

        sql = f"""
            SELECT
                [PARTICULARS] AS [PARTICULAR],
                [MONTH],
                [AMOUNT]
                {select_optional}
            FROM [{q['TBL_BIR']}]
            WHERE [PARTICULARS] = ?
            ORDER BY [MONTH]
        """
        try:
            with self._cursor() as cur:
                cur.execute(sql, payee_name)
                cols = [desc[0] for desc in cur.description]
                rows = cur.fetchall()

            df = pd.DataFrame([list(r) for r in rows], columns=cols)
            logger.debug("Fetched %d rows for payee '%s'.", len(df), payee_name)
            return df
        except DatabaseError:
            raise
        except Exception as exc:
            raise DatabaseError(
                f"Failed to fetch data for '{payee_name}': {exc}"
            ) from exc

    def get_all_payees_data(self) -> Dict[str, pd.DataFrame]:
        """
        Fetch data for ALL payees grouped by PAYEE_NAME.
        Returns {payee_name: DataFrame}.
        """
        q = self._config.query
        optional_cols = self._get_optional_columns(q)
        select_optional = (
            ", ".join(f"[{c}]" for c in optional_cols)
            if optional_cols else ""
        )
        select_optional = f", {select_optional}" if select_optional else ""

        sql = f"""
            SELECT
                [PARTICULARS] AS [PARTICULAR],
                [MONTH],
                [AMOUNT]
                {select_optional}
            FROM [{q['TBL_BIR']}]
            WHERE [PARTICULARS] IS NOT NULL
            ORDER BY [PARTICULARS], [MONTH]
        """
        try:
            with self._cursor() as cur:
                cur.execute(sql)
                cols = [desc[0] for desc in cur.description]
                rows = cur.fetchall()

            df = pd.DataFrame([list(r) for r in rows], columns=cols)
            grouped = {
                name: grp.reset_index(drop=True)
                for name, grp in df.groupby("PARTICULAR")
            }
            logger.info("Loaded data for %d payees.", len(grouped))
            return grouped
        except DatabaseError:
            raise
        except Exception as exc:
            raise DatabaseError(f"Failed to fetch all payee data: {exc}") from exc

    def _get_optional_columns(self, q: Dict) -> List[str]:
        """
        Check which optional columns exist in the table
        and return those that are present.
        """
        optional = ["payee_tin_col", "atc_col", "description_col", "tax_withheld_col"]
        present = []
        try:
            with self._cursor() as cur:
                cur.execute(
                    f"SELECT TOP 0 * FROM [{q['TBL_BIR']}]"
                )
                existing = {desc[0].lower() for desc in cur.description}
            for key in optional:
                col = q.get(key, "")
                if col and col.lower() in existing:
                    present.append(col)
        except Exception:
            pass
        return present

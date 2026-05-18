"""
validators.py - BIR 2307 data validation and business logic
"""
import datetime
import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)

MONTH_ORDER = [
    "january", "february", "march",
    "april", "may", "june",
    "july", "august", "september",
    "october", "november", "december",
]

QUARTER_MAP = {
    1: "Q1", 2: "Q1", 3: "Q1",
    4: "Q2", 5: "Q2", 6: "Q2",
    7: "Q3", 8: "Q3", 9: "Q3",
    10: "Q4", 11: "Q4", 12: "Q4",
}

QUARTER_MONTHS = {
    "Q1": [1, 2, 3],
    "Q2": [4, 5, 6],
    "Q3": [7, 8, 9],
    "Q4": [10, 11, 12],
}

MONTH_NAMES = {
    1: "January", 2: "February", 3: "March",
    4: "April", 5: "May", 6: "June",
    7: "July", 8: "August", 9: "September",
    10: "October", 11: "November", 12: "December",
}


@dataclass
class ValidationResult:
    is_valid: bool = True
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def add_error(self, msg: str) -> None:
        self.errors.append(msg)
        self.is_valid = False

    def add_warning(self, msg: str) -> None:
        self.warnings.append(msg)


@dataclass
class PayeeRecord:
    """Parsed, validated data ready for Excel filling."""
    payee_name: str
    payee_tin: Optional[str]
    quarter: str
    year: int
    date_from: str          # MM/DD/YYYY
    date_to: str            # MM/DD/YYYY
    month1_amount: float
    month2_amount: float
    month3_amount: float
    month1_label: str
    month2_label: str
    month3_label: str
    atc_code: str
    income_description: str
    tax_withheld: Optional[float]
    tax_rate: float
    validation: ValidationResult = field(default_factory=ValidationResult)


def parse_month_string(month_str) -> Tuple[Optional[int], Optional[int]]:
    """
    Parse a month value from common formats.
    Accepts dates, timestamps, and strings like 'January 2026', '01/2026', '2026-01', or '2026-01-01'.
    Returns (month_number, year) or (None, None) on failure.
    """
    if pd.isna(month_str):
        return None, None
    if isinstance(month_str, (datetime.date, datetime.datetime)):
        return month_str.month, month_str.year
    s = str(month_str).strip()
    if not s:
        return None, None

    # Try "Month YYYY" or "Mon YYYY"
    parts = s.replace(",", "").split()
    if len(parts) == 2:
        month_part, year_part = parts
        try:
            year = int(year_part)
        except ValueError:
            year_part, month_part = parts
            try:
                year = int(year_part)
            except ValueError:
                return None, None

        month_lower = month_part.lower()
        for i, name in enumerate(MONTH_ORDER, start=1):
            if name.startswith(month_lower) or month_lower.startswith(name[:3]):
                return i, year

    # Try MM/YYYY or YYYY-MM or YYYY/MM
    for sep in ["/", "-"]:
        tokens = s.split(sep)
        if len(tokens) == 2:
            try:
                a, b = int(tokens[0]), int(tokens[1])
                if 1 <= a <= 12 and b > 1000:
                    return a, b
                if 1 <= b <= 12 and a > 1000:
                    return b, a
            except ValueError:
                pass

    # Try full date patterns like YYYY-MM-DD or MM/DD/YYYY
    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
    ):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.month, dt.year
        except ValueError:
            pass

    return None, None


def detect_quarter(month_numbers: List[int]) -> Optional[str]:
    """Determine which quarter the months belong to."""
    if not month_numbers:
        return None
    quarters = {QUARTER_MAP.get(m) for m in month_numbers}
    if len(quarters) == 1:
        return quarters.pop()
    logger.warning("Months span multiple quarters: %s", month_numbers)
    # Use the quarter of the first month
    return QUARTER_MAP.get(month_numbers[0])


def validate_and_build_record(
    df: pd.DataFrame,
    payee_name: str,
    atc_code_default: str,
    description_default: str,
    tax_rate_default: float,
    period_from: Optional[str] = None,
    period_to: Optional[str] = None,
    month_col: str = "MONTH",
    amount_col: str = "AMOUNT",
    date_format: str = "%m/%d/%Y",
    payee_tin_col: Optional[str] = None,
    atc_col: Optional[str] = None,
    description_col: Optional[str] = None,
    tax_withheld_col: Optional[str] = None,
) -> Optional[PayeeRecord]:
    """
    Validate a payee's DataFrame and build a PayeeRecord.
    Returns None if validation fails critically.
    """
    result = ValidationResult()

    if df.empty:
        result.add_error("No records found.")
        logger.error("[%s] %s", payee_name, result.errors[-1])
        return None

    # ── Parse months ──────────────────────────────────────────────────────
    parsed = []
    for _, row in df.iterrows():
        raw_amt = row.get(amount_col)
        if pd.isna(raw_amt) or (isinstance(raw_amt, str) and not raw_amt.strip()):
            continue

        m, y = parse_month_string(row[month_col])
        if m is None:
            result.add_error(f"Cannot parse month: '{row[month_col]}'")
        else:
            parsed.append((m, y, row))

    if not parsed:
        logger.error("[%s] No parseable months.", payee_name)
        return None

    # Sort by (year, month)
    parsed.sort(key=lambda x: (x[1], x[0]))

    # ── Quarter detection ─────────────────────────────────────────────────
    month_nums = [p[0] for p in parsed]
    year_candidates = [p[1] for p in parsed if p[1] and p[1] > 1900]
    year = max(year_candidates) if year_candidates else parsed[0][1]
    quarter = detect_quarter(month_nums)

    if quarter is None:
        result.add_error("Could not determine quarter.")
        return None

    expected_months = QUARTER_MONTHS[quarter]

    # ── Duplicate check ───────────────────────────────────────────────────
    seen: Dict[int, int] = {}
    for m, y, _ in parsed:
        seen[m] = seen.get(m, 0) + 1
    duplicates = [MONTH_NAMES[m] for m, cnt in seen.items() if cnt > 1]
    if duplicates:
        result.add_warning(f"Duplicate months detected: {', '.join(duplicates)}")

    # ── Missing month check ───────────────────────────────────────────────
    present_months = set(seen.keys())
    missing = [MONTH_NAMES[m] for m in expected_months if m not in present_months]
    if missing:
        result.add_warning(f"Missing months: {', '.join(missing)} (will use 0.00)")

    # ── Build monthly amounts dict ────────────────────────────────────────
    month_amounts: Dict[int, float] = {m: 0.0 for m in expected_months}
    atc_code = atc_code_default
    income_desc = description_default
    payee_tin: Optional[str] = None
    tax_withheld_total: Optional[float] = None

    for m_num, y_num, row in parsed:
        if payee_tin_col and payee_tin_col in df.columns and not payee_tin:
            raw_tin = row.get(payee_tin_col)
            if raw_tin is not None and str(raw_tin).strip():
                payee_tin = str(raw_tin).strip()
        if m_num not in expected_months:
            result.add_warning(
                f"Month {MONTH_NAMES.get(m_num, m_num)} is outside {quarter} - skipped."
            )
            continue
        try:
            amt = float(row[amount_col])
        except (ValueError, TypeError):
            result.add_error(f"Invalid amount for {MONTH_NAMES.get(m_num)}: '{row[amount_col]}'")
            amt = 0.0
        if amt < 0:
            result.add_warning(f"Negative amount for {MONTH_NAMES.get(m_num)}: {amt}")
        month_amounts[m_num] = month_amounts.get(m_num, 0.0) + amt

        # Optional columns
        if atc_col and atc_col in df.columns and row.get(atc_col):
            atc_code = str(row[atc_col]).strip()
        if description_col and description_col in df.columns and row.get(description_col):
            income_desc = str(row[description_col]).strip()
        if tax_withheld_col and tax_withheld_col in df.columns:
            tw = row.get(tax_withheld_col)
            if tw is not None and str(tw).strip():
                try:
                    v = float(tw)
                    tax_withheld_total = (tax_withheld_total or 0.0) + v
                except (ValueError, TypeError):
                    pass

    # Assign amounts to month slots 1/2/3
    m1, m2, m3 = expected_months
    # Apply explicit period override from settings if provided.
    if period_from and period_from.strip():
        from_date = period_from.strip()
    else:
        from_month_num = m1
        from_date = f"01/01/{year}" if from_month_num == 1 else _first_day(from_month_num, year)

    if period_to and period_to.strip():
        to_date = period_to.strip()
    else:
        to_month_num = m3
        to_date = _last_day(to_month_num, year)

    record = PayeeRecord(
        payee_name=payee_name,
        payee_tin=payee_tin,
        quarter=quarter,
        year=year,
        date_from=from_date,
        date_to=to_date,
        month1_amount=month_amounts[m1],
        month2_amount=month_amounts[m2],
        month3_amount=month_amounts[m3],
        month1_label=MONTH_NAMES[m1],
        month2_label=MONTH_NAMES[m2],
        month3_label=MONTH_NAMES[m3],
        atc_code=atc_code,
        income_description=income_desc,
        tax_withheld=tax_withheld_total,
        tax_rate=tax_rate_default,
        validation=result,
    )

    if result.errors:
        logger.error("[%s] Validation errors: %s", payee_name, result.errors)
    if result.warnings:
        logger.warning("[%s] Validation warnings: %s", payee_name, result.warnings)
    if not result.is_valid:
        return None

    logger.info(
        "[%s] Record built — %s %d | M1=%.2f M2=%.2f M3=%.2f",
        payee_name, quarter, year,
        record.month1_amount, record.month2_amount, record.month3_amount,
    )
    return record


def _first_day(month: int, year: int) -> str:
    return f"{month:02d}/01/{year}"


def _last_day(month: int, year: int) -> str:
    import calendar
    last = calendar.monthrange(year, month)[1]
    return f"{month:02d}/{last:02d}/{year}"

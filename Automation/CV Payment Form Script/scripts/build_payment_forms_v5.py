"""
Court Visitor Payment form auto-filler (v5)

- Reads Excel at EXCEL_PATH (read-only; never writes back)
- Prompts for month (accepts many formats; see parse_month_input)
- Filters to that month (optionally month-to-date)
- Auto-selects the correct template table (>= 8 rows; header match)
- Finds the actual column indexes from the header row
- Replaces ONLY the placeholder text in each cell (keeps static text like 'C-1-PB-')
- Saves outputs as '9_1Sept2025 Court Visitor Payment.docx', etc., in OUTPUT_DIR

Install once:
  python -m pip install pandas python-docx python-dateutil openpyxl
"""
from __future__ import annotations
import os
import re
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from dateutil.relativedelta import relativedelta
from typing import List, Dict, Any, Optional, Sequence

import pandas as pd
from docx import Document

# ===========================
# CONFIG — EDIT THESE IF NEEDED
# ===========================
EXCEL_PATH = r"C:\Users\may\OneDrive\Guardian Docs\ward_guardian_info.xlsx"  # READ-ONLY
SHEET_NAME = 0                                 # first sheet
DATE_COLUMN = "visitdate"                      # used to select the month
OUTPUT_DIR = r"G:\My Drive\Guardianship files\Payments to review and submit"
TEMPLATE_PATH = r"C:\GuardianAutomation\2025 Court Visitor Payment Invoice - fields.docx"  # your local .docx

# Header bits to identify the right table (case-insensitive contains)
EXPECTED_HEADER_BITS = [
    "cause",                      # Cause No.
    "date appointed",
    "date of court visit",
    "date of court visitor report",
]

# Excel aliases for each logical field
EXCEL_ALIASES: Dict[str, Sequence[str]] = {
    "cause": ("causeno",),
    "date_appointed": ("Dateappointed", "dateappointed"),
    "date_visit": ("visitdate",),
    "date_report": ("datesubmitted", "datesubmited"),
}

# Date formatting on the form
DATE_FORMAT = "%m/%d/%Y"  # e.g., 09/05/2025

# Output file name pattern
FILENAME_PATTERN = "{BILL_MONTH_NUM}_{FORM_NO}{BILL_MONTH_NAME_SHORT}{BILL_YEAR} Court Visitor Payment.docx"
# ===========================


def month_short_name(dt: date) -> str:
    """Return short month name with 'Sept' for September (instead of 'Sep')."""
    s = dt.strftime("%b")
    return "Sept" if s == "Sep" else s


def parse_month_input(raw: str) -> Optional[date]:
    """
    Accepts: 'last', '', 'YYYY-MM', 'YYYY/MM', 'MM-YYYY', 'MM/YYYY', 'M-YYYY', 'M/YYYY',
             'YYYY-M', 'YYYY/M', '9', '09'
    Returns date set to the FIRST of the target month, or None if unparseable.
    """
    raw = (raw or "").strip().lower()
    today = date.today()

    if raw == "":
        return today.replace(day=1)
    if raw == "last":
        return (today.replace(day=1) - relativedelta(months=1))

    if raw.isdigit():
        m = int(raw)
        if 1 <= m <= 12:
            return date(today.year, m, 1)

    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})", raw)
    if m:
        y, mon = int(m.group(1)), int(m.group(2))
        if 1 <= mon <= 12:
            return date(y, mon, 1)
    m = re.fullmatch(r"(\d{1,2})[-/](\d{4})", raw)
    if m:
        mon, y = int(m.group(1)), int(m.group(2))
        if 1 <= mon <= 12:
            return date(y, mon, 1)

    for fmt in ("%Y-%m", "%Y/%m", "%m-%Y", "%m/%Y"):
        try:
            dt = datetime.strptime(raw, fmt).date()
            return dt.replace(day=1)
        except ValueError:
            pass

    return None


def prompt_for_month() -> (date, bool):
    raw = input(
        "Enter month to compile (e.g., 9/2025, 09-2025, 2025-9, 2025/9, YYYY-MM). "
        "Blank = current month, or 'last' = previous month: "
    )
    anchor = parse_month_input(raw)
    if anchor is None:
        print("Could not parse; defaulting to current month.")
        anchor = date.today().replace(day=1)

    today = date.today()
    use_mtd = False
    if anchor.year == today.year and anchor.month == today.month:
        yn = input("Use month-to-date for the current month? [Y/n]: ").strip().lower()
        use_mtd = (yn in ("", "y", "yes"))
    else:
        yn = input("Compile the FULL month (not capped to today)? [Y/n]: ").strip().lower()
        use_mtd = not (yn in ("", "y", "yes"))
    return anchor, use_mtd


@dataclass
class BillingWindow:
    start: date
    end: date  # exclusive

    @staticmethod
    def for_month(anchor: date, to_date: Optional[date] = None) -> "BillingWindow":
        start = anchor.replace(day=1)
        end = (start + relativedelta(months=1))
        if to_date is not None:
            end = min(end, to_date + timedelta(days=1))
        return BillingWindow(start, end)


def ensure_output_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def first_existing_col(df: pd.DataFrame, candidates: Sequence[str]) -> Optional[str]:
    for c in candidates:
        if c in df.columns:
            return c
    lower_map = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return None


def load_rows_for_window(bill_start: date, use_month_to_date: bool) -> pd.DataFrame:
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")

    if DATE_COLUMN not in df.columns:
        raise ValueError(f"Expected date column '{DATE_COLUMN}' not found in Excel. Found: {list(df.columns)}")

    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], errors="coerce").dt.date

    to_date_limit = date.today() if use_month_to_date else None
    window = BillingWindow.for_month(bill_start, to_date=to_date_limit)
    mask = (df[DATE_COLUMN] >= window.start) & (df[DATE_COLUMN] < window.end)
    month_df = df.loc[mask].copy()

    # Sort predictably
    sort_keys = [DATE_COLUMN]
    if "causeno" in month_df.columns:
        sort_keys.append("causeno")
    month_df.sort_values(sort_keys, inplace=True)

    # Resolve Excel headers for each logical field
    resolved: Dict[str, str] = {}
    for logical, aliases in EXCEL_ALIASES.items():
        col = first_existing_col(month_df, aliases)
        if not col:
            raise ValueError(
                f"None of the aliases {aliases} exist in the Excel header. "
                f"Headers present: {list(month_df.columns)}"
            )
        resolved[logical] = col
    month_df.attrs["resolved"] = resolved

    return month_df.reset_index(drop=True)


def find_target_table(doc: Document) -> int:
    """
    Choose the table that:
      - has >= 8 rows (1 header + 7 data rows),
      - has >= 4 columns,
      - header contains EXPECTED_HEADER_BITS (case-insensitive).
    Fallback: the table with the most rows.
    """
    best_idx = None
    best_score = -1
    fallback_idx = None
    fallback_rows = -1

    for idx, t in enumerate(doc.tables):
        row_count = len(t.rows)
        col_count = len(t.columns) if row_count else 0

        if row_count > fallback_rows:
            fallback_rows = row_count
            fallback_idx = idx

        if row_count < 8 or col_count < 4:
            continue

        try:

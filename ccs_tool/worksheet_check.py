"""Compare today's extraction vs existing multi-sheet worksheets."""

from __future__ import annotations

from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import pandas as pd

# Presets matching your existing scripts
WORKSHEET_PRESETS = {
    "LDB": {
        "label": "LDB worksheet check",
        "worksheet_sheets": ["KAMs", "KAMs - To correct", "Pricing", "Credit"],
        "filter_ccs_in_reference": True,
        "hint_extraction": "e.g. LDB12.xlsx (today's extraction)",
        "hint_worksheet": "e.g. CHECK LDB.xlsx (existing worksheet)",
    },
    "CPD": {
        "label": "CPD worksheet check",
        "worksheet_sheets": ["Export", "Cleared", "Pricing", "Cleared new"],
        "filter_ccs_in_reference": True,
        "hint_extraction": "e.g. Check CPD.xlsx or CPD extraction",
        "hint_worksheet": "e.g. CPD07.xlsx (existing worksheet)",
    },
    "LPD": {
        "label": "LPD worksheet check (same sheet pattern as LDB)",
        "worksheet_sheets": ["KAMs", "KAMs - To correct", "Pricing", "Credit"],
        "filter_ccs_in_reference": True,
        "hint_extraction": "today's LPD extraction",
        "hint_worksheet": "existing LPD worksheet",
    },
    "PPD": {
        "label": "PPD worksheet check (same sheet pattern as LDB)",
        "worksheet_sheets": ["KAMs", "KAMs - To correct", "Pricing", "Credit"],
        "filter_ccs_in_reference": True,
        "hint_extraction": "today's PPD extraction",
        "hint_worksheet": "existing PPD worksheet",
    },
}

AMOUNT_ALIASES = ("SUBI $", "Amount (CoCode Crcy)", "Amount")


def list_excel_files(folder: str | Path) -> List[Path]:
    path = Path(folder).expanduser()
    if not path.is_dir():
        raise FileNotFoundError(f"Folder not found: {path}")
    files = sorted(
        [p for p in path.iterdir() if p.suffix.lower() in {".xlsx", ".xls"} and not p.name.startswith("~$")],
        key=lambda p: p.name.lower(),
    )
    return files


def _normalize_amount_column(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure a column named 'SUBI $' exists (rename from Amount aliases if needed)."""
    out = df.copy()
    if "SUBI $" in out.columns:
        return out
    for alias in AMOUNT_ALIASES[1:]:
        if alias in out.columns:
            out = out.rename(columns={alias: "SUBI $"})
            return out
    raise ValueError(
        f"Need amount column one of {AMOUNT_ALIASES}. Found: {list(out.columns)[:20]}..."
    )


def _require_assignment(df: pd.DataFrame, label: str) -> None:
    if "Assignment" not in df.columns:
        raise ValueError(f"{label} missing column 'Assignment'")


def load_worksheet_keys(
    worksheet_path: str | Path,
    sheets: Sequence[str],
) -> Tuple[pd.DataFrame, List[str]]:
    """Load selected sheets; return unique SUBI $|Assignment keys + list of sheets read."""
    path = Path(worksheet_path)
    read_ok: List[str] = []
    parts: List[pd.DataFrame] = []
    errors: List[str] = []

    for sheet in sheets:
        try:
            df = pd.read_excel(path, sheet_name=sheet)
            df = _normalize_amount_column(df)
            _require_assignment(df, f"Sheet '{sheet}'")
            parts.append(df[["SUBI $", "Assignment"]].copy())
            read_ok.append(sheet)
        except Exception as e:
            errors.append(f"{sheet}: {e}")

    if not parts:
        raise ValueError(
            "Could not read any worksheet sheets. "
            + (" | ".join(errors) if errors else "No sheets configured.")
        )

    combined = pd.concat(parts, ignore_index=True)
    combined = combined.drop_duplicates(subset=["SUBI $", "Assignment"])
    return combined, read_ok


def find_not_on_worksheet(
    extraction: pd.DataFrame,
    worksheet_path: str | Path,
    sheets: Sequence[str],
    *,
    filter_ccs_in_reference: bool = True,
) -> Tuple[pd.DataFrame, Dict[str, object]]:
    """
    Rows in today's extraction that are NOT already on the worksheet
    (matched by SUBI $ + Assignment).
    """
    df_a = _normalize_amount_column(extraction)
    _require_assignment(df_a, "Extraction")

    if filter_ccs_in_reference:
        if "Reference" in df_a.columns:
            df_a = df_a[~df_a["Reference"].astype(str).str.contains("CCS", case=False, na=False)]
        # if no Reference column, skip filter silently

    keys_b, sheets_read = load_worksheet_keys(worksheet_path, sheets)

    df_a = df_a.copy()
    df_a["_key"] = df_a["SUBI $"].astype(str) + "|" + df_a["Assignment"].astype(str)
    keys_b = keys_b.copy()
    keys_b["_key"] = keys_b["SUBI $"].astype(str) + "|" + keys_b["Assignment"].astype(str)

    missing = df_a[~df_a["_key"].isin(set(keys_b["_key"]))].drop(columns="_key")

    meta = {
        "extraction_rows_after_filter": len(df_a),
        "worksheet_unique_keys": len(keys_b),
        "sheets_read": sheets_read,
        "not_on_worksheet": len(missing),
    }
    return missing, meta

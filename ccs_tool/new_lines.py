"""Find new lines vs yesterday (CPD and LDB/LPD/PPD)."""

from __future__ import annotations

from typing import Dict, Tuple

import pandas as pd

from .config import AMOUNT_COL, CPD_CUSTOMER_GROUPS, NON_CPD_DIVS


def _require_cols(df: pd.DataFrame, cols: list[str], label: str) -> None:
    missing = [c for c in cols if c not in df.columns]
    if missing:
        raise ValueError(f"{label} missing columns: {missing}")


def _make_key(df: pd.DataFrame) -> pd.Series:
    return df[AMOUNT_COL].astype(str) + "|" + df["Assignment"].astype(str)


def find_new_non_cpd(df_new: pd.DataFrame, df_last: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """New LDB/LPD/PPD rows not in yesterday (Amount|Assignment), Reference without CCS."""
    required = ["Div", AMOUNT_COL, "Assignment", "Reference"]
    _require_cols(df_new, required, "Today")
    _require_cols(df_last, required, "Yesterday")

    filtered = df_new[df_new["Div"].isin(NON_CPD_DIVS)].copy()
    filtered = filtered[~filtered["Reference"].astype(str).str.contains("CCS", case=False, na=False)]

    filtered["_key"] = _make_key(filtered)
    last_keys = set(_make_key(df_last))
    diff = filtered[~filtered["_key"].isin(last_keys)].drop(columns="_key")

    sheets: Dict[str, pd.DataFrame] = {}
    if not diff.empty:
        sheets["All_New"] = diff
        for div in NON_CPD_DIVS:
            part = diff[diff["Div"] == div]
            if not part.empty:
                sheets[div] = part
    return sheets


def find_new_cpd(
    df_new: pd.DataFrame, df_last: pd.DataFrame
) -> Tuple[Dict[str, pd.DataFrame], pd.DataFrame]:
    """New CPD rows not in yesterday (Amount|Assignment), grouped by customer."""
    required = ["Div", AMOUNT_COL, "Assignment", "Reference", "Customer"]
    _require_cols(df_new, required, "Today")
    _require_cols(df_last, [AMOUNT_COL, "Assignment"], "Yesterday")

    filtered = df_new[df_new["Div"] == "CPD"].copy()
    filtered = filtered[~filtered["Reference"].astype(str).str.contains("CCS", case=False, na=False)]

    filtered["_key"] = list(zip(filtered[AMOUNT_COL], filtered["Assignment"]))
    last_keys = set(zip(df_last[AMOUNT_COL], df_last["Assignment"]))
    result = filtered[~filtered["_key"].isin(last_keys)].drop(columns="_key")

    result_clean = result.dropna(subset=["Customer"]).copy()
    result_clean["Customer"] = result_clean["Customer"].astype(str).str.strip()

    sheets: Dict[str, pd.DataFrame] = {}
    unassigned_rows = []

    if not result_clean.empty:
        sheets["All_Results"] = result_clean.drop_duplicates()

    grouped = {name: [] for name in CPD_CUSTOMER_GROUPS}
    for _, row in result_clean.iterrows():
        customer = str(row["Customer"])
        matched = False
        for group, customers in CPD_CUSTOMER_GROUPS.items():
            if customer in customers:
                grouped[group].append(row)
                matched = True
                break
        if not matched:
            unassigned_rows.append(row)

    for group, rows in grouped.items():
        if rows:
            sheets[group] = pd.DataFrame(rows)

    unassigned = pd.DataFrame(unassigned_rows) if unassigned_rows else pd.DataFrame()
    return sheets, unassigned

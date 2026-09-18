"""Check Assignment / Item Text / Div changes on existing Dispute IDs."""

from __future__ import annotations

from typing import Dict

import pandas as pd

from .config import CPD_CUSTOMER_GROUPS, NON_CPD_DIVS


def _normalize_dispute_fields(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Dispute ID"] = (
        out["Dispute ID"].astype(str).str.replace(".0", "", regex=False).str.strip()
    )
    if "Assignment" in out.columns:
        out["Assignment"] = out["Assignment"].astype(str).str.strip()
    if "Item Text" in out.columns:
        out["Item Text"] = out["Item Text"].astype(str).str.strip()
    out = out[out["Dispute ID"].str.lower() != "nan"]
    return out


def find_discrepancies_cpd(df_new: pd.DataFrame, df_last: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """CPD: same Dispute ID with changes in Div, Assignment, or Item Text."""
    required = ["Div", "Dispute ID", "Assignment", "Item Text"]
    for label, df in (("Today", df_new), ("Yesterday", df_last)):
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"{label} missing columns: {missing}")

    new = df_new.copy()
    last = df_last.copy()
    new["Source"] = "Today"
    last["Source"] = "Yesterday"
    combined = pd.concat([new, last], ignore_index=True)

    def _changed(group: pd.DataFrame):
        if len(group) < 2:
            return None
        if (
            group["Div"].nunique() > 1
            or group["Assignment"].nunique() > 1
            or group["Item Text"].nunique() > 1
        ):
            return group
        return None

    differing = combined.groupby("Dispute ID", group_keys=False).apply(_changed)
    sheets: Dict[str, pd.DataFrame] = {}

    if differing is None or differing.empty:
        return sheets

    sheets["All_Differing_Rows"] = differing
    cpd_rows = differing[differing["Div"] == "CPD"]
    if cpd_rows.empty:
        return sheets

    sheets["CPD_All"] = cpd_rows
    if "Customer" in cpd_rows.columns:
        for group_name, customers in CPD_CUSTOMER_GROUPS.items():
            customers = {str(c) for c in customers}
            group_rows = cpd_rows[cpd_rows["Customer"].astype(str).isin(customers)]
            if not group_rows.empty:
                sheets[group_name] = group_rows
    return sheets


def find_discrepancies_non_cpd(
    df_new: pd.DataFrame, df_last: pd.DataFrame
) -> Dict[str, pd.DataFrame]:
    """LDB/LPD/PPD: same Dispute ID with Assignment or Item Text change."""
    required = ["Div", "Dispute ID", "Assignment", "Item Text"]
    for label, df in (("Today", df_new), ("Yesterday", df_last)):
        missing = [c for c in required if c not in df.columns]
        if missing:
            raise ValueError(f"{label} missing columns: {missing}")

    new = _normalize_dispute_fields(df_new)
    last = _normalize_dispute_fields(df_last)

    new_f = new[new["Div"].isin(NON_CPD_DIVS)]
    last_f = last[last["Div"].isin(NON_CPD_DIVS)]
    common_ids = set(new_f["Dispute ID"]).intersection(set(last_f["Dispute ID"]))

    rows = []
    for dispute_id in common_ids:
        new_rows = new_f[new_f["Dispute ID"] == dispute_id]
        last_rows = last_f[last_f["Dispute ID"] == dispute_id]
        for _, new_row in new_rows.iterrows():
            for _, last_row in last_rows.iterrows():
                if (
                    new_row["Assignment"] != last_row["Assignment"]
                    or new_row["Item Text"] != last_row["Item Text"]
                ):
                    n = new_row.to_dict()
                    l = last_row.to_dict()
                    n["Source"] = "Today"
                    l["Source"] = "Yesterday"
                    rows.append(n)
                    rows.append(l)

    disc = pd.DataFrame(rows)
    sheets: Dict[str, pd.DataFrame] = {}
    if disc.empty:
        return sheets

    # Drop duplicate pair noise from cartesian product when possible
    sheets["All_Discrepancies"] = disc.drop_duplicates()
    for div in NON_CPD_DIVS:
        part = sheets["All_Discrepancies"][sheets["All_Discrepancies"]["Div"] == div]
        if not part.empty:
            sheets[div] = part
    return sheets

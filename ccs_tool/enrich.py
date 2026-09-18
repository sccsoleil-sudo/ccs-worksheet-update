"""Step 1: populate Div, Brand, Axe from S4 export."""

from __future__ import annotations

import pandas as pd

from .config import BRAND_AXE_MAPPING, DIV_MAPPING


def _get_brand_axe(row: pd.Series) -> pd.Series:
    item_text = str(row["Item Text"])[:4].upper() if pd.notna(row["Item Text"]) else ""
    div = row["Div"]
    for entry in BRAND_AXE_MAPPING:
        if entry["acronym"] == item_text and entry["div"] == div:
            return pd.Series([entry["brand"], entry["axe"]])
    return pd.Series(["[UNIDENTIFIED]", "*General*"])


def enrich_div_brand_axe(df: pd.DataFrame) -> pd.DataFrame:
    """Add Div / Brand / Axe columns and move them to the front."""
    required = ["Business Area", "Item Text"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    out = df.copy()
    out["Div"] = out["Business Area"].map(DIV_MAPPING).fillna("[UNIDENTIFIED]")
    out[["Brand", "Axe"]] = out.apply(_get_brand_axe, axis=1)

    front = ["Div", "Brand", "Axe"]
    rest = [c for c in out.columns if c not in front]
    return out[front + rest]


def enrichment_summary(df: pd.DataFrame):
    """Quick counts by Div and unidentified brands."""
    by_div = df["Div"].value_counts(dropna=False).rename_axis("Div").reset_index(name="Rows")
    unidentified = int((df["Brand"] == "[UNIDENTIFIED]").sum())
    return by_div, unidentified

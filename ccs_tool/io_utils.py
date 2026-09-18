"""Helpers: Excel I/O and download bytes."""

from __future__ import annotations

from io import BytesIO
from typing import Dict

import pandas as pd


def read_excel(uploaded_or_path) -> pd.DataFrame:
    return pd.read_excel(uploaded_or_path)


def sheets_to_excel_bytes(sheets: Dict[str, pd.DataFrame], empty_message: str) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if not sheets:
            pd.DataFrame([{"Status": empty_message}]).to_excel(
                writer, sheet_name="Summary", index=False
            )
        else:
            for name, df in sheets.items():
                safe = str(name)[:31]
                df.to_excel(writer, sheet_name=safe, index=False)
    buffer.seek(0)
    return buffer.getvalue()

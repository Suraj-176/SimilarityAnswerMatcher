"""File loading utilities shared by comparison flows."""

from __future__ import annotations

import pandas as pd


def read_uploaded_file(uploaded_file, sheet_name=None):
    """Read a Streamlit uploaded file as CSV or Excel.

    Behavior mirrors the reference app:
    - `.csv` files are parsed as CSV with delimiter sniffing and encoding fallbacks.
    - Other files are parsed as Excel first, then fallback to CSV.
    """
    if uploaded_file is None:
        return None

    fname = getattr(uploaded_file, "name", "").lower()

    def try_read_csv_with_encodings(fileobj):
        last_exc = None
        for enc in ("utf-8", "cp1252", "latin1"):
            try:
                fileobj.seek(0)
                return pd.read_csv(fileobj, encoding=enc, sep=None, engine="python")
            except Exception as exc:  # pragma: no cover - passthrough from parser
                last_exc = exc
                continue
        if last_exc is None:
            raise RuntimeError("Failed to read CSV: unknown error")
        raise last_exc

    if fname.endswith(".csv"):
        return try_read_csv_with_encodings(uploaded_file)

    try:
        uploaded_file.seek(0)
        xl = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")
        if isinstance(xl, dict):
            if sheet_name is None:
                return next(iter(xl.values()))
            if sheet_name in xl:
                return xl[sheet_name]
            str_key = str(sheet_name)
            if str_key in xl:
                return xl[str_key]
            return next(iter(xl.values()))
        return xl
    except Exception:
        return try_read_csv_with_encodings(uploaded_file)

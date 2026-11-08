"""
importer.py
===========
Multi-format Data Importer for Excel <-> Website Automation Assistant.
Reads data from:
1. Excel (.xlsx, .xls) - supports multiple sheets
2. CSV (.csv) - handles UTF-8 and Hindi Unicode text
3. Word (.docx) - extracts structured tables AND "Label: Value" paragraphs into a clean DataFrame
"""

from __future__ import annotations
import io
import os
import re
from typing import Dict, List, Optional, Any
import pandas as pd

try:
    import docx
except ImportError:
    docx = None


def _ensure_file_like(file_or_bytes: Any) -> Any:
    if isinstance(file_or_bytes, (bytes, bytearray)):
        return io.BytesIO(file_or_bytes)
    return file_or_bytes


def load_excel_or_csv(file_obj: Any, filename: str) -> Dict[str, pd.DataFrame]:
    """
    Reads an uploaded Excel or CSV file (or raw bytes) and returns a dictionary of {sheet_name: DataFrame}.
    Preserves string formatting for leading-zero PIN codes, mobile numbers, and Aadhaar numbers.
    """
    sheets_dict: Dict[str, pd.DataFrame] = {}
    ext = os.path.splitext(filename)[1].lower()
    stream = _ensure_file_like(file_obj)

    if ext == ".csv":
        df = pd.read_csv(stream, dtype=str, keep_default_na=False)
        clean_name = os.path.splitext(os.path.basename(filename))[0] or "Sheet1"
        sheets_dict[clean_name] = df
    else:
        xls = pd.ExcelFile(stream)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str, keep_default_na=False)
            sheets_dict[sheet_name] = df

    return sheets_dict


def load_spreadsheet_sheets(file_or_bytes: Any, filename: str) -> Dict[str, pd.DataFrame]:
    """Alias used by app.py to load all sheets from Excel or CSV."""
    return load_excel_or_csv(file_or_bytes, filename)


def load_word_document(file_or_bytes: Any, filename: str = "document.docx") -> Dict[str, Any]:
    """
    Extracts structured records from a Word (.docx) file.
    Supports two common government/citizen document formats:
    1. Word Tables: First row is treated as headers, subsequent rows as records.
       (Also handles 2-column Key-Value tables!)
    2. Paragraph "Label: Value" or "Label - Value" blocks:
       e.g.:
       Name: Ramesh Kumar
       Father Name: Suresh Kumar
       Mobile: 9876543210
    Returns a dict with:
      - "tables": List[pd.DataFrame]
      - "key_values": List[Dict[str, str]]
      - "combined_df": pd.DataFrame
    """
    if docx is None:
        raise RuntimeError("python-docx is not installed. Run 'pip install python-docx'.")

    stream = _ensure_file_like(file_or_bytes)
    doc = docx.Document(stream)
    extracted_records: List[Dict[str, str]] = []
    table_dfs: List[pd.DataFrame] = []
    kv_blocks: List[Dict[str, str]] = []

    # 1. Extract from Word Tables
    for table in doc.tables:
        if not table.rows:
            continue

        first_row_cells = [c.text.strip() for c in table.rows[0].cells]
        if (
            len(first_row_cells) == 2
            and any(k in first_row_cells[0].lower() for k in ["name", "नाम", "field", "label", "आवेदक"])
            and not any(k in first_row_cells[1].lower() for k in ["father", "पिता", "mobile", "मोबाइल"])
        ):
            kv_record: Dict[str, str] = {}
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                if len(cells) >= 2 and cells[0]:
                    key = cells[0].rstrip(":-").strip()
                    kv_record[key] = cells[1]
            if kv_record:
                extracted_records.append(kv_record)
                kv_blocks.append(kv_record)
                table_dfs.append(pd.DataFrame([kv_record]))
        else:
            headers = [c.text.strip() or f"Column_{i+1}" for i, c in enumerate(table.rows[0].cells)]
            t_rows: List[Dict[str, str]] = []
            for row in table.rows[1:]:
                vals = [c.text.strip() for c in row.cells]
                if any(vals):
                    rec = {headers[i]: (vals[i] if i < len(vals) else "") for i in range(len(headers))}
                    extracted_records.append(rec)
                    t_rows.append(rec)
            if t_rows:
                table_dfs.append(pd.DataFrame(t_rows))

    # 2. Extract from "Label: Value" Paragraphs
    current_para_record: Dict[str, str] = {}
    kv_pattern = re.compile(r"^([^:\-—]{2,40})\s*[:\-—]\s*(.+)$")

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            if len(current_para_record) >= 3:
                extracted_records.append(current_para_record)
                kv_blocks.append(current_para_record)
                current_para_record = {}
            continue

        if text.startswith("---") or text.startswith("==="):
            if current_para_record:
                extracted_records.append(current_para_record)
                kv_blocks.append(current_para_record)
                current_para_record = {}
            continue

        match = kv_pattern.match(text)
        if match:
            key = match.group(1).strip()
            val = match.group(2).strip()
            if key in current_para_record and len(current_para_record) >= 2:
                extracted_records.append(current_para_record)
                kv_blocks.append(current_para_record)
                current_para_record = {}
            current_para_record[key] = val

    if current_para_record:
        extracted_records.append(current_para_record)
        kv_blocks.append(current_para_record)

    combined_df = pd.DataFrame(extracted_records).fillna("") if extracted_records else pd.DataFrame()
    return {
        "tables": table_dfs,
        "key_values": kv_blocks,
        "combined_df": combined_df,
    }


def filter_rows_by_selection(
    df: pd.DataFrame,
    mode: str = "all",
    start_row: int = 1,
    end_row: Optional[int] = None,
    selected_indices: Optional[List[int]] = None
) -> pd.DataFrame:
    """
    Filters DataFrame rows based on user selection:
    - 'all': all rows
    - 'range': 1-indexed inclusive row range [start_row, end_row]
    - 'selected': specific 0-indexed row indices
    """
    if df.empty:
        return df

    if mode == "range":
        s_idx = max(0, start_row - 1)
        e_idx = end_row if end_row is not None else len(df)
        return df.iloc[s_idx:e_idx].copy()

    elif mode == "selected" and selected_indices is not None:
        valid_indices = [i for i in selected_indices if 0 <= i < len(df)]
        return df.iloc[valid_indices].copy()

    return df.copy()

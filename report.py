"""
report.py
=========
Proof & Final Audit Report Generator (Excel Copy + PDF Report).
Guarantees:
- NEVER overwrites the original uploaded Excel file.
- Writes `Portal_Status`, `Application_Ack_No`, `Submission_Time`, `Error_Message`, and `Screenshot_Proof`
  into a formatted copy of the user's spreadsheet (with frozen top row, header styling, and auto-fit widths).
- Generates a clean PDF Summary Report with success/fail counts and masked PII for official record-keeping.
"""

from __future__ import annotations
import io
from typing import Dict, List, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

from utils import mask_sensitive_value

try:
    from fpdf import FPDF
except ImportError:
    FPDF = None


def export_formatted_excel_bytes(
    df: pd.DataFrame,
    sheet_title: str = "Extracted_Data"
) -> bytes:
    """
    B7: Exports any DataFrame to a professionally styled Excel workbook (.xlsx)
    with dark slate header styling, frozen header row (`A2`), status colour coding,
    and automatic column widths.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title[:31]
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_ok = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_fail = PatternFill(start_color="FECDD3", end_color="FECDD3", fill_type="solid")
    fill_warn = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    cols = list(df.columns)
    status_col_idx = None
    for i, c in enumerate(cols, start=1):
        if "status" in str(c).lower():
            status_col_idx = i
            break

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=1):
        ws.append(row)
        for c_idx in range(1, len(row) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border
            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif status_col_idx and c_idx == status_col_idx:
                val_low = str(cell.value or "").lower()
                if any(k in val_low for k in ["submitted", "approved", "verified", "match"]):
                    cell.fill = fill_ok
                elif any(k in val_low for k in ["failed", "rejected", "error", "missing"]):
                    cell.fill = fill_fail
                elif any(k in val_low for k in ["filled", "pending", "skipped", "mismatch"]):
                    cell.fill = fill_warn

    ws.freeze_panes = "A2"
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=12)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 4, 13), 45)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def build_updated_excel_with_results(
    original_df: pd.DataFrame,
    execution_results: Dict[int, Dict[str, Any]],
    sheet_title: str = "Portal_Results"
) -> bytes:
    """
    Creates a copy of `original_df`, appends the portal outcome columns
    (Portal_Status, Application_Ack_No, Submission_Time, Error_Message, Screenshot_Proof),
    and returns a styled `.xlsx` byte stream with frozen headers and auto column widths.
    """
    df_copy = original_df.copy()

    statuses: List[str] = []
    ack_nos: List[str] = []
    timestamps: List[str] = []
    errors: List[str] = []
    screenshots: List[str] = []

    for idx in range(len(df_copy)):
        res = execution_results.get(idx, execution_results.get(str(idx), {}))
        statuses.append(res.get("status", "Pending"))
        ack_nos.append(res.get("ack_no") or res.get("reference_no", ""))
        timestamps.append(res.get("timestamp", ""))
        errors.append(res.get("error_message", ""))
        screenshots.append(res.get("screenshot", ""))

    df_copy["Portal_Status"] = statuses
    df_copy["Application_Ack_No"] = ack_nos
    df_copy["Submission_Time"] = timestamps
    df_copy["Error_Message"] = errors
    df_copy["Screenshot_Proof"] = screenshots

    return export_formatted_excel_bytes(df_copy, sheet_title=sheet_title)


def export_results_excel(
    original_df: pd.DataFrame,
    outcomes: Dict[int, Dict[str, Any]],
    raw_original_bytes: Optional[bytes] = None,
) -> bytes:
    """Wrapper used by app.py to export non-destructive copy of Excel with outcomes."""
    return build_updated_excel_with_results(original_df, outcomes)


def generate_pdf_audit_report(
    df: Optional[pd.DataFrame],
    execution_results: Dict[int, Dict[str, Any]],
    portal_name: str = "SSO / eMitra Portal"
) -> bytes:
    """
    Generates a clean PDF Batch Execution Summary Report with masked Aadhaar/Mobile numbers,
    total counts (Submitted, Filled, Failed, Skipped), and row-level acknowledgement numbers.
    """
    total = len(df) if df is not None and not df.empty else len(execution_results)
    submitted = sum(1 for r in execution_results.values() if r.get("status") == "Submitted")
    filled = sum(1 for r in execution_results.values() if r.get("status") == "Filled")
    failed = sum(1 for r in execution_results.values() if str(r.get("status", "")).startswith("Fail"))
    skipped = sum(1 for r in execution_results.values() if "Skip" in str(r.get("status", "")))

    if FPDF is None:
        lines = [
            f"EXCEL <-> WEBSITE AUTOMATION AUDIT REPORT ({portal_name})",
            f"Generated At: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}",
            f"Total Rows: {total} | Submitted: {submitted} | Filled: {filled} | Failed: {failed} | Skipped: {skipped}",
            "-" * 75,
        ]
        for idx, res in sorted(execution_results.items(), key=lambda x: int(x[0])):
            ack = res.get("ack_no") or res.get("reference_no") or "-"
            lines.append(
                f"Row {int(idx)+1:03d} | Status: {res.get('status', 'Pending'):<14} | Ack: {ack} | Time: {res.get('timestamp', '-')}"
            )
        return "\n".join(lines).encode("utf-8")

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(0, 10, f"Portal Automation Audit Report - {portal_name}", ln=True)

    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 7, f"Generated At: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True)
    pdf.cell(
        0,
        8,
        f"Total Rows: {total}   |   Submitted: {submitted}   |   Filled: {filled}   |   Failed: {failed}   |   Skipped: {skipped}",
        ln=True,
    )
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(18, 8, "Row #", border=1)
    pdf.cell(48, 8, "Applicant / ID", border=1)
    pdf.cell(35, 8, "Status", border=1)
    pdf.cell(45, 8, "Ack / App No", border=1)
    pdf.cell(42, 8, "Timestamp", border=1, ln=True)

    pdf.set_font("Helvetica", "", 8)
    for idx, res in sorted(execution_results.items(), key=lambda x: int(x[0])):
        row_i = int(idx)
        raw_id = f"Row-{row_i + 1}"
        if df is not None and not df.empty and row_i < len(df):
            first_col = str(df.columns[0])
            raw_id = mask_sensitive_value(df.iloc[row_i][first_col], first_col)
        safe_id = str(raw_id).encode("ascii", "replace").decode("ascii")[:22]
        status_str = str(res.get("status", "Pending"))[:18]
        ack_str = str(res.get("ack_no") or res.get("reference_no") or "-")[:22]
        time_str = str(res.get("timestamp", "-"))[:19]

        pdf.cell(18, 7, str(row_i + 1), border=1)
        pdf.cell(48, 7, safe_id, border=1)
        pdf.cell(35, 7, status_str, border=1)
        pdf.cell(45, 7, ack_str, border=1)
        pdf.cell(42, 7, time_str, border=1, ln=True)

    return bytes(pdf.output())


def generate_pdf_summary_report(
    outcomes: Dict[int, Dict[str, Any]],
    df: Optional[pd.DataFrame] = None,
) -> bytes:
    """Wrapper used by app.py to generate the PDF summary report."""
    return generate_pdf_audit_report(df=df, execution_results=outcomes)

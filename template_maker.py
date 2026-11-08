"""
template_maker.py
=================
STEP 2: Auto-Generate Ready-to-Fill Excel Format (.xlsx) from a Website URL Scan.

Capabilities:
1. Creates a 3-sheet Excel workbook (.xlsx) directly from scanned web form fields:
   - Sheet 1 ("Form_Data_Template"):
     * One column per scanned form field in exact form order.
     * Required columns highlighted with a distinct header & soft tint.
     * Row 1: Clean Column Header ( frozen at row 3 so both Header and Hint row stay visible, or frozen at A2 ).
     * Row 2: Notes / Hints row explaining format, limits, required/optional, and an example value (plus Excel Cell Comments).
     * Excel DataValidation dropdown lists (`openpyxl.worksheet.datavalidation.DataValidation`) attached to cells (rows 3..200) for every dropdown/radio field using the website's actual options from the "Allowed Values" sheet.
     * Extra tracking columns appended at the end:
       `Photo_Document_Filename` (if the form has file upload fields),
       `Status`, `Application_Acknowledgement_No`, `Submission_Time`, `Error_Message`.
   - Sheet 2 ("Instructions"):
     * Step-by-step instructions in simple Hindi and English explaining how to fill data, format dates, prepare photos, and run the assistant.
   - Sheet 3 ("Allowed Values"):
     * Lists every dropdown/radio column and all valid option values scraped from the website.
2. Reshapes an existing user Excel file ("I already have my own Excel") into the generated website template format using the column mapping.
"""

from __future__ import annotations
import io
import re
from typing import Dict, List, Any, Optional
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.worksheet.datavalidation import DataValidation


def make_clean_column_name(field: Dict[str, Any], index: int) -> str:
    """
    Produces a clean, human-friendly Excel column header from a scanned form field.
    Prefers Gemini's `suggested_column` if present, otherwise cleans the label/name/id.
    """
    if field.get("suggested_column"):
        return str(field["suggested_column"]).strip()

    raw_label = str(field.get("label") or field.get("name") or field.get("id") or f"Field_{index}").strip()
    # Remove trailing asterisks and colons
    cleaned = re.sub(r"[\*\:\-]+$", "", raw_label).strip()
    if not cleaned:
        cleaned = f"Field_{index}"
    return cleaned


def build_column_hint_text(field: Dict[str, Any]) -> str:
    """Builds a concise Hindi + English hint string for the template's hint row & cell comment."""
    f_type = str(field.get("type", "text")).lower()
    req = bool(field.get("required", False))
    max_len = field.get("maxlength") or field.get("max_length")
    options = field.get("options") or []
    validation_type = str(field.get("validation_rule") or "").lower()

    parts: List[str] = ["Required (अनिवार्य)" if req else "Optional (वैकल्पिक)"]

    if f_type == "file":
        parts.append("Filename in photos folder (e.g. 101_Ramesh.jpg, <= 50KB)")
    elif options:
        preview_opts = ", ".join(str(o) for o in options[:4] if not str(o).startswith("--"))
        parts.append(f"Dropdown [{preview_opts}]")
    elif "mobile" in validation_type or "mobile" in str(field.get("name", "")).lower():
        parts.append("10-digit Mobile (e.g. 9876543210)")
    elif "aadhaar" in validation_type or "aadhaar" in str(field.get("name", "")).lower():
        parts.append("12-digit Aadhaar (e.g. 456789012345)")
    elif "pin" in validation_type or "pin" in str(field.get("name", "")).lower():
        parts.append("6-digit PIN Code (e.g. 302001)")
    elif f_type == "date" or "dob" in str(field.get("name", "")).lower() or "date" in validation_type:
        parts.append("Date DD/MM/YYYY (e.g. 15/08/1992)")
    elif f_type == "checkbox":
        parts.append("Yes / No (हाँ / नहीं)")
    elif f_type == "number":
        parts.append("Numeric value (केवल संख्या)")
    else:
        parts.append("Text (टेक्स्ट)")

    if max_len:
        parts.append(f"Max {max_len} chars")

    return " | ".join(parts)


def generate_excel_template_from_scan(
    scanned_fields: List[Dict[str, Any]],
    portal_url: str = "",
    template_name: str = "Portal_Form_Template",
    include_sample_hint_row: bool = True,
) -> bytes:
    """
    Generates a ready-to-fill 3-sheet `.xlsx` Excel template from `scanned_fields`:
    - Sheet 1: `Form_Data_Template` (Headers, Hint comments, DataValidation dropdowns, Tracking columns)
    - Sheet 2: `Instructions` (Hindi & English step-by-step instructions)
    - Sheet 3: `Allowed Values` (Dropdown & Radio option lists referenced by Excel DataValidation)
    """
    wb = openpyxl.Workbook()

    # -----------------------------------------------------------------------
    # SHEET 1: Form_Data_Template
    # -----------------------------------------------------------------------
    ws_data = wb.active
    ws_data.title = "Form_Data_Template"
    ws_data.views.sheetView[0].showGridLines = True

    # -----------------------------------------------------------------------
    # SHEET 2: Instructions (Hindi + English)
    # -----------------------------------------------------------------------
    ws_inst = wb.create_sheet("Instructions")
    ws_inst.views.sheetView[0].showGridLines = True

    # -----------------------------------------------------------------------
    # SHEET 3: Allowed Values (For Dropdowns & Radios)
    # -----------------------------------------------------------------------
    ws_allowed = wb.create_sheet("Allowed Values")
    ws_allowed.views.sheetView[0].showGridLines = True

    # Styling definitions
    req_header_fill = PatternFill(start_color="7C2D12", end_color="7C2D12", fill_type="solid")   # Deep Amber/Rust for Required
    opt_header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")   # Dark Slate for Optional
    meta_header_fill = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")  # Deep Emerald for Status/Output
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    hint_fill = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    hint_font = Font(name="Calibri", size=9, italic=True, color="713F12")

    req_cell_tint = PatternFill(start_color="FFFBEB", end_color="FFFBEB", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    has_file_upload = any(str(f.get("type", "")).lower() == "file" for f in scanned_fields)

    # Build ordered column definitions from scanned fields
    columns_meta: List[Dict[str, Any]] = []
    used_names: Dict[str, int] = {}

    for idx, field in enumerate(scanned_fields, start=1):
        base_col = make_clean_column_name(field, idx)
        if base_col in used_names:
            used_names[base_col] += 1
            col_title = f"{base_col}_{used_names[base_col]}"
        else:
            used_names[base_col] = 1
            col_title = base_col

        raw_opts = [
            str(o).strip()
            for o in (field.get("options") or [])
            if str(o).strip() and not str(o).strip().startswith("--")
        ]

        columns_meta.append({
            "col_title": col_title,
            "required": bool(field.get("required", False)),
            "hint": build_column_hint_text(field),
            "options": raw_opts,
            "is_system": False,
            "field": field,
        })

    # Append extra columns requested by user specification
    if has_file_upload and not any("photo" in c["col_title"].lower() or "file" in c["col_title"].lower() for c in columns_meta):
        columns_meta.append({
            "col_title": "Photo_Document_Filename",
            "required": False,
            "hint": "File name inside photos folder (e.g. 101_Ramesh_Kumar.jpg)",
            "options": [],
            "is_system": True,
            "field": {},
        })

    for sys_col, sys_hint in [
        ("Status", "Auto-filled by Assistant (Pending / Filled / Submitted / Failed / Skipped)"),
        ("Application_Acknowledgement_No", "Auto-filled by Assistant after portal submission"),
        ("Submission_Time", "Auto-filled timestamp (YYYY-MM-DD HH:MM:SS)"),
        ("Error_Message", "Auto-filled error or portal message if any"),
    ]:
        columns_meta.append({
            "col_title": sys_col,
            "required": False,
            "hint": sys_hint,
            "options": [],
            "is_system": True,
            "field": {},
        })

    # Populate Sheet 1 (Form_Data_Template) & Sheet 3 (Allowed Values)
    allowed_col_counter = 0

    for c_idx, col_info in enumerate(columns_meta, start=1):
        col_letter = get_column_letter(c_idx)
        title = col_info["col_title"]
        is_req = col_info["required"]
        is_sys = col_info["is_system"]
        hint_text = col_info["hint"]
        opts = col_info["options"]

        # Row 1: Column Header
        cell_h = ws_data.cell(row=1, column=c_idx, value=title)
        cell_h.font = header_font
        cell_h.fill = meta_header_fill if is_sys else (req_header_fill if is_req else opt_header_fill)
        cell_h.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_h.border = thin_border
        cell_h.comment = Comment(f"{title}\n{hint_text}", "URL-First Automation Assistant")

        # Row 2: Format / Limit / Example Hint Row (Optional or kept as Row 2)
        if include_sample_hint_row:
            cell_hint = ws_data.cell(row=2, column=c_idx, value=hint_text)
            cell_hint.font = hint_font
            cell_hint.fill = hint_fill
            cell_hint.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell_hint.border = thin_border

        # Style empty input cells (Rows 3..52) so required columns stand out
        start_data_row = 3 if include_sample_hint_row else 2
        for r_idx in range(start_data_row, start_data_row + 50):
            c_cell = ws_data.cell(row=r_idx, column=c_idx)
            c_cell.border = thin_border
            if is_req and not is_sys:
                c_cell.fill = req_cell_tint

        # Attach Excel DataValidation dropdown if field has options
        if opts:
            allowed_col_counter += 1
            allow_letter = get_column_letter(allowed_col_counter)
            ah_cell = ws_allowed.cell(row=1, column=allowed_col_counter, value=title)
            ah_cell.font = header_font
            ah_cell.fill = opt_header_fill
            ah_cell.border = thin_border

            for opt_i, opt_val in enumerate(opts, start=2):
                oc = ws_allowed.cell(row=opt_i, column=allowed_col_counter, value=opt_val)
                oc.border = thin_border

            ws_allowed.column_dimensions[allow_letter].width = max(len(title) + 5, 22)

            # Reference 'Allowed Values' sheet range in DataValidation
            last_opt_row = len(opts) + 1
            formula = f"='Allowed Values'!${allow_letter}$2:${allow_letter}${last_opt_row}"
            dv = DataValidation(
                type="list",
                formula1=formula,
                allow_blank=not is_req,
                showDropDown=False,
                showErrorMessage=True,
                errorTitle="Invalid Option / गलत विकल्प",
                error="Please select a valid value from the dropdown list (कृपया ड्रॉपडाउन सूची से सही विकल्प चुनें).",
            )
            ws_data.add_data_validation(dv)
            dv.add(f"{col_letter}{start_data_row}:{col_letter}200")

        # Auto column width
        ws_data.column_dimensions[col_letter].width = min(max(len(title) + 6, 18), 42)

    ws_data.row_dimensions[1].height = 28
    if include_sample_hint_row:
        ws_data.row_dimensions[2].height = 24
        ws_data.freeze_panes = "A3"
    else:
        ws_data.freeze_panes = "A2"

    if allowed_col_counter == 0:
        ws_allowed.cell(row=1, column=1, value="No dropdown/radio fields with fixed options were found on this page.")

    # Populate Sheet 2: Instructions (Bilingual Hindi + English)
    instructions_rows = [
        ["📋 URL-First Excel <-> Website Automation Assistant — Instructions / निर्देश"],
        [f"Target Website URL: {portal_url or 'Scanned Page'}"],
        [f"Template Name: {template_name}"],
        [""],
        ["🇮🇳 हिंदी निर्देश (Step-by-Step Hindi Instructions):"],
        ["1. 'Form_Data_Template' शीट में पंक्ति 3 (Row 3) से अपना डेटा भरना शुरू करें (पंक्ति 1 हेडर है और पंक्ति 2 में नियम/उदाहरण लिखे हैं)।"],
        ["2. जिन कॉलम का हेडर भूरा/लाल (Dark Amber) है और सेल हल्के पीले हैं, वे अनिवार्य (Required *) फील्ड्स हैं—उन्हें खाली न छोड़ें।"],
        ["3. ड्रॉपडाउन वाले कॉलम (जैसे Gender, Category, District) में क्लिक करने पर वेबसाइट के असली विकल्पों की लिस्ट खुलेगी—उसी में से चुनें।"],
        ["4. मोबाइल नंबर पूरे 10 अंक का, आधार 12 अंक का, पिन कोड 6 अंक का और जन्म तिथि DD/MM/YYYY फॉर्मेट (जैसे 15/08/1992) में लिखें।"],
        ["5. फोटो/दस्तावेज़ अपलोड के लिए अपनी फोटो फाइल का नाम (जैसे 101_Ramesh_Kumar.jpg) संबंधित कॉलम में लिखें; ऐप उसे अपने-आप 50KB से कम साइज में कंप्रेस करके अपलोड कर देगा।"],
        ["6. आखिरी 4 कॉलम (Status, Application_Acknowledgement_No, Submission_Time, Error_Message) खाली छोड़ दें—ऐप फॉर्म भरने के बाद इन्हें अपने-आप भरेगा।"],
        [""],
        ["🇬🇧 English Instructions:"],
        ["1. Enter your applicant records in the 'Form_Data_Template' sheet starting from Row 3 (Row 1 = Headers, Row 2 = Field Format Hints)."],
        ["2. Highlighted columns are Required (*) on the website. Do not leave them blank."],
        ["3. Columns with dropdown lists are pre-configured with Excel Data Validation matching the website's exact options (see 'Allowed Values' sheet)."],
        ["4. Enter Mobile as 10 digits, Aadhaar as 12 digits, PIN code as 6 digits, and Dates as DD/MM/YYYY."],
        ["5. Leave Status, Application_Acknowledgement_No, Submission_Time, and Error_Message blank — the assistant populates them automatically in a copy of your file."],
    ]
    for r_i, r_val in enumerate(instructions_rows, start=1):
        ws_inst.append(r_val)
        c = ws_inst.cell(row=r_i, column=1)
        if r_i == 1:
            c.font = Font(name="Calibri", size=14, bold=True, color="0F172A")
        elif r_val and (r_val[0].startswith("🇮🇳") or r_val[0].startswith("🇬🇧")):
            c.font = Font(name="Calibri", size=11, bold=True, color="065F46")
        else:
            c.font = Font(name="Calibri", size=10)
    ws_inst.column_dimensions["A"].width = 115

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def reshape_user_excel_to_template(
    user_df: pd.DataFrame,
    scanned_fields: List[Dict[str, Any]],
    mappings: List[Dict[str, Any]],
) -> pd.DataFrame:
    """
    Takes a user's existing DataFrame (`user_df`) and the column-to-field `mappings`
    and reshapes it into a clean DataFrame whose columns match the scanned website form order.
    """
    selector_to_user_col: Dict[str, str] = {}
    for m in mappings:
        sel = str(m.get("field_selector", "")).strip()
        u_col = str(m.get("data_column", "")).strip()
        if sel and u_col:
            selector_to_user_col[sel] = u_col

    reshaped_data: Dict[str, List[Any]] = {}
    num_rows = len(user_df)

    for idx, field in enumerate(scanned_fields, start=1):
        col_title = make_clean_column_name(field, idx)
        sel = str(field.get("selector", "")).strip()
        mapped_user_col = selector_to_user_col.get(sel, "")

        if mapped_user_col and mapped_user_col in user_df.columns:
            reshaped_data[col_title] = user_df[mapped_user_col].tolist()
        elif col_title in user_df.columns:
            reshaped_data[col_title] = user_df[col_title].tolist()
        else:
            reshaped_data[col_title] = [""] * num_rows

    reshaped_df = pd.DataFrame(reshaped_data)
    for sys_col in ["Status", "Application_Acknowledgement_No", "Submission_Time", "Error_Message"]:
        reshaped_df[sys_col] = ["Pending" if sys_col == "Status" else ""] * num_rows

    return reshaped_df

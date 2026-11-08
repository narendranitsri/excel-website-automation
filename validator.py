"""
validator.py
============
Pre-Flight Citizen & Portal Data Validator.
Validates rows before portal submission:
- Mandatory fields (non-blank check)
- Mobile number (10 digits, starts with 6-9)
- Aadhaar / JanAadhaar number (12 digits)
- PIN code (6 digits)
- Date format validation (DD/MM/YYYY or YYYY-MM-DD)
- Dropdown values against website's allowed option list
- Duplicate rows detection
Provides clear Hindi, Hinglish, and English error messages per row
(e.g. "Row 12: mobile number 9 digit ka hai, 10 hona chahiye").
"""

from __future__ import annotations
import re
from typing import Dict, List, Any, Optional
import pandas as pd


class ValidationIssue:
    def __init__(
        self,
        row_index: int,
        column: str,
        issue_type: str,
        message_en: str,
        message_hi: str,
        current_value: Any,
        severity: str = "error",
    ):
        self.row_index = row_index  # 0-based row index (row_index + 1 is 1-based Excel row)
        self.column = column
        self.issue_type = issue_type
        self.message_en = message_en
        self.message_hi = message_hi
        self.current_value = current_value
        self.severity = severity

    def to_dict(self) -> Dict[str, Any]:
        return {
            "row": self.row_index + 1,
            "row_index": self.row_index,
            "column": self.column,
            "type": self.issue_type,
            "severity": self.severity,
            "value": str(self.current_value),
            "message_en": self.message_en,
            "message_hi": self.message_hi,
        }


def detect_column_role(col_name: str) -> str:
    """Infers semantic validation role from column header (English, Hindi, Hinglish)."""
    c = col_name.strip().lower()
    if any(k in c for k in ["mobile", "phone", "contact", "मोबाइल", "फ़ोन", "फोन"]):
        return "mobile"
    if any(k in c for k in ["aadhaar", "aadhar", "uid", "आधार"]):
        return "aadhaar"
    if any(k in c for k in ["pincode", "pin code", "pin_code", "pin", "पिन", "postal"]):
        return "pincode"
    if any(k in c for k in ["dob", "birth", "date", "तिथि", "जन्म"]):
        return "date"
    if any(k in c for k in ["name", "applicant", "नाम", "आवेदक"]):
        return "mandatory_name"
    return "general"


def validate_dataset(
    df: pd.DataFrame,
    mandatory_columns: Optional[List[str]] = None,
    allowed_dropdown_options: Optional[Dict[str, List[str]]] = None,
) -> List[ValidationIssue]:
    """
    Validates all rows in `df` and returns a list of `ValidationIssue` objects.
    Includes friendly Hinglish/Hindi messages such as:
    "Row 12: mobile number 9 digit ka hai, 10 hona chahiye"
    Also validates dropdown columns against `allowed_dropdown_options` if provided.
    """
    issues: List[ValidationIssue] = []
    if df is None or df.empty:
        return issues

    inferred_mandatory = mandatory_columns if mandatory_columns is not None else [
        col for col in df.columns if detect_column_role(str(col)) in ["mandatory_name", "mobile", "aadhaar"]
    ]
    dropdown_rules = allowed_dropdown_options or {}

    # 1. Check Duplicate Rows
    dup_mask = df.duplicated(keep="first")
    for idx, is_dup in enumerate(dup_mask):
        if is_dup:
            row_num = idx + 1
            first_col = str(df.columns[0])
            issues.append(
                ValidationIssue(
                    row_index=idx,
                    column=first_col,
                    issue_type="duplicate",
                    message_en=f"Row {row_num}: Duplicate row detected.",
                    message_hi=f"Row {row_num}: यह पंक्ति डुप्लिकेट (दोहराई गई) है / duplicate row hai.",
                    current_value=df.iloc[idx][first_col],
                    severity="warning",
                )
            )

    # 2. Row-by-Row Field Validation
    for idx in range(len(df)):
        row_num = idx + 1
        row_data = df.iloc[idx]

        for col in df.columns:
            col_str = str(col)
            raw_val = row_data[col]
            val_str = str(raw_val).strip() if pd.notna(raw_val) else ""
            if val_str.lower() in ["nan", "none", "null"]:
                val_str = ""

            if re.match(r"^\d+\.0$", val_str):
                val_str = val_str[:-2]

            role = detect_column_role(col_str)

            # Check Mandatory Blank
            if col_str in inferred_mandatory and not val_str:
                issues.append(
                    ValidationIssue(
                        row_index=idx,
                        column=col_str,
                        issue_type="blank_mandatory",
                        message_en=f"Row {row_num}: Mandatory column '{col_str}' is blank.",
                        message_hi=f"Row {row_num}: ज़रूरी कॉलम '{col_str}' खाली है (blank hai).",
                        current_value="",
                        severity="error",
                    )
                )
                continue

            if not val_str:
                continue

            # Check Dropdown Allowed Options (if scanned options exist for this column)
            if col_str in dropdown_rules and dropdown_rules[col_str]:
                valid_opts = dropdown_rules[col_str]
                lower_map = {o.strip().lower(): o for o in valid_opts}
                if val_str.lower() not in lower_map:
                    issues.append(
                        ValidationIssue(
                            row_index=idx,
                            column=col_str,
                            issue_type="invalid_dropdown_option",
                            message_en=f"Row {row_num}: Value '{val_str}' in '{col_str}' is not in website allowed options ({', '.join(valid_opts[:5])}).",
                            message_hi=f"Row {row_num}: '{col_str}' में '{val_str}' वेबसाइट की मान्य सूची में नहीं है ({', '.join(valid_opts[:4])})।",
                            current_value=val_str,
                            severity="error",
                        )
                    )

            # Mobile Validation (10 digits)
            if role == "mobile":
                digits = re.sub(r"\D", "", val_str)
                if digits.startswith("91") and len(digits) == 12:
                    digits = digits[2:]
                if len(digits) != 10 or digits[0] not in "6789":
                    issues.append(
                        ValidationIssue(
                            row_index=idx,
                            column=col_str,
                            issue_type="invalid_mobile",
                            message_en=f"Row {row_num}: Mobile number has {len(digits)} digits ('{val_str}'), must be 10 digits.",
                            message_hi=f"Row {row_num}: mobile number {len(digits)} digit ka hai, 10 hona chahiye (मोबाइल नंबर 10 अंक का होना चाहिए).",
                            current_value=val_str,
                            severity="error",
                        )
                    )

            # Aadhaar Validation (12 digits)
            elif role == "aadhaar":
                digits = re.sub(r"\D", "", val_str)
                if len(digits) != 12:
                    issues.append(
                        ValidationIssue(
                            row_index=idx,
                            column=col_str,
                            issue_type="invalid_aadhaar",
                            message_en=f"Row {row_num}: Aadhaar number has {len(digits)} digits, must be 12 digits.",
                            message_hi=f"Row {row_num}: Aadhaar number {len(digits)} digit ka hai, 12 hona chahiye (आधार नंबर 12 अंक का होना चाहिए).",
                            current_value=val_str,
                            severity="error",
                        )
                    )

            # PIN Code Validation (6 digits)
            elif role == "pincode":
                digits = re.sub(r"\D", "", val_str)
                if len(digits) != 6 or digits[0] == "0":
                    issues.append(
                        ValidationIssue(
                            row_index=idx,
                            column=col_str,
                            issue_type="invalid_pincode",
                            message_en=f"Row {row_num}: PIN code '{val_str}' has {len(digits)} digits, must be 6 digits.",
                            message_hi=f"Row {row_num}: PIN code {len(digits)} digit ka hai, 6 hona chahiye (पिन कोड 6 अंक का होना चाहिए).",
                            current_value=val_str,
                            severity="error",
                        )
                    )

            # Date Validation
            elif role == "date":
                dt = pd.to_datetime(val_str, dayfirst=True, errors="coerce")
                if pd.isna(dt):
                    issues.append(
                        ValidationIssue(
                            row_index=idx,
                            column=col_str,
                            issue_type="invalid_date",
                            message_en=f"Row {row_num}: Date '{val_str}' is not in a valid DD/MM/YYYY format.",
                            message_hi=f"Row {row_num}: तारीख '{val_str}' सही फॉर्मेट (DD/MM/YYYY) में नहीं है।",
                            current_value=val_str,
                            severity="error",
                        )
                    )

    return issues


def validate_dataframe(
    df: pd.DataFrame,
    mandatory_columns: Optional[List[str]] = None,
    allowed_dropdown_options: Optional[Dict[str, List[str]]] = None,
) -> Dict[str, Any]:
    """Dictionary summary wrapper around `validate_dataset`."""
    issue_objs = validate_dataset(
        df,
        mandatory_columns=mandatory_columns,
        allowed_dropdown_options=allowed_dropdown_options,
    )
    invalid_cells_map: Dict[str, Dict[str, Any]] = {}
    duplicate_rows: List[int] = []

    for i in issue_objs:
        d = i.to_dict()
        invalid_cells_map[f"{i.row_index}_{i.column}"] = d
        if i.issue_type == "duplicate":
            duplicate_rows.append(i.row_index + 1)

    return {
        "is_valid": len(issue_objs) == 0,
        "total_issues": len(issue_objs),
        "issues": [i.to_dict() for i in issue_objs],
        "duplicate_rows": duplicate_rows,
        "invalid_cells_map": invalid_cells_map,
    }

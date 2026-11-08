"""
llm.py
======
Centralized Google Gemini AI Provider for "URL-First Excel <-> Website Automation Assistant".
All LLM calls are isolated in this single module so the provider can be maintained or changed easily.

STRICT COMPLIANCE & PRIVACY:
- Uses the official `google-genai` package (`from google import genai`).
- Reads API Key from `st.secrets["GEMINI_API_KEY"]` first, then `os.environ.get("GEMINI_API_KEY")`. Never hardcoded.
- Reads Model from `os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")`.
- Requests JSON only (`response_mime_type="application/json"`) and validates against strict schemas.
- NEVER runs `eval()` or `exec()` on generated output.
- Retries up to 3 times with exponential backoff on 429 / quota errors, returning friendly Hindi/English messages.
- Receives ONLY column names, field labels, and 1-2 MASKED sample values (never raw Aadhaar, mobile, or bank numbers).
"""

from __future__ import annotations
import json
import os
import re
import time
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd

from utils import sanitize_sample_rows_for_ai

try:
    import streamlit as st
except ImportError:
    st = None

try:
    from google import genai
    from google.genai import types as genai_types
except ImportError:
    genai = None
    genai_types = None


BILINGUAL_SYNONYMS: List[Tuple[List[str], List[str]]] = [
    (["applicant name", "full name", "candidate name", "name", "आवेदक का नाम", "पूरा नाम", "नाम"], ["applicant_name", "fullname", "name", "candidate"]),
    (["father name", "father's name", "husband name", "पिता का नाम", "पिता", "pita ka naam"], ["father_name", "fname", "father"]),
    (["mother name", "mother's name", "माता का नाम", "माता", "mata ka naam"], ["mother_name", "mname", "mother"]),
    (["date of birth", "dob", "birth date", "जन्म तिथि", "जन्म दिनांक", "janm tithi"], ["dob", "birth_date", "dateofbirth"]),
    (["gender", "sex", "लिंग", "ling"], ["gender", "sex"]),
    (["category", "caste", "social category", "श्रेणी", "वर्ग", "जाति"], ["category", "caste"]),
    (["mobile", "mobile number", "phone", "मोबाइल नंबर", "मोबाइल", "फ़ोन"], ["mobile", "phone", "contact_no"]),
    (["aadhaar", "aadhar", "uid", "jan aadhaar", "आधार नंबर", "आधार संख्या", "जनाधार"], ["aadhaar", "aadhar", "uid", "janaadhaar"]),
    (["application no", "application id", "app no", "registration no", "आवेदन संख्या", "रजिस्ट्रेशन नंबर"], ["application_no", "app_id", "reg_no"]),
    (["status", "application status", "स्थिति", "आवेदन की स्थिति"], ["status", "app_status"]),
    (["district", "जिला", "zila"], ["district", "dist"]),
    (["block", "tehsil", "panchayat samiti", "ब्लॉक", "तहसील", "पंचायत समिति"], ["block", "tehsil"]),
    (["village", "ward", "gram panchayat", "गाँव", "ग्राम", "वार्ड"], ["village", "gram", "ward"]),
    (["pincode", "pin code", "pin_code", "postal code", "पिन कोड", "पिनकोड"], ["pincode", "pin", "postal"]),
    (["annual income", "income", "वार्षिक आय", "आय"], ["annual_income", "income"]),
    (["address", "permanent address", "पता", "स्थायी पता"], ["address", "addr"]),
    (["photo", "applicant photo", "passport photo", "photo_file", "फोटो", "आवेदक की फोटो"], ["applicant_photo", "photo", "image", "pic"]),
    (["document", "id proof", "aadhaar pdf", "दस्तावेज़", "पहचान पत्र"], ["document", "doc", "id_proof"]),
]


def get_gemini_api_key() -> str:
    """
    Retrieves GEMINI_API_KEY from `st.secrets["GEMINI_API_KEY"]` first,
    then falls back to environment variable `GEMINI_API_KEY`. Never hardcoded.
    """
    if st is not None:
        try:
            if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
                return str(st.secrets["GEMINI_API_KEY"]).strip()
        except Exception:
            pass
    return os.environ.get("GEMINI_API_KEY", "").strip()


def get_gemini_model_name() -> str:
    """Reads model name from environment variable GEMINI_MODEL (default: gemini-2.5-flash)."""
    return os.environ.get("GEMINI_MODEL", "gemini-2.5-flash").strip() or "gemini-2.5-flash"


def validate_mappings_schema(mappings: Any) -> List[Dict[str, Any]]:
    """
    Validates the JSON mapping list returned by Gemini against our strict schema.
    Never executes any code; only extracts safe string/number/boolean properties.
    """
    if not isinstance(mappings, list):
        return []
    allowed_transforms = {"none", "date_ddmmyyyy", "date_yyyymmdd", "upper", "lower", "title", "digits_only"}
    clean_list: List[Dict[str, Any]] = []
    for item in mappings:
        if not isinstance(item, dict):
            continue
        sel = str(item.get("field_selector", "")).strip()
        if not sel:
            continue
        transform = str(item.get("transform", "none")).strip().lower()
        if transform not in allowed_transforms:
            transform = "none"
        clean_list.append({
            "field_selector": sel,
            "field_label": str(item.get("field_label", sel)).strip(),
            "field_type": str(item.get("field_type", "text")).strip().lower(),
            "data_column": str(item.get("data_column", "")).strip(),
            "transform": transform,
            "depends_on_selector": item.get("depends_on_selector") if isinstance(item.get("depends_on_selector"), str) else None,
            "wait_after_ms": int(item.get("wait_after_ms", 200)) if str(item.get("wait_after_ms", "")).isdigit() else 200,
        })
    return clean_list


def call_gemini_json_with_retry(
    system_instruction: str,
    user_prompt: str,
    max_retries: int = 3
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Calls Google Gemini API (`google-genai`) requesting `response_mime_type="application/json"`.
    Retries up to 3 times with exponential backoff on 429 / quota / transient network errors.
    Returns: (parsed_json_dict_or_None, friendly_bilingual_error_or_None)
    """
    api_key = get_gemini_api_key()
    if not api_key:
        return None, (
            "GEMINI_API_KEY सेट नहीं है (Not set). कृपया साइडबार में या st.secrets / environment variable में अपनी Gemini API Key डालें। "
            "(Using built-in Smart Bilingual Dictionary fallback)."
        )

    if genai is None:
        return None, "google-genai पैकेज इंस्टॉल नहीं है। कृपया 'pip install google-genai' चलाएं।"

    model_name = get_gemini_model_name()
    client = genai.Client(api_key=api_key)

    for attempt in range(1, max_retries + 1):
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=user_prompt,
                config=genai_types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json"
                ) if genai_types else None
            )
            raw_text = (response.text or "").strip()
            match = re.search(r"\{.*\}", raw_text, re.DOTALL)
            json_str = match.group(0) if match else raw_text
            parsed = json.loads(json_str)
            if isinstance(parsed, dict):
                return parsed, None
            return {"data": parsed}, None

        except Exception as err:
            err_str = str(err)
            is_quota = "429" in err_str or "quota" in err_str.lower() or "resource_exhausted" in err_str.lower()
            if is_quota and attempt < max_retries:
                backoff_sec = 2 ** attempt
                time.sleep(backoff_sec)
                continue
            elif is_quota:
                return None, (
                    "⚠️ Gemini API कोटा सीमा (429 Quota Exceeded) पूरी हो गई है (3 बार प्रयास किया गया)। "
                    "ऐप ने स्मार्ट हिंदी/इंग्लिश डिक्शनरी मैपिंग का उपयोग किया है। (Quota limit reached; used smart local fallback)."
                )
            else:
                return None, f"Gemini API त्रुटि (Error): {err_str}"

    return None, "Gemini API call failed after retries."


def clean_and_enrich_scan_with_gemini(scanned_fields: List[Dict[str, Any]]) -> Tuple[List[Dict[str, Any]], str]:
    """
    STEP 1 (Gemini Scan Cleanup):
    Takes the raw scanned form fields from a URL and asks Gemini to:
    - Group fields into logical sections
    - Give each field a short, clean Excel column name (`suggested_column`)
    - Guess data type & validation rule (`10-digit mobile`, `12-digit Aadhaar`, `6-digit PIN`, `date_ddmmyyyy`, `dropdown`)
    - Confirm required fields
    Falls back deterministically if offline or no API key is configured.
    """
    # Deterministic local enrichment first
    enriched: List[Dict[str, Any]] = []
    for idx, f in enumerate(scanned_fields, start=1):
        fc = dict(f)
        combined = f"{fc.get('label', '')} {fc.get('name', '')} {fc.get('id', '')}".lower()
        short_col = fc.get("name") or fc.get("id") or f"Field_{idx}"
        short_col = re.sub(r"[^a-zA-Z0-9_]", "_", str(short_col)).strip("_").title()

        for keywords, canonical_list in BILINGUAL_SYNONYMS:
            if any(k in combined for k in keywords):
                short_col = canonical_list[0].title()
                break

        val_rule = "text"
        if "mobile" in combined or "phone" in combined or "मोबाइल" in combined:
            val_rule = "10-digit mobile"
        elif "aadhaar" in combined or "aadhar" in combined or "आधार" in combined:
            val_rule = "12-digit Aadhaar"
        elif "pin" in combined or "पिन" in combined:
            val_rule = "6-digit PIN code"
        elif "dob" in combined or "date" in combined or "जन्म" in combined or fc.get("type") == "date":
            val_rule = "date DD/MM/YYYY"
        elif fc.get("options"):
            val_rule = "dropdown allowed options"

        fc["suggested_column"] = short_col
        fc["validation_rule"] = val_rule
        enriched.append(fc)

    if not get_gemini_api_key():
        return enriched, "Used smart bilingual field analyzer (local)."

    system_prompt = """You are a Web Form Schema Analyzer.
Given a list of form fields scanned from a website (labels in Hindi/English/Hinglish), return JSON enriching each field with:
- "selector": exact selector from input
- "section": logical section title
- "suggested_column": short clean Excel column name (e.g. "Applicant_Name", "Father_Name", "Mobile", "Aadhaar", "District")
- "validation_rule": one of ["10-digit mobile", "12-digit Aadhaar", "6-digit PIN code", "date DD/MM/YYYY", "dropdown allowed options", "file upload", "text", "number"]
- "required": boolean
Return ONLY valid JSON: {"fields": [...]}
"""
    minimal_input = [
        {
            "selector": f.get("selector"),
            "label": f.get("label"),
            "type": f.get("type"),
            "required": f.get("required"),
            "has_options": bool(f.get("options")),
        }
        for f in enriched
    ]
    parsed, err = call_gemini_json_with_retry(system_prompt, json.dumps({"fields": minimal_input}, ensure_ascii=False))
    if parsed and isinstance(parsed.get("fields"), list):
        by_sel = {str(item.get("selector")): item for item in parsed["fields"] if isinstance(item, dict)}
        for f in enriched:
            g_item = by_sel.get(str(f.get("selector")))
            if g_item:
                if g_item.get("suggested_column"):
                    f["suggested_column"] = str(g_item["suggested_column"]).strip()
                if g_item.get("section"):
                    f["section"] = str(g_item["section"]).strip()
                if g_item.get("validation_rule"):
                    f["validation_rule"] = str(g_item["validation_rule"]).strip()
        return enriched, f"✨ Gemini ({get_gemini_model_name()}) ने फॉर्म स्कैन को साफ करके एक्सेल कॉलम नाम तैयार किए।"

    return enriched, err or "Enriched scan locally."


def heuristic_field_mapping(
    data_columns: List[str],
    portal_fields: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    Deterministic bilingual matcher that pairs Excel/Word columns with portal form fields
    (Hindi/English/Hinglish) when offline or as a fallback.
    """
    mappings: List[Dict[str, Any]] = []
    district_selector: Optional[str] = None
    block_selector: Optional[str] = None

    for pf in portal_fields:
        f_label = str(pf.get("label", "")).lower()
        f_name = str(pf.get("name", "")).lower()
        f_id = str(pf.get("id", "")).lower()
        f_sug = str(pf.get("suggested_column", "")).lower()
        f_type = str(pf.get("type", "text")).lower()
        f_sel = pf.get("selector", "")
        combined_portal = f"{f_label} {f_name} {f_id} {f_sug}"

        matched_col = ""
        for col in data_columns:
            col_clean = col.strip().lower().replace("_", " ")
            if col.strip().lower() == f_sug or col_clean in combined_portal or any(part in combined_portal for part in col_clean.split() if len(part) > 3):
                matched_col = col
                break

        if not matched_col:
            for keywords, _ in BILINGUAL_SYNONYMS:
                if any(k in combined_portal for k in keywords):
                    for col in data_columns:
                        c_low = col.strip().lower().replace("_", " ")
                        if any(k in c_low for k in keywords):
                            matched_col = col
                            break
                if matched_col:
                    break

        transform = "none"
        if f_type == "date" or "dob" in combined_portal or "जन्म" in combined_portal:
            transform = "date_yyyymmdd" if f_type == "date" else "date_ddmmyyyy"
        elif any(k in combined_portal for k in ["mobile", "aadhaar", "pincode", "मोबाइल", "आधार", "पिन"]):
            transform = "digits_only"

        depends_on = pf.get("depends_on_selector")
        wait_ms = 200
        if any(k in combined_portal for k in ["district", "जिला"]):
            district_selector = f_sel
            wait_ms = 600
        elif any(k in combined_portal for k in ["block", "tehsil", "ब्लॉक", "तहसील"]):
            block_selector = f_sel
            depends_on = depends_on or district_selector
            wait_ms = 600
        elif any(k in combined_portal for k in ["village", "gram", "गाँव", "ग्राम"]):
            depends_on = depends_on or block_selector
            wait_ms = 600

        mappings.append({
            "field_selector": f_sel,
            "field_label": pf.get("label", f_name or f_id),
            "field_type": f_type,
            "data_column": matched_col,
            "transform": transform,
            "depends_on_selector": depends_on,
            "wait_after_ms": wait_ms,
            "required": pf.get("required", False),
            "allowed_options": pf.get("options", []),
        })

    return mappings


def suggest_form_mapping_with_gemini(
    df: pd.DataFrame,
    portal_fields: List[Dict[str, Any]]
) -> Tuple[List[Dict[str, Any]], str]:
    """
    Uses Gemini (`GEMINI_MODEL`, default `gemini-2.5-flash`) to map
    Excel/Word columns to scanned web form fields.
    Sends ONLY column names, field labels, and 1-2 MASKED sample rows.
    """
    data_columns = [str(c) for c in df.columns]
    raw_samples = df.head(2).to_dict(orient="records") if not df.empty else []
    masked_samples = sanitize_sample_rows_for_ai(raw_samples)

    system_prompt = """You are a Smart Bilingual Portal Form Mapper.
Match the user's Excel/Word columns to the scanned web form fields.
Understand Hindi, English, and Hinglish labels (e.g. "पिता का नाम" = "Father_Name", "जिला" = "District").
Return ONLY valid JSON matching this schema:
{
  "mappings": [
    {
      "field_selector": "#applicant_name",
      "field_label": "आवेदक का नाम / Applicant Name",
      "field_type": "text",
      "data_column": "Applicant_Name",
      "transform": "none",
      "depends_on_selector": null,
      "wait_after_ms": 200
    }
  ]
}
Allowed transforms: "none", "date_ddmmyyyy", "date_yyyymmdd", "upper", "lower", "title", "digits_only".
"""
    user_payload = json.dumps({
        "excel_columns": data_columns,
        "masked_sample_values": masked_samples,
        "web_form_fields": portal_fields,
    }, ensure_ascii=False, indent=2)

    parsed, err = call_gemini_json_with_retry(system_prompt, user_payload)
    if parsed and "mappings" in parsed:
        validated = validate_mappings_schema(parsed["mappings"])
        if validated:
            return validated, f"✨ Gemini ({get_gemini_model_name()}) ने स्मार्ट फील्ड मैपिंग तैयार की।"

    fallback = heuristic_field_mapping(data_columns, portal_fields)
    status_note = err if err else "Mapped via Smart Bilingual Dictionary."
    return fallback, status_note


def generate_field_mapping_with_gemini(
    df: pd.DataFrame,
    scanned_fields: List[Dict[str, Any]]
) -> Dict[str, Any]:
    """Wrapper used by app.py returning dict with `mappings`, `used_fallback`, and `message_hi`."""
    mappings, msg = suggest_form_mapping_with_gemini(df, scanned_fields)
    used_fallback = not msg.startswith("✨")
    return {
        "mappings": mappings,
        "used_fallback": used_fallback,
        "message_hi": msg,
    }


def suggest_extraction_cleanup_with_gemini(
    raw_headers: List[str],
    masked_sample_rows: List[Dict[str, str]]
) -> Tuple[Dict[str, str], str]:
    """Suggests clean, standardized Hindi/English column headers for tables scraped from a website."""
    system_prompt = """You are a Data Extraction Schema Cleaner.
Given raw table headers scraped from a website and 1-2 masked sample rows, suggest clean, standardized column names.
Return ONLY valid JSON:
{
  "header_mapping": {
    "raw_header_1": "Clean_Column_Name_1"
  }
}
"""
    payload = json.dumps({
        "raw_headers": raw_headers,
        "masked_samples": sanitize_sample_rows_for_ai(masked_sample_rows[:2]),
    }, ensure_ascii=False, indent=2)

    parsed, err = call_gemini_json_with_retry(system_prompt, payload)
    if parsed and "header_mapping" in parsed and isinstance(parsed["header_mapping"], dict):
        clean_map = {str(k): str(v) for k, v in parsed["header_mapping"].items()}
        return clean_map, f"✨ Gemini ({get_gemini_model_name()}) ने टेबल कॉलम नाम साफ़ किए।"

    default_map = {h: re.sub(r"\s+", "_", str(h).strip()) for h in raw_headers}
    return default_map, err or "Standardized column headers locally."


def suggest_extraction_selectors_with_gemini(
    raw_headers: List[str],
    sample_rows: Optional[List[Dict[str, str]]] = None
) -> Dict[str, Any]:
    mapping, msg = suggest_extraction_cleanup_with_gemini(raw_headers, sample_rows or [])
    return {"header_mapping": mapping, "message_hi": msg}


class PortalTemplateStore:
    """Manages reusable JSON templates for URL Scans, Form Filling, and Website Extraction."""

    def __init__(self, templates_dir: str = "templates"):
        self.templates_dir = templates_dir
        os.makedirs(self.templates_dir, exist_ok=True)

    def save_template(
        self,
        name: str = "",
        portal_url: str = "",
        mode: str = "excel_to_website",
        mappings: Optional[List[Dict[str, Any]]] = None,
        scanned_fields: Optional[List[Dict[str, Any]]] = None,
        extra_config: Optional[Dict[str, Any]] = None,
        template_name: Optional[str] = None,
        template_type: Optional[str] = None,
    ) -> str:
        final_name = template_name or name or "Portal_Template"
        final_type = template_type or mode
        safe_name = re.sub(r"[^a-zA-Z0-9_\- ]", "", final_name).strip().replace(" ", "_") or "template"
        path = os.path.join(self.templates_dir, f"{safe_name}.json")
        payload = {
            "template_name": final_name,
            "template_type": final_type,
            "portal_url": portal_url,
            "updated_at": pd.Timestamp.now().isoformat(),
            "extra_config": extra_config or {},
            "scanned_fields": scanned_fields or [],
            "mappings": mappings or [],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        return path

    def list_templates(self, template_type: Optional[str] = None) -> List[str]:
        names: List[str] = []
        if not os.path.exists(self.templates_dir):
            return names
        for fname in sorted(os.listdir(self.templates_dir)):
            if fname.endswith(".json"):
                if template_type is None:
                    names.append(fname[:-5])
                else:
                    try:
                        with open(os.path.join(self.templates_dir, fname), "r", encoding="utf-8") as f:
                            data = json.load(f)
                        if data.get("template_type") == template_type:
                            names.append(fname[:-5])
                    except Exception:
                        pass
        return names

    def load_template(self, name: str) -> Optional[Dict[str, Any]]:
        if not name:
            return None
        fname = name if name.endswith(".json") else f"{name}.json"
        path = os.path.join(self.templates_dir, fname)
        if not os.path.exists(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

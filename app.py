"""
app.py
======
URL-First Excel <-> Website Automation Assistant (Streamlit Main Entry Point)

Supports 4 top-level workflow modes via the top Mode Switch:
1. "Scan & Make Excel Format" (Steps 1 & 2):
   - Open any website URL in a visible Chromium browser on Desktop, or seamlessly in Cloud Browser / Direct-Scan mode on Streamlit Community Cloud.
   - Scan EVERYTHING: form fields, dropdown options, cascading dropdowns (District -> Block -> Village), multi-step forms, file uploads, HTML tables, pagination, and detail links.
   - Show "Form fields found: N, Tables found: M, Uploads: K" report.
   - Gemini cleans up the scan and auto-generates a ready-to-fill 3-sheet Excel template (.xlsx) with DataValidation dropdown lists, required highlights, notes row, "Instructions" sheet, and "Allowed Values" sheet.
   - Option: "I already have my own Excel" — upload it, map columns with Gemini, and reshape it into the generated template format.
2. "Fill Website from Excel" (Step 3):
   - Import Excel/CSV/Word, choose sheet and row range, pre-validate (mandatory blanks, 10-digit mobile, 12-digit Aadhaar, 6-digit PIN, dates, allowed dropdown values, duplicates).
   - Fill website row-by-row with fuzzy dropdown matching, cascading dropdown waits, photo auto-resize/compression (<=50 KB), Dry Run (Test Mode), Semi-Automatic review ("Approve & Submit"), crash resume, retries, and proof screenshots.
3. "Website to Excel" (Step 4):
   - Extract any scanned table/list across multiple pages (`max_pages`), Search-and-Fetch status for an Excel of IDs, scrape Detail Pages, and download linked PDFs/certificates.
4. "Fill + Verify" (Step 5):
   - Re-verify submitted application numbers on the website (`Submitted` -> `Approved` / `Pending` / `Rejected`) and reconcile Excel vs. Website data into a color-coded `Differences` Excel sheet.
"""

import os
import io
import json
import pandas as pd
import streamlit as st

from importer import load_spreadsheet_sheets, load_word_document
from validator import validate_dataset
from browser import PortalBrowserManager, fetch_url_html_direct
from scanner import scan_full_page_live, scan_full_html_string
from template_maker import generate_excel_template_from_scan, reshape_user_excel_to_template
from llm import (
    get_gemini_api_key,
    get_gemini_model_name,
    clean_and_enrich_scan_with_gemini,
    generate_field_mapping_with_gemini,
    PortalTemplateStore,
)
from uploader import match_document_for_row
from filler import fill_row_into_portal
from runner import BatchExecutionState
from extractor import (
    detect_html_tables_from_string,
    extract_paginated_table_live,
    search_and_fetch_from_excel_live,
    extract_detail_pages_live,
    clean_extracted_dataframe,
)
from reconcile import reconcile_excel_vs_website, export_reconciliation_excel
from report import export_results_excel, export_formatted_excel_bytes, generate_pdf_summary_report
from utils import (
    mask_sensitive_value,
    parse_voice_assistant_command,
)

st.set_page_config(
    page_title="URL-First Excel <-> Website Automation Assistant",
    page_icon="🌐",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Session State Initialization
# ---------------------------------------------------------------------------
if "ui_lang" not in st.session_state:
    st.session_state.ui_lang = "Bilingual (Hindi + English)"
if "scan_report" not in st.session_state:
    st.session_state.scan_report = None
if "scanned_fields" not in st.session_state:
    st.session_state.scanned_fields = []
if "active_df" not in st.session_state:
    st.session_state.active_df = None
if "raw_excel_bytes" not in st.session_state:
    st.session_state.raw_excel_bytes = None
if "field_mappings" not in st.session_state:
    st.session_state.field_mappings = []
if "browser_mgr" not in st.session_state:
    st.session_state.browser_mgr = PortalBrowserManager(profile_dir="./browser_profile")
if "batch_state" not in st.session_state:
    st.session_state.batch_state = BatchExecutionState()
if "extracted_df" not in st.session_state:
    st.session_state.extracted_df = None
if "reconcile_summary_df" not in st.session_state:
    st.session_state.reconcile_summary_df = None
if "reconcile_diff_df" not in st.session_state:
    st.session_state.reconcile_diff_df = None

template_store = PortalTemplateStore("./templates")


def scan_url_or_active_browser(url_to_scan: str, start_from_current: bool = True, scan_multistep: bool = True):
    """
    Opens `url_to_scan` in Playwright (or uses Direct HTTP fetch fallback on restricted cloud containers)
    and returns `(ok, message, scan_report)`.
    """
    mgr: PortalBrowserManager = st.session_state.browser_mgr

    if not mgr.page and not mgr.fallback_html:
        with st.spinner("Launching browser / connecting to URL (auto-installing Chromium if needed)..."):
            ok, msg = mgr.launch_headed_browser(url_to_scan)
            if not ok:
                return False, msg, None
    elif mgr.page and not start_from_current and url_to_scan:
        try:
            mgr.page.goto(url_to_scan, wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass
    elif not mgr.page and url_to_scan:
        ok_http, html_or_err, final_u = fetch_url_html_direct(url_to_scan)
        if ok_http:
            mgr.fallback_html = html_or_err
            mgr.fallback_url = final_u
        else:
            return False, html_or_err, None

    if mgr.page:
        report = scan_full_page_live(
            mgr.page,
            probe_cascading=True,
            scan_multi_steps=scan_multistep,
        )
    else:
        report = scan_full_html_string(mgr.fallback_html, source_url=mgr.fallback_url or url_to_scan)

    enriched_fields, ai_msg = clean_and_enrich_scan_with_gemini(report.get("fields", []))
    report["fields"] = enriched_fields
    st.session_state.scan_report = report
    st.session_state.scanned_fields = enriched_fields

    # If the page also contains HTML tables, pre-populate extracted_df with the first table for convenience
    if report.get("tables") and st.session_state.extracted_df is None:
        html_src = mgr.page.content() if mgr.page else mgr.fallback_html
        dfs = detect_html_tables_from_string(html_src)
        if dfs:
            st.session_state.extracted_df = clean_extracted_dataframe(dfs[0])

    return True, f"✅ {report['summary_text']} | {ai_msg}", report


# ---------------------------------------------------------------------------
# Sidebar Configuration & Safety Controls
# ---------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Assistant Settings")
    st.caption("URL-First Human-in-the-Loop Automation")

    st.session_state.ui_lang = st.selectbox(
        "भाषा / Interface Language",
        ["Bilingual (Hindi + English)", "English", "हिंदी (Hindi)"],
        index=0,
    )

    st.divider()
    st.subheader("🔑 Google Gemini API (`llm.py`)")
    active_key = get_gemini_api_key()
    active_model = get_gemini_model_name()
    if active_key:
        st.success(f"GEMINI_API_KEY active (`{active_model}`)")
    else:
        st.warning("GEMINI_API_KEY not found in st.secrets or env.")
        manual_key = st.text_input("Enter GEMINI_API_KEY (Session Only):", type="password")
        if manual_key:
            os.environ["GEMINI_API_KEY"] = manual_key
            st.rerun()

    st.divider()
    st.subheader("🌐 Playwright Browser & URL")
    default_demo_path = os.path.abspath("./samples/demo_portal.html")
    default_url = f"file://{default_demo_path}" if os.path.exists(default_demo_path) else "https://sso.rajasthan.gov.in"
    portal_url = st.text_input("Website URL:", value=default_url)

    col_b1, col_b2 = st.columns(2)
    with col_b1:
        if st.button("🚀 Open URL", use_container_width=True):
            with st.spinner("Opening URL (auto-installing Chromium on first cloud run)..."):
                ok, msg = st.session_state.browser_mgr.launch_headed_browser(portal_url)
            if ok:
                st.success(msg)
            else:
                st.error(msg)
    with col_b2:
        if st.button("🛑 Close Browser", use_container_width=True):
            st.session_state.browser_mgr.close()
            st.toast("Browser closed.")

    # Show live browser preview when running in Cloud Headless Mode (e.g., Streamlit Cloud on Mobile)
    live_shot = st.session_state.browser_mgr.get_live_screenshot_bytes()
    if live_shot:
        with st.expander("🖥️ Live Browser View (Cloud Preview)", expanded=True):
            st.image(live_shot, caption=st.session_state.browser_mgr.page.url, use_container_width=True)

    st.divider()
    st.subheader("🛡️ Safety & Rate Limiting")
    action_delay = st.slider("Delay between actions (sec):", 0.5, 5.0, 1.2, 0.1)
    dry_run_mode = st.toggle("🧪 Dry Run / Test Mode (Fill 1st row only, never submit)", value=False)
    semi_auto_mode = st.toggle("Semi-Automatic Mode (Stop before Submit)", value=True)
    auto_submit_confirmed = False
    if not semi_auto_mode and not dry_run_mode:
        st.warning("⚠️ Auto-Submit Enabled! The app will click Submit automatically after filling each row.")
        auto_submit_confirmed = st.checkbox("I confirm Auto-Submit without manual review per row.")

    mask_pii_ui = st.toggle("Mask Aadhaar/Mobile/Bank in UI & Logs", value=True)

    st.divider()
    with st.expander("🔒 Privacy & Terms of Use Notice", expanded=False):
        st.markdown(
            """
            - **All Data Stays on Your Computer**: Files are processed locally. Gemini receives only field labels, column names, and 1–2 masked sample values (`XXXXXXXX1234`).
            - **No Login / OTP / CAPTCHA Bypass**: When the site asks for them, the assistant pauses with *"Please login / enter OTP / captcha in the browser, then click Continue"*.
            - **Authorised Access Only**: Use only on data and websites you are authorised to access. Always check the website's Terms of Use.
            """
        )

# ---------------------------------------------------------------------------
# Top Header & 4-Mode Switch
# ---------------------------------------------------------------------------
st.title("🌐 URL-First Excel ↔ Website Automation Assistant")
st.markdown(
    "**Step 1:** Give a Website URL & Scan Everything → **Step 2:** Auto-Generate Ready-to-Fill Excel Format → **Step 3:** Fill Website from Excel → **Step 4:** Extract Website Data to Excel → **Step 5:** Fill + Verify & Reconcile."
)

mode = st.radio(
    "Select Workflow Mode / कार्य मोड चुनें:",
    [
        "Scan & Make Excel Format",
        "Fill Website from Excel",
        "Website to Excel",
        "Fill + Verify",
    ],
    horizontal=True,
)

# Voice / Natural-Language Command Bar
with st.expander("🎙️ Voice / Natural-Language Command Bar (Hindi, English, Hinglish)", expanded=False):
    col_v1, col_v2 = st.columns([3, 1])
    with col_v1:
        voice_text = st.text_input(
            "Speak or type a command (e.g., 'is URL ko scan karo', 'Excel format banao', 'pehli 10 rows bharo', 'sirf failed rows dobara chalao', 'table Excel mein nikalo', 'sabka status check karo'):",
            placeholder="is URL ko scan karo | Excel format banao | pehli 10 rows bharo | sabka status check karo",
        )
    with col_v2:
        try:
            from streamlit_mic_recorder import speech_to_text
            spoken = speech_to_text(language="hi-IN", start_prompt="🎤 बोलें / Speak", stop_prompt="⏹️ Stop", key="mic_cmd")
            if spoken:
                voice_text = spoken
        except Exception:
            st.caption("Install `streamlit-mic-recorder` for live microphone button.")

    if voice_text:
        parsed_cmd = parse_voice_assistant_command(voice_text)
        st.info(f"**Recognized Command**: `{parsed_cmd['intent']}` — {parsed_cmd['message_hi']} ({parsed_cmd['message_en']})")

st.divider()


# ===========================================================================
# MODE 1: SCAN & MAKE EXCEL FORMAT (STEPS 1 & 2)
# ===========================================================================
if mode == "Scan & Make Excel Format":
    st.subheader("🔍 Step 1: URL Scan (Scan Forms, Dropdowns, Multi-Step Wizards, Uploads & Tables)")

    col_u1, col_u2, col_u3 = st.columns([2, 1, 1])
    with col_u1:
        target_scan_url = st.text_input("Website URL to Open & Scan:", value=portal_url)
    with col_u2:
        start_from_current = st.checkbox("Start from the page I am on now (after Login)", value=False)
    with col_u3:
        scan_multistep = st.checkbox("Safe Multi-Step Scan (Never Submits)", value=True)

    if st.session_state.browser_mgr.page:
        obstacle = st.session_state.browser_mgr.detect_security_obstacle()
        if obstacle["paused"]:
            st.warning(f"⏸️ **{obstacle['reason']}** — {obstacle['reason_hi']}")
            if st.button("▶️ Continue After Login / OTP / CAPTCHA"):
                st.rerun()

    col_sc1, col_sc2, col_sc3 = st.columns(3)
    with col_sc1:
        if st.button("🌐 Open URL & Scan This Page", type="primary", use_container_width=True):
            ok, msg, _ = scan_url_or_active_browser(
                url_to_scan=target_scan_url,
                start_from_current=start_from_current,
                scan_multistep=scan_multistep,
            )
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    with col_sc2:
        if st.button("🧪 Scan Bundled Local Demo Portal (`demo_portal.html`)", use_container_width=True):
            demo_path = "./samples/demo_portal.html"
            if os.path.exists(demo_path):
                with open(demo_path, "r", encoding="utf-8") as f:
                    report = scan_full_html_string(f.read(), source_url=demo_path)
                enriched_fields, ai_msg = clean_and_enrich_scan_with_gemini(report["fields"])
                report["fields"] = enriched_fields
                st.session_state.scan_report = report
                st.session_state.scanned_fields = enriched_fields
                st.success(f"✅ {report['summary_text']} | {ai_msg}")

    with col_sc3:
        saved_templates = template_store.list_templates()
        chosen_tpl = st.selectbox("📂 Or Load Saved Template JSON:", ["-- Select Template --"] + saved_templates)
        if chosen_tpl != "-- Select Template --":
            tpl_obj = template_store.load_template(chosen_tpl)
            if tpl_obj:
                st.session_state.field_mappings = tpl_obj.get("mappings", [])
                if tpl_obj.get("scanned_fields"):
                    st.session_state.scanned_fields = tpl_obj["scanned_fields"]
                st.toast(f"Loaded template `{chosen_tpl}`")

    if st.session_state.scan_report or st.session_state.scanned_fields:
        rep = st.session_state.scan_report or {
            "fields": st.session_state.scanned_fields,
            "tables": [],
            "uploads": [f for f in st.session_state.scanned_fields if f.get("type") == "file"],
            "summary_text": f"Form fields found: {len(st.session_state.scanned_fields)}",
        }
        fields = st.session_state.scanned_fields
        tables_found = rep.get("tables", [])

        st.markdown(f"### 📊 URL Scan Report — `{rep.get('summary_text')}`")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Form Fields Found", len(fields))
        m2.metric("Tables Found", len(tables_found))
        m3.metric("File Upload Fields", len(rep.get("uploads", [])))
        m4.metric("Dropdown / Radio Fields", sum(1 for f in fields if f.get("options")))

        if fields:
            fields_preview_df = pd.DataFrame([
                {
                    "Order": i + 1,
                    "Section": f.get("section", "Form"),
                    "Website Field Label": f.get("label", ""),
                    "Excel Column Name": f.get("suggested_column") or f.get("name") or f.get("id"),
                    "Type": f.get("type", "text"),
                    "Required": "Yes *" if f.get("required") else "No",
                    "Validation / Format": f.get("validation_rule", "text"),
                    "Allowed Options (Dropdown/Radio)": ", ".join(f.get("options", [])[:6]),
                    "Selector": f.get("selector", ""),
                }
                for i, f in enumerate(fields)
            ])
            st.dataframe(fields_preview_df, use_container_width=True)

        if tables_found:
            with st.expander(f"📋 Detected HTML Tables on Page ({len(tables_found)})", expanded=True):
                for tbl in tables_found:
                    st.markdown(f"**Table #{tbl['table_index']}: {tbl['title']}** ({tbl['row_count']} rows)")
                    st.dataframe(pd.DataFrame(tbl["preview_rows"]), use_container_width=True)

        st.divider()
        st.subheader("📥 Step 2: Auto-Generate Ready-to-Fill Excel Format (.xlsx)")
        st.markdown(
            "Generates an Excel template with **1 column per form field**, **highlighted required columns**, **Excel DataValidation dropdown lists** from the website's actual options, **Photo/Status/Application No columns**, an **Instructions** sheet (Hindi/English), and an **Allowed Values** sheet."
        )

        col_g1, col_g2 = st.columns([2, 1])
        with col_g1:
            tpl_save_name = st.text_input("Template Name (for saving JSON & Excel):", value="SSO_New_Registration")
        with col_g2:
            if st.button("💾 Save Template Mapping JSON", use_container_width=True):
                auto_mappings = [
                    {
                        "field_selector": f.get("selector", ""),
                        "field_label": f.get("label", ""),
                        "field_type": f.get("type", "text"),
                        "data_column": f.get("suggested_column") or f.get("name") or f.get("id"),
                        "transform": "date_ddmmyyyy" if "date" in str(f.get("validation_rule", "")) else "none",
                        "depends_on_selector": f.get("depends_on_selector"),
                        "wait_after_ms": 500 if f.get("options") else 200,
                        "required": bool(f.get("required")),
                        "allowed_options": f.get("options", []),
                    }
                    for f in fields
                ]
                st.session_state.field_mappings = auto_mappings
                saved_path = template_store.save_template(
                    name=tpl_save_name,
                    portal_url=target_scan_url,
                    mode="excel_to_website",
                    mappings=auto_mappings,
                    scanned_fields=fields,
                )
                st.success(f"Saved reusable template to `{saved_path}`!")

        excel_template_bytes = generate_excel_template_from_scan(
            scanned_fields=fields,
            portal_url=target_scan_url,
            template_name=tpl_save_name,
        )
        st.download_button(
            "📥 Download Auto-Generated Website Excel Format (.xlsx) — 3 Sheets (Template + Instructions + Allowed Values)",
            data=excel_template_bytes,
            file_name=f"{tpl_save_name}_Template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        st.divider()
        st.subheader("🔄 Option: 'I Already Have My Own Excel' (Map & Reshape to Website Format)")
        own_file = st.file_uploader("Upload your existing Excel/CSV/Word file to map & reshape:", type=["xlsx", "csv", "docx"], key="own_excel_reshape")
        if own_file is not None:
            raw_b = own_file.read()
            if own_file.name.lower().endswith(".docx"):
                u_df = load_word_document(raw_b, own_file.name)["combined_df"]
            else:
                sheets = load_spreadsheet_sheets(raw_b, own_file.name)
                u_df = sheets[list(sheets.keys())[0]]

            res_map = generate_field_mapping_with_gemini(u_df, fields)
            st.session_state.field_mappings = res_map["mappings"]
            st.info(res_map["message_hi"])

            edited_map = st.data_editor(pd.DataFrame(st.session_state.field_mappings), use_container_width=True, key="reshape_map_editor")
            st.session_state.field_mappings = edited_map.to_dict(orient="records")

            reshaped_df = reshape_user_excel_to_template(u_df, fields, st.session_state.field_mappings)
            st.markdown("#### ✨ Reshaped Data Preview (Matching Website Template Columns)")
            st.dataframe(reshaped_df.head(10), use_container_width=True)

            reshaped_bytes = export_formatted_excel_bytes(reshaped_df, sheet_title="Form_Data_Template")
            st.download_button(
                "📥 Download Reshaped Excel in Website Format (.xlsx)",
                data=reshaped_bytes,
                file_name=f"{tpl_save_name}_Reshaped_Data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


# ===========================================================================
# MODE 2: FILL WEBSITE FROM EXCEL (STEP 3)
# ===========================================================================
elif mode == "Fill Website from Excel":
    tab_import, tab_run, tab_report = st.tabs([
        "1️⃣ Import & Pre-Flight Validation",
        "2️⃣ Row-by-Row & Batch Filler",
        "3️⃣ Proof Screenshots & Updated Excel Copy",
    ])

    with tab_import:
        st.subheader("📂 Upload Filled Excel Template, Custom Excel/CSV, or Word (.docx)")
        col_u1, col_u2 = st.columns([2, 1])
        with col_u1:
            uploaded_file = st.file_uploader("Choose Excel (.xlsx/.csv) or Word (.docx)", type=["xlsx", "xls", "csv", "docx"])
        with col_u2:
            st.markdown("**Or test with bundled sample data:**")
            if st.button("📥 Load `sample_applicants.xlsx`", use_container_width=True):
                sample_path = "./samples/sample_applicants.xlsx"
                if os.path.exists(sample_path):
                    with open(sample_path, "rb") as f:
                        raw_bytes = f.read()
                    st.session_state.raw_excel_bytes = raw_bytes
                    sheets = load_spreadsheet_sheets(raw_bytes, "sample_applicants.xlsx")
                    st.session_state.active_df = sheets[list(sheets.keys())[0]]
                    st.success(f"Loaded `{sample_path}` ({len(st.session_state.active_df)} rows)")

        if uploaded_file is not None:
            raw_bytes = uploaded_file.read()
            fname = uploaded_file.name
            if fname.lower().endswith(".docx"):
                doc_res = load_word_document(raw_bytes, fname)
                st.session_state.active_df = doc_res["combined_df"]
            else:
                st.session_state.raw_excel_bytes = raw_bytes
                sheets = load_spreadsheet_sheets(raw_bytes, fname)
                sheet_name = st.selectbox("Select Sheet:", list(sheets.keys()))
                df_loaded = sheets[sheet_name]
                if not df_loaded.empty and any("Required (" in str(v) or "Optional (" in str(v) for v in df_loaded.iloc[0].values):
                    df_loaded = df_loaded.iloc[1:].reset_index(drop=True)
                st.session_state.active_df = df_loaded

        if st.session_state.active_df is not None:
            df = st.session_state.active_df
            st.markdown("#### ✏️ Editable Data Grid (Fix red validation issues directly below)")
            edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key="step3_editor")
            st.session_state.active_df = edited_df

            allowed_opts_map = {}
            for f in st.session_state.scanned_fields:
                col_key = f.get("suggested_column") or f.get("name")
                if col_key and f.get("options"):
                    allowed_opts_map[str(col_key)] = f["options"]

            mandatory_cols = st.multiselect(
                "Mandatory Columns (अनिवार्य कॉलम):",
                options=list(edited_df.columns),
                default=[c for c in edited_df.columns if any(k in c.lower() for k in ["name", "aadhaar", "mobile", "district"])][:4],
            )
            issues = validate_dataset(edited_df, mandatory_columns=mandatory_cols, allowed_dropdown_options=allowed_opts_map)
            if not issues:
                st.success("🎉 All rows passed validation (Mandatory, Mobile 10-digit, Aadhaar 12-digit, PIN 6-digit, Dates, Dropdown values, Duplicates)!")
            else:
                st.error(f"Found {len(issues)} validation issue(s). Please fix them in the grid above:")
                st.dataframe(
                    pd.DataFrame([
                        {
                            "Row": i.row_index + 1,
                            "Column": i.column,
                            "Value": i.current_value,
                            "Message (HI/Hinglish)": i.message_hi,
                            "Message (EN)": i.message_en,
                        }
                        for i in issues
                    ]),
                    use_container_width=True,
                )

    with tab_run:
        st.subheader("🤖 Fill Website from Excel (Semi-Auto Review & Batch Safety)")
        if st.session_state.active_df is None:
            st.info("Please load an Excel/Word file in Tab 1 first.")
        else:
            df = st.session_state.active_df
            if not st.session_state.field_mappings and st.session_state.scanned_fields:
                st.session_state.field_mappings = generate_field_mapping_with_gemini(df, st.session_state.scanned_fields)["mappings"]
            elif not st.session_state.field_mappings:
                tpl = template_store.load_template("SSO_Citizen_Registration")
                if tpl:
                    st.session_state.field_mappings = tpl.get("mappings", [])

            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                photos_folder = st.text_input("📁 Photos / Documents Folder:", value="./samples/photos")
            with col_p2:
                max_photo_kb = st.number_input("Max Photo KB (Auto-Resize):", value=50, min_value=10, max_value=500)
            with col_p3:
                curr_idx = st.number_input(
                    "Active Row Index (0-based):",
                    min_value=0,
                    max_value=max(0, len(df) - 1),
                    value=min(st.session_state.batch_state.current_row_index, max(0, len(df) - 1)),
                )
                st.session_state.batch_state.current_row_index = curr_idx

            current_row = df.iloc[curr_idx].to_dict()
            masked_preview = {k: (mask_sensitive_value(v, k) if mask_pii_ui else v) for k, v in current_row.items()}
            st.json(masked_preview)

            matched_photo = match_document_for_row(current_row, photos_folder, row_number=curr_idx + 1)
            if matched_photo:
                st.caption(f"📎 Matched Photo/Document: `{matched_photo}` (Will be compressed ≤ {max_photo_kb} KB)")
            else:
                st.caption("⚠️ No matching photo file found in folder for this row.")

            if st.session_state.browser_mgr.page:
                obstacle = st.session_state.browser_mgr.detect_security_obstacle()
                if obstacle["paused"]:
                    st.error(f"⏸️ {obstacle['reason']}")
                    if st.button("▶️ Continue After Login / OTP / CAPTCHA", type="primary"):
                        st.rerun()

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                if st.button("✍️ Fill Current Row (Stop Before Submit)", type="primary", use_container_width=True):
                    if not st.session_state.browser_mgr.page:
                        ok, msg = st.session_state.browser_mgr.launch_headed_browser(portal_url)
                        if not ok or not st.session_state.browser_mgr.page:
                            st.error("Playwright interactive browser is required for form filling. Run locally or check cloud Chromium packages.")
                    if st.session_state.browser_mgr.page:
                        fill_res = fill_row_into_portal(
                            page=st.session_state.browser_mgr.page,
                            row_data=current_row,
                            mappings=st.session_state.field_mappings,
                            photos_dir=photos_folder,
                            max_photo_kb=max_photo_kb,
                            action_delay=action_delay,
                            row_number=curr_idx + 1,
                        )
                        if fill_res["status"] == "PAUSED_FOR_HUMAN":
                            st.warning(fill_res["message"])
                        elif fill_res["status"] == "READY_FOR_REVIEW":
                            st.session_state.batch_state.record_row_outcome(curr_idx, status="Filled")
                            if dry_run_mode:
                                st.info("🧪 Dry Run / Test Mode active: Filled 1st row and stopped without submitting.")
                            else:
                                st.success("✅ Form filled! Review in browser and click 'Approve & Submit' below.")
                        else:
                            st.session_state.batch_state.record_row_outcome(curr_idx, status="Failed", error_msg=fill_res["message"])
                            st.error(fill_res["message"])

            with c2:
                if st.button("✅ Approve & Submit Row", disabled=dry_run_mode, use_container_width=True):
                    if st.session_state.browser_mgr.page:
                        sub_res = st.session_state.browser_mgr.click_submit_and_capture(row_index=curr_idx)
                        st.session_state.batch_state.record_row_outcome(
                            row_index=curr_idx,
                            status=sub_res["status"],
                            reference_no=sub_res["reference_no"],
                            screenshot_path=sub_res["screenshot_path"],
                            error_msg=sub_res.get("error", ""),
                        )
                        st.session_state.batch_state.current_row_index = min(len(df) - 1, curr_idx + 1)
                        st.success(f"Submitted Row #{curr_idx + 1}! Application No: `{sub_res['reference_no']}`")
                    else:
                        st.error("Browser is not open.")

            with c3:
                if st.button("⏭️ Skip Row", use_container_width=True):
                    st.session_state.batch_state.record_row_outcome(curr_idx, status="Skipped", error_msg="Skipped by user")
                    st.session_state.batch_state.current_row_index = min(len(df) - 1, curr_idx + 1)
                    st.rerun()

            with c4:
                failed_indices = st.session_state.batch_state.get_failed_row_indices()
                if st.button(f"🔄 Retry Failed Rows ({len(failed_indices)})", use_container_width=True):
                    if failed_indices:
                        st.session_state.batch_state.current_row_index = failed_indices[0]
                        st.rerun()

    with tab_report:
        st.subheader("📊 Proof Screenshots, Status Log & Non-Destructive Excel Export")
        outcomes = st.session_state.batch_state.row_outcomes
        if not outcomes:
            st.info("No rows processed yet in this session.")
        else:
            st.dataframe(pd.DataFrame.from_dict(outcomes, orient="index"), use_container_width=True)
            if st.session_state.active_df is not None:
                updated_excel_bytes = export_results_excel(
                    original_df=st.session_state.active_df,
                    outcomes=outcomes,
                    raw_original_bytes=st.session_state.raw_excel_bytes,
                )
                d1, d2 = st.columns(2)
                with d1:
                    st.download_button(
                        "📥 Download Updated Excel Copy (Never Overwrites Original)",
                        data=updated_excel_bytes,
                        file_name="applicants_results_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                with d2:
                    pdf_bytes = generate_pdf_summary_report(outcomes, df=st.session_state.active_df)
                    st.download_button(
                        "📄 Download PDF Audit Summary Report",
                        data=pdf_bytes,
                        file_name="automation_audit_report.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )


# ===========================================================================
# MODE 3: WEBSITE TO EXCEL DATA (STEP 4)
# ===========================================================================
elif mode == "Website to Excel":
    tab_table, tab_search, tab_detail = st.tabs([
        "1️⃣ Multi-Page Table & List Extractor",
        "2️⃣ Search-and-Fetch Status from Excel",
        "3️⃣ Detail-Page Scraper & File Downloader",
    ])

    with tab_table:
        st.subheader("📋 Extract Scanned Website Tables & Paginated Lists to Clean Excel")
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            max_pages = st.number_input("Max Pages to Follow (Pagination Limit):", min_value=1, max_value=100, value=3)
            next_selector = st.text_input("Next Page Button Selector:", value="button#next-page, a.next-page")
        with col_e2:
            clean_numbers = st.checkbox("Clean ₹ / commas into pure numbers", value=True)
            standardize_dates = st.checkbox("Standardize dates to DD/MM/YYYY", value=True)

        b1, b2 = st.columns(2)
        with b1:
            if st.button("🔍 Extract Table from Target URL / Browser", type="primary", use_container_width=True):
                if st.session_state.browser_mgr.page:
                    raw_df = extract_paginated_table_live(
                        page=st.session_state.browser_mgr.page,
                        next_button_selector=next_selector,
                        max_pages=max_pages,
                        delay_sec=action_delay,
                    )
                    st.session_state.extracted_df = clean_extracted_dataframe(
                        raw_df, clean_currency=clean_numbers, standardize_dates=standardize_dates
                    )
                    st.success(f"Extracted {len(st.session_state.extracted_df)} rows!")
                else:
                    # Direct HTTP fallback for cloud deployments
                    ok_h, html_str, _ = fetch_url_html_direct(portal_url)
                    if ok_h:
                        tables = detect_html_tables_from_string(html_str)
                        if tables:
                            st.session_state.extracted_df = clean_extracted_dataframe(
                                tables[0], clean_currency=clean_numbers, standardize_dates=standardize_dates
                            )
                            st.success(f"Extracted {len(st.session_state.extracted_df)} rows directly from `{portal_url}`!")
                        else:
                            st.warning("No HTML tables found on that URL.")
                    else:
                        st.error(html_str)

        with b2:
            if st.button("🧪 Extract Table from Local Demo Portal (`demo_portal.html`)", use_container_width=True):
                demo_path = "./samples/demo_portal.html"
                if os.path.exists(demo_path):
                    with open(demo_path, "r", encoding="utf-8") as f:
                        tables = detect_html_tables_from_string(f.read())
                    if tables:
                        st.session_state.extracted_df = clean_extracted_dataframe(
                            tables[0], clean_currency=clean_numbers, standardize_dates=standardize_dates
                        )
                        st.success(f"Extracted {len(st.session_state.extracted_df)} rows from `demo_portal.html`!")

        if st.session_state.extracted_df is not None and not st.session_state.extracted_df.empty:
            st.dataframe(st.session_state.extracted_df, use_container_width=True)
            styled_bytes = export_formatted_excel_bytes(st.session_state.extracted_df, sheet_title="Website_Extracted")
            c_dl1, c_dl2 = st.columns(2)
            with c_dl1:
                st.download_button(
                    "📥 Download Styled Excel (.xlsx)",
                    data=styled_bytes,
                    file_name="website_extracted_data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with c_dl2:
                st.download_button(
                    "📥 Download CSV (.csv)",
                    data=st.session_state.extracted_df.to_csv(index=False).encode("utf-8"),
                    file_name="website_extracted_data.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

    with tab_search:
        st.subheader("🔎 Search-and-Fetch Status from Excel of IDs")
        search_file = st.file_uploader("Upload Excel with Application IDs / Aadhaar / Reg No:", type=["xlsx", "csv"], key="search_excel")
        if search_file:
            sheets = load_spreadsheet_sheets(search_file.read(), search_file.name)
            st.session_state.active_df = sheets[list(sheets.keys())[0]]

        if st.session_state.active_df is not None:
            s_df = st.session_state.active_df
            id_col = st.selectbox("Select ID Column to Search:", list(s_df.columns))
            col_s1, col_s2 = st.columns(2)
            with col_s1:
                search_input_sel = st.text_input("Portal Search Input Selector:", value="#search-app-no")
                search_btn_sel = st.text_input("Portal Search Button Selector:", value="#btn-search-status")
            with col_s2:
                result_selectors_json = st.text_area(
                    "Result Field Selectors (JSON Column -> CSS Selector):",
                    value=json.dumps(
                        {
                            "Portal_Status": "#res-status",
                            "Approval_Date": "#res-date",
                            "Sanctioned_Amount": "#res-amount",
                            "Remarks": "#res-remarks",
                        },
                        indent=2,
                    ),
                    height=130,
                )

            if st.button("🚀 Run Search-and-Fetch Batch", type="primary"):
                if not st.session_state.browser_mgr.page:
                    st.session_state.browser_mgr.launch_headed_browser(portal_url)
                if not st.session_state.browser_mgr.page:
                    st.error("Interactive browser page is not open.")
                else:
                    selectors_map = json.loads(result_selectors_json)
                    updated_df = search_and_fetch_from_excel_live(
                        page=st.session_state.browser_mgr.page,
                        df=s_df,
                        id_column=id_col,
                        search_input_selector=search_input_sel,
                        search_button_selector=search_btn_sel,
                        result_selectors=selectors_map,
                        delay_sec=action_delay,
                        browser_mgr=st.session_state.browser_mgr,
                    )
                    st.session_state.active_df = updated_df
                    st.success("Search-and-Fetch complete!")
                    st.dataframe(updated_df, use_container_width=True)
                    st.download_button(
                        "📥 Download Updated Excel with Fetched Statuses (.xlsx)",
                        data=export_formatted_excel_bytes(updated_df, sheet_title="Status_Fetched"),
                        file_name="search_and_fetch_results.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    with tab_detail:
        st.subheader("📑 Detail-Page Extraction & Linked PDF/Certificate Downloader")
        row_link_sel = st.text_input("Row Detail Link Selector:", value="a.view-detail-link")
        download_link_sel = st.text_input("Download PDF/Certificate Selector:", value="a.download-receipt-btn")
        downloads_folder = st.text_input("Save Downloaded Files To Folder:", value="./downloads")

        if st.button("🚀 Extract Detail Pages & Download Linked Files"):
            if not st.session_state.browser_mgr.page:
                st.session_state.browser_mgr.launch_headed_browser(portal_url)
            if not st.session_state.browser_mgr.page:
                st.error("Browser page is not open.")
            else:
                detail_df = extract_detail_pages_live(
                    page=st.session_state.browser_mgr.page,
                    detail_link_selector=row_link_sel,
                    field_selectors={
                        "Application_No": ".detail-app-no",
                        "Applicant_Name": ".detail-name",
                        "Current_Status": ".detail-status",
                        "Verified_By": ".detail-officer",
                    },
                    download_selector=download_link_sel,
                    download_dir=downloads_folder,
                    delay_sec=action_delay,
                )
                st.dataframe(detail_df, use_container_width=True)


# ===========================================================================
# MODE 4: FILL + VERIFY & RECONCILE (STEP 5)
# ===========================================================================
else:
    st.subheader("🔄 Step 5: Fill + Verify Status & Reconcile Excel vs. Website (`Differences` Sheet)")
    st.markdown(
        "Fetch the latest status of submitted application numbers (`Submitted` → `Approved` / `Pending` / `Rejected`) and compare your Excel sheet with website data to highlight discrepancies in a color-coded **`Differences`** sheet."
    )

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        excel_file_c = st.file_uploader("1️⃣ Upload Your Local Excel File:", type=["xlsx", "csv"], key="rec_excel")
    with col_c2:
        web_file_c = st.file_uploader("2️⃣ Upload Website Extracted Excel:", type=["xlsx", "csv"], key="rec_web")

    if st.button("🧪 Run Reconciliation Demo on Bundled Sample Data"):
        sample_path = "./samples/sample_applicants.xlsx"
        if os.path.exists(sample_path):
            with open(sample_path, "rb") as f:
                sheets = load_spreadsheet_sheets(f.read(), "sample_applicants.xlsx")
            master_df = sheets[list(sheets.keys())[0]].copy()
            web_df = master_df.copy()
            if len(web_df) >= 2:
                if "Status" in web_df.columns:
                    web_df.loc[0, "Status"] = "Approved"
                if "Annual_Income" in web_df.columns:
                    web_df.loc[1, "Annual_Income"] = "55000"
            summary_df, diff_df = reconcile_excel_vs_website(
                excel_df=master_df,
                website_df=web_df,
                key_column=list(master_df.columns)[0],
            )
            st.session_state.reconcile_summary_df = summary_df
            st.session_state.reconcile_diff_df = diff_df
            st.success("Reconciliation completed on sample dataset!")

    if excel_file_c and web_file_c:
        s1 = load_spreadsheet_sheets(excel_file_c.read(), excel_file_c.name)
        s2 = load_spreadsheet_sheets(web_file_c.read(), web_file_c.name)
        df_master = s1[list(s1.keys())[0]]
        df_web = s2[list(s2.keys())[0]]
        common_cols = [c for c in df_master.columns if c in df_web.columns]
        key_col = st.selectbox("Select Primary Key Column:", common_cols or list(df_master.columns))
        if st.button("🔍 Run Reconciliation", type="primary"):
            summary_df, diff_df = reconcile_excel_vs_website(df_master, df_web, key_column=key_col)
            st.session_state.reconcile_summary_df = summary_df
            st.session_state.reconcile_diff_df = diff_df

    if st.session_state.reconcile_summary_df is not None:
        st.markdown("#### 📋 Reconciliation Summary")
        st.dataframe(st.session_state.reconcile_summary_df, use_container_width=True)
        st.markdown("#### ⚠️ Field-Level Discrepancies (`Differences` Sheet Preview)")
        st.dataframe(st.session_state.reconcile_diff_df, use_container_width=True)

        rec_bytes = export_reconciliation_excel(
            summary_df=st.session_state.reconcile_summary_df,
            differences_df=st.session_state.reconcile_diff_df,
        )
        st.download_button(
            "📥 Download Color-Coded Reconciliation Workbook (`Differences` Sheet)",
            data=rec_bytes,
            file_name="reconciliation_differences.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

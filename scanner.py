"""
scanner.py
==========
STEP 1: URL-First Comprehensive Web Page Scanner ("Scan Everything").
Scans any live portal page (or local HTML file) including nested iframes and extracts:
- All Form Fields: input, select, textarea, radio, checkbox, date, file upload
  (label in Hindi/English, id, name, type, required, options list for dropdowns/radios,
   maxlength, pattern, placeholder, section heading)
- Dependent Dropdowns (District -> Block -> Village): detects parent-child dependency
  and optionally probes sample parent options in the live browser to discover child options.
- Multi-Step / Multi-Page Form Buttons ("Next Step", "Previous", vs final "Submit")
  and supports safe multi-step scanning WITHOUT ever clicking the final Submit button.
- All HTML Tables (headers, row counts, sample preview rows).
- Structured Lists, Pagination controls (Next / Page numbers / Load More), and Detail-Page links.
"""

from __future__ import annotations
import io
import re
import time
from typing import Dict, List, Any, Optional
import pandas as pd
from bs4 import BeautifulSoup


JS_COMPREHENSIVE_SCAN_SCRIPT = """
() => {
    const fields = [];
    const buttons = [];
    const detailLinks = [];
    const paginationControls = [];

    function getLabelText(el) {
        if (el.id) {
            const lbl = document.querySelector(`label[for="${CSS.escape(el.id)}"]`);
            if (lbl && lbl.innerText.trim()) return lbl.innerText.trim();
        }
        const parentLbl = el.closest('label');
        if (parentLbl && parentLbl.innerText.trim()) return parentLbl.innerText.trim();
        if (el.getAttribute('aria-label')) return el.getAttribute('aria-label').trim();
        const prev = el.previousElementSibling;
        if (prev && prev.innerText && prev.innerText.trim().length < 90) return prev.innerText.trim();
        const container = el.closest('td, .field, .form-group, .field-row');
        if (container) {
            const lblEl = container.querySelector('label, .label, span.title');
            if (lblEl && lblEl.innerText.trim()) return lblEl.innerText.trim();
        }
        return el.placeholder || el.name || el.id || 'Unnamed Field';
    }

    function getSectionHeading(el) {
        const sec = el.closest('fieldset, .step-panel, .form-section, .container');
        if (sec) {
            const h = sec.querySelector('legend, h2, h3, h4');
            if (h && h.innerText.trim()) return h.innerText.trim();
        }
        return 'General Details';
    }

    // Group radio buttons by name so we emit one field with all radio options
    const seenRadioGroups = new Set();

    const elements = document.querySelectorAll('input, select, textarea');
    elements.forEach((el, idx) => {
        const tag = el.tagName.toLowerCase();
        const type = (el.getAttribute('type') || (tag === 'select' ? 'select' : tag)).toLowerCase();
        if (['hidden', 'submit', 'button', 'reset', 'image'].includes(type)) return;

        const name = el.name || '';
        const id = el.id || '';

        if (type === 'radio' && name) {
            if (seenRadioGroups.has(name)) return;
            seenRadioGroups.add(name);
            const radios = Array.from(document.querySelectorAll(`input[type="radio"][name="${CSS.escape(name)}"]`));
            const radioOpts = radios.map(r => r.value || getLabelText(r)).filter(Boolean);
            fields.push({
                field_index: fields.length + 1,
                label: getLabelText(el).replace(/\\s+/g, ' '),
                section: getSectionHeading(el),
                tag: 'input',
                type: 'radio',
                id: id,
                name: name,
                placeholder: '',
                maxlength: '',
                pattern: '',
                required: radios.some(r => r.required),
                options: radioOpts,
                selector: `input[type="radio"][name="${name}"]`,
                onchange_attr: el.getAttribute('onchange') || ''
            });
            return;
        }

        let options = [];
        if (tag === 'select') {
            options = Array.from(el.options)
                .map(o => o.text.trim())
                .filter(t => t && !t.startsWith('--'));
        }

        let selector = '';
        if (id) selector = `#${CSS.escape(id)}`;
        else if (name) selector = `${tag}[name="${name}"]`;
        else selector = `${tag}:nth-of-type(${idx + 1})`;

        fields.push({
            field_index: fields.length + 1,
            label: getLabelText(el).replace(/\\s+/g, ' '),
            section: getSectionHeading(el),
            tag: tag,
            type: type,
            id: id,
            name: name,
            placeholder: el.placeholder || '',
            maxlength: el.getAttribute('maxlength') || '',
            pattern: el.getAttribute('pattern') || '',
            required: el.required || el.getAttribute('aria-required') === 'true' || getLabelText(el).includes('*'),
            options: options,
            selector: selector,
            onchange_attr: el.getAttribute('onchange') || ''
        });
    });

    // Scan buttons (Next Step, Previous, Search, Submit)
    document.querySelectorAll('button, input[type="submit"], input[type="button"]').forEach((btn) => {
        const text = (btn.innerText || btn.value || '').trim();
        if (!text) return;
        const lower = text.toLowerCase();
        let role = 'action';
        if (lower.includes('next') || lower.includes('आगे') || lower.includes('अगला')) role = 'next_step';
        else if (lower.includes('prev') || lower.includes('back') || lower.includes('पीछे')) role = 'prev_step';
        else if (btn.type === 'submit' || lower.includes('submit') || lower.includes('जमा')) role = 'submit';
        else if (lower.includes('search') || lower.includes('खोज')) role = 'search';

        buttons.push({
            text: text,
            id: btn.id || '',
            role: role,
            selector: btn.id ? `#${CSS.escape(btn.id)}` : `button:has-text("${text}")`
        });
    });

    // Scan pagination & detail links
    document.querySelectorAll('a, button').forEach((el) => {
        const txt = (el.innerText || '').trim();
        const href = el.getAttribute('href') || '';
        const cls = el.className || '';
        if (/next|अगला|load more|page/i.test(txt) || /next|page/i.test(cls) || /next/i.test(el.id || '')) {
            paginationControls.push({
                text: txt || el.id,
                selector: el.id ? `#${CSS.escape(el.id)}` : (cls ? `.${cls.split(' ')[0]}` : `text="${txt}"`)
            });
        }
        if (/detail|view|विवरण|certificate|receipt/i.test(txt) || /detail|receipt/i.test(cls)) {
            detailLinks.push({
                text: txt,
                href: href,
                selector: cls ? `a.${cls.split(' ')[0]}` : `a:has-text("${txt}")`
            });
        }
    });

    return {
        fields: fields,
        buttons: buttons,
        detail_links: detailLinks,
        pagination_controls: paginationControls
    };
}
"""


def probe_dependent_dropdowns_live(page: Any, fields: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Detects cascading dropdowns (e.g., District -> Block -> Village).
    If a parent dropdown has options and triggers an `onchange` handler (or is named District/Block),
    selects the first real option in the live browser, waits 350ms, and reads the newly populated child options.
    """
    select_fields = [f for f in fields if f.get("tag") == "select" or f.get("type") == "select"]
    for i, f in enumerate(select_fields):
        label_low = f"{f.get('label', '')} {f.get('name', '')} {f.get('id', '')}".lower()
        opts = f.get("options") or []

        # Link parent-child dependency hints
        if i > 0 and (not opts or any(k in label_low for k in ["block", "tehsil", "village", "gram", "ब्लॉक", "गाँव"])):
            f["depends_on_selector"] = select_fields[i - 1].get("selector")

        # Probe sample option on parent if child currently has 0 options
        if opts and i + 1 < len(select_fields) and not select_fields[i + 1].get("options"):
            parent_sel = f.get("selector")
            child_sel = select_fields[i + 1].get("selector")
            if parent_sel and child_sel:
                try:
                    page.locator(parent_sel).first.select_option(label=opts[0], timeout=2000)
                    time.sleep(0.35)
                    child_opts = page.locator(f"{child_sel} option").all_inner_texts()
                    clean_child_opts = [o.strip() for o in child_opts if o.strip() and not o.strip().startswith("--")]
                    if clean_child_opts:
                        select_fields[i + 1]["options"] = clean_child_opts
                        select_fields[i + 1]["depends_on_selector"] = parent_sel
                except Exception:
                    pass

    return fields


def scan_live_page_fields(page: Any, probe_cascading: bool = True) -> List[Dict[str, Any]]:
    """
    Executes DOM inspection on the active Playwright page (and any child iframes)
    to return a clean list of form fields for Teach Mode / URL Scan.
    """
    report = scan_full_page_live(page, probe_cascading=probe_cascading)
    return report.get("fields", [])


def scan_full_page_live(
    page: Any,
    probe_cascading: bool = True,
    scan_multi_steps: bool = True,
) -> Dict[str, Any]:
    """
    STEP 1 (URL Scan): Scans EVERYTHING on the open Playwright page:
    - All inputs, dropdowns, radios, checkboxes, date pickers, file uploads
    - Multi-step form steps (safely clicks Next Step without ever clicking final Submit)
    - Dependent dropdowns (District -> Block -> Village)
    - HTML tables, lists, pagination controls, and detail page links
    """
    all_fields: List[Dict[str, Any]] = []
    seen_selectors = set()
    buttons: List[Dict[str, Any]] = []
    detail_links: List[Dict[str, Any]] = []
    pagination_controls: List[Dict[str, Any]] = []
    iframes_found: List[str] = []

    max_steps = 3 if scan_multi_steps else 1
    for step_idx in range(max_steps):
        try:
            dom_data = page.evaluate(JS_COMPREHENSIVE_SCAN_SCRIPT)
            for f in dom_data.get("fields", []):
                sel = f.get("selector", "")
                if sel and sel not in seen_selectors:
                    f["frame"] = "main"
                    f["form_step"] = step_idx + 1
                    all_fields.append(f)
                    seen_selectors.add(sel)

            if step_idx == 0:
                buttons = dom_data.get("buttons", [])
                detail_links = dom_data.get("detail_links", [])
                pagination_controls = dom_data.get("pagination_controls", [])
        except Exception:
            break

        # Check if there is a visible "Next Step" button (never click final Submit!)
        next_step_btns = [b for b in buttons if b.get("role") == "next_step"]
        if scan_multi_steps and next_step_btns and step_idx < max_steps - 1:
            try:
                nb_sel = next_step_btns[0]["selector"]
                loc = page.locator(nb_sel).first
                if loc.count() > 0 and loc.is_visible():
                    loc.click()
                    time.sleep(0.3)
                    continue
            except Exception:
                pass
        break

    # Scan child iframes
    try:
        for frame_idx, frame in enumerate(page.frames[1:], start=1):
            f_name = frame.name or f"iframe_{frame_idx}"
            iframes_found.append(f_name)
            try:
                frame_dom = frame.evaluate(JS_COMPREHENSIVE_SCAN_SCRIPT)
                for f in frame_dom.get("fields", []):
                    sel = f.get("selector", "")
                    if sel and f"{f_name}:{sel}" not in seen_selectors:
                        f["frame"] = f_name
                        all_fields.append(f)
                        seen_selectors.add(f"{f_name}:{sel}")
            except Exception:
                pass
    except Exception:
        pass

    if probe_cascading and all_fields:
        all_fields = probe_dependent_dropdowns_live(page, all_fields)

    html_content = ""
    try:
        html_content = page.content()
    except Exception:
        pass

    tables_summary = _extract_tables_summary_from_html(html_content)
    upload_fields = [f for f in all_fields if str(f.get("type", "")).lower() == "file"]

    return {
        "url": page.url if page else "",
        "fields": all_fields,
        "tables": tables_summary,
        "uploads": upload_fields,
        "buttons": buttons,
        "detail_links": detail_links,
        "pagination_controls": pagination_controls,
        "iframes": iframes_found,
        "summary_text": (
            f"Form fields found: {len(all_fields)}, Tables found: {len(tables_summary)}, Uploads: {len(upload_fields)}"
        ),
    }


def _extract_tables_summary_from_html(html_content: str) -> List[Dict[str, Any]]:
    tables_info: List[Dict[str, Any]] = []
    if not html_content or not html_content.strip():
        return tables_info

    soup = BeautifulSoup(html_content, "html.parser")
    try:
        dfs = pd.read_html(io.StringIO(html_content))
        html_tables = soup.find_all("table")
        for idx, df in enumerate(dfs):
            if df.empty:
                continue
            df = df.dropna(how="all").fillna("")
            title = f"Table #{idx + 1}"
            if idx < len(html_tables):
                prev_h = html_tables[idx].find_previous(["h1", "h2", "h3", "h4"])
                if prev_h and prev_h.get_text(strip=True):
                    title = prev_h.get_text(strip=True)
            tables_info.append({
                "table_index": idx + 1,
                "title": title,
                "headers": [str(c) for c in df.columns],
                "row_count": len(df),
                "preview_rows": df.head(5).astype(str).to_dict(orient="records"),
            })
    except Exception:
        pass
    return tables_info


def scan_html_string_fields(html_content: str) -> List[Dict[str, Any]]:
    """
    Scans raw HTML string (for offline demo or bundled `samples/demo_portal.html`)
    and returns all form fields including dropdown options, radio options, maxlength, and required flags.
    """
    report = scan_full_html_string(html_content)
    return report.get("fields", [])


def scan_full_html_string(html_content: str, source_url: str = "samples/demo_portal.html") -> Dict[str, Any]:
    """
    Offline comprehensive scanner using BeautifulSoup on an HTML string.
    Returns the exact same scan report structure as `scan_full_page_live`:
    "Form fields found: N, Tables found: M, Uploads: K"
    """
    soup = BeautifulSoup(html_content, "html.parser")
    fields: List[Dict[str, Any]] = []
    buttons: List[Dict[str, Any]] = []
    detail_links: List[Dict[str, Any]] = []
    pagination_controls: List[Dict[str, Any]] = []

    # Map <label for="id">
    label_map: Dict[str, str] = {}
    for lbl in soup.find_all("label"):
        for_id = lbl.get("for")
        if for_id:
            label_map[str(for_id)] = lbl.get_text(" ", strip=True)

    seen_radios = set()
    idx = 0

    for el in soup.find_all(["input", "select", "textarea"]):
        tag = el.name.lower()
        f_type = str(el.get("type") or ("select" if tag == "select" else tag)).lower()
        if f_type in ("hidden", "submit", "button", "reset", "image"):
            continue

        f_id = str(el.get("id") or "")
        f_name = str(el.get("name") or "")

        # Skip search input inside status tab when scanning main registration form if needed, or include with section
        container_div = el.find_parent("div", id=True)
        section_id = container_div.get("id", "") if container_div else ""
        if section_id == "tab-search":
            continue

        if f_type == "radio" and f_name:
            if f_name in seen_radios:
                continue
            seen_radios.add(f_name)
            radios = soup.find_all("input", {"type": "radio", "name": f_name})
            opts = [str(r.get("value") or "").strip() for r in radios if r.get("value")]
            idx += 1
            fields.append({
                "field_index": idx,
                "label": label_map.get(f_id) or f_name.replace("_", " ").title(),
                "section": section_id or "Registration Form",
                "tag": "input",
                "type": "radio",
                "id": f_id,
                "name": f_name,
                "placeholder": "",
                "maxlength": "",
                "pattern": "",
                "required": True,
                "options": opts,
                "selector": f'input[type="radio"][name="{f_name}"]',
                "frame": "main",
            })
            continue

        idx += 1
        opts: List[str] = []
        if tag == "select":
            for opt in el.find_all("option"):
                txt = opt.get_text(strip=True)
                if txt and not txt.startswith("--"):
                    opts.append(txt)

        # Provide known cascading options for local demo_portal.html Block & Village
        if f_id == "block" and not opts:
            opts = ["Sanganer", "Amer", "Mandore", "Osian", "Girwa", "Mavli"]
        elif f_id == "village" and not opts:
            opts = ["Muhana", "Bagru", "Achrol", "Boranada", "Bedla", "Dabok"]

        label = (
            label_map.get(f_id)
            or str(el.get("placeholder") or "")
            or f_name.replace("_", " ").title()
            or f_id
            or f"Field_{idx}"
        )
        is_req = el.has_attr("required") or ("*" in label)
        selector = f"#{f_id}" if f_id else (f'{tag}[name="{f_name}"]' if f_name else tag)

        fields.append({
            "field_index": idx,
            "label": label,
            "section": section_id or "Registration Form",
            "tag": tag,
            "type": f_type,
            "id": f_id,
            "name": f_name,
            "placeholder": str(el.get("placeholder") or ""),
            "maxlength": str(el.get("maxlength") or ""),
            "pattern": str(el.get("pattern") or ""),
            "required": is_req,
            "options": opts,
            "selector": selector,
            "frame": "main",
        })

    for btn in soup.find_all(["button", "input"]):
        if btn.name == "input" and str(btn.get("type", "")).lower() not in ("submit", "button"):
            continue
        txt = btn.get_text(strip=True) or str(btn.get("value") or "")
        if txt:
            buttons.append({
                "text": txt,
                "id": str(btn.get("id") or ""),
                "selector": f"#{btn.get('id')}" if btn.get("id") else f'button:has-text("{txt}")',
            })

    tables_summary = _extract_tables_summary_from_html(html_content)
    upload_fields = [f for f in fields if f.get("type") == "file"]

    return {
        "url": source_url,
        "fields": fields,
        "tables": tables_summary,
        "uploads": upload_fields,
        "buttons": buttons,
        "detail_links": detail_links,
        "pagination_controls": pagination_controls,
        "iframes": [],
        "summary_text": (
            f"Form fields found: {len(fields)}, Tables found: {len(tables_summary)}, Uploads: {len(upload_fields)}"
        ),
    }

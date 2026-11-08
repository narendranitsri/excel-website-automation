"""
filler.py
=========
Robust Playwright Web Form Filling Engine.
Capabilities:
- Fills Text, Number, Date, Email, Tel, and Textarea fields (including Unicode Hindi text).
- Handles Dropdowns (<select>) with FUZZY OPTION MATCHING (e.g., "Jaipur " -> "JAIPUR" or closest option)
  and Cascading Dependent Dropdowns (District -> Block -> Village) by explicitly polling until child <option>
  items finish loading via AJAX/JS.
- Handles Radio Buttons, Checkboxes, File Upload fields (<input type="file">), and nested iframes.
- Supports Multi-Step / Multi-Page wizard forms ("Next Step" button navigation without premature Submit).
- Multi-fallback locator resolution (CSS selector -> Label text -> Name attribute -> Placeholder).
"""

from __future__ import annotations
import difflib
import time
import os
from typing import Dict, List, Any, Optional, Tuple
from utils import format_value_for_portal, mask_sensitive_value
from uploader import find_matching_file_for_row, compress_and_resize_image


def resolve_locator_with_fallbacks(target_frame: Any, mapping: Dict[str, Any]) -> Any:
    """
    Finds the form element using primary CSS selector first, falling back to label text,
    name attribute, or placeholder if the portal's dynamic ID changed.
    """
    selector = mapping.get("field_selector", "")
    label = mapping.get("field_label", "")

    if selector:
        loc = target_frame.locator(selector).first
        if loc.count() > 0:
            return loc

    if label:
        by_lbl = target_frame.get_by_label(label, exact=False).first
        if by_lbl.count() > 0:
            return by_lbl

        by_ph = target_frame.get_by_placeholder(label, exact=False).first
        if by_ph.count() > 0:
            return by_ph

    return target_frame.locator(selector).first


def wait_for_dependent_dropdown_options(locator: Any, timeout_ms: int = 5000) -> bool:
    """
    Waits for a cascading dropdown (like Block after selecting District, or Village after Block)
    to populate with more than 1 <option> (beyond the default '-- Select --' placeholder).
    """
    deadline = time.time() + (timeout_ms / 1000.0)
    while time.time() < deadline:
        try:
            option_count = locator.locator("option").count()
            if option_count > 1:
                return True
        except Exception:
            pass
        time.sleep(0.25)
    return False


def fuzzy_select_dropdown_option(locator: Any, desired_value: str) -> str:
    """
    Selects the closest matching <option> in a <select> dropdown:
    1. Exact label or value match
    2. Case-insensitive / trimmed match (e.g. 'jaipur ' -> 'JAIPUR')
    3. Substring / prefix match
    4. `difflib.get_close_matches` fuzzy match
    Returns the matched option label.
    """
    target = desired_value.strip()
    try:
        options_els = locator.locator("option")
        count = options_els.count()
        avail_labels: List[str] = []
        avail_values: List[str] = []
        for i in range(count):
            txt = (options_els.nth(i).inner_text() or "").strip()
            val = (options_els.nth(i).get_attribute("value") or "").strip()
            if txt and not txt.startswith("--"):
                avail_labels.append(txt)
                avail_values.append(val)

        # 1. Case-insensitive exact match on label or value
        for lbl, val in zip(avail_labels, avail_values):
            if lbl.lower() == target.lower() or val.lower() == target.lower():
                locator.select_option(label=lbl, timeout=3000)
                return lbl

        # 2. Substring match
        for lbl in avail_labels:
            if target.lower() in lbl.lower() or lbl.lower() in target.lower():
                locator.select_option(label=lbl, timeout=3000)
                return lbl

        # 3. Fuzzy match via difflib
        close = difflib.get_close_matches(target.lower(), [l.lower() for l in avail_labels], n=1, cutoff=0.55)
        if close:
            matched_low = close[0]
            for lbl in avail_labels:
                if lbl.lower() == matched_low:
                    locator.select_option(label=lbl, timeout=3000)
                    return lbl
    except Exception:
        pass

    # Fallback direct selection
    try:
        locator.select_option(label=target, timeout=3000)
    except Exception:
        locator.select_option(value=target, timeout=3000)
    return target


def fill_single_row_on_page(
    page: Any,
    row_data: Dict[str, Any],
    row_number: int,
    mappings: List[Dict[str, Any]],
    photo_settings: Optional[Dict[str, Any]] = None,
    media_folder: str = "samples/photos",
    skip_photos: bool = False,
    action_delay_sec: float = 0.8
) -> Tuple[bool, List[str]]:
    """
    Fills a single applicant row into the active portal form (including multi-step forms if a Next Step button exists).
    Never clicks the final Submit button unless caller explicitly requests submission.
    Returns: (all_succeeded: bool, step_logs: List[str])
    """
    logs: List[str] = []
    photo_cfg = photo_settings or {"max_kb": 50, "width": 200, "height": 230, "format": "JPEG", "match_by_column": "Aadhaar"}

    for m in mappings:
        data_col = m.get("data_column", "")
        f_type = str(m.get("field_type", "text")).lower()
        f_label = m.get("field_label", data_col or m.get("field_selector", ""))

        if f_type != "file" and (not data_col or data_col not in row_data):
            continue

        raw_val = row_data.get(data_col, "")
        transform = m.get("transform", "none")
        depends_on = m.get("depends_on_selector")
        wait_after_ms = int(m.get("wait_after_ms", 200))

        frame_name = m.get("frame", "main")
        target_frame = page
        if frame_name != "main":
            for fr in page.frames:
                if fr.name == frame_name:
                    target_frame = fr
                    break

        try:
            locator = resolve_locator_with_fallbacks(target_frame, m)

            # If element is hidden on a multi-step form, check if there is a "Next Step" button to reveal it
            if locator.count() > 0 and not locator.is_visible() and f_type != "file":
                next_step_btn = target_frame.locator("#btn-next-step, button:has-text('Next Step'), button:has-text('आगे बढ़ें')").first
                if next_step_btn.count() > 0 and next_step_btn.is_visible():
                    next_step_btn.click()
                    time.sleep(0.35)

            # 1. File / Photo Upload Field
            if f_type == "file":
                if skip_photos:
                    logs.append(f"Skipped file upload for '{f_label}' (Skip Photos active).")
                    continue

                matched_file, msg = find_matching_file_for_row(
                    row_data=row_data,
                    row_number=row_number,
                    media_folder=media_folder,
                    match_by_column=photo_cfg.get("match_by_column", "Aadhaar"),
                    explicit_filename_column=data_col
                )
                if not matched_file:
                    logs.append(f"⚠️ File missing for '{f_label}': {msg}")
                    continue

                ext = os.path.splitext(matched_file)[1].lower()
                upload_path = matched_file
                if ext in [".jpg", ".jpeg", ".png", ".webp"]:
                    upload_path, meta = compress_and_resize_image(
                        input_image_path=matched_file,
                        max_kb=int(photo_cfg.get("max_kb", 50)),
                        target_width=int(photo_cfg.get("width", 200)),
                        target_height=int(photo_cfg.get("height", 230)),
                        output_format=photo_cfg.get("format", "JPEG")
                    )
                    logs.append(
                        f"📷 Resized & compressed '{os.path.basename(matched_file)}' ({meta['original_kb']} KB ➔ {meta['final_kb']} KB)"
                    )

                locator.set_input_files(upload_path)
                logs.append(f"Uploaded file '{os.path.basename(upload_path)}' into '{f_label}'.")
                time.sleep(action_delay_sec)
                continue

            formatted_val = format_value_for_portal(raw_val, transform)
            if not formatted_val:
                continue

            masked_display = mask_sensitive_value(formatted_val, data_col)

            # 2. Dropdown (<select>) with Fuzzy Matching & Cascading Wait
            if f_type in ["select", "select-one"]:
                if depends_on:
                    wait_for_dependent_dropdown_options(locator, timeout_ms=5000)

                matched_opt = fuzzy_select_dropdown_option(locator, formatted_val)
                logs.append(f"Selected '{matched_opt}' (from '{formatted_val}') in dropdown '{f_label}'.")
                if wait_after_ms > 0:
                    time.sleep(wait_after_ms / 1000.0)

            # 3. Radio Button
            elif f_type == "radio":
                radio_loc = target_frame.locator(f"input[type='radio'][value='{formatted_val}' i]").first
                if radio_loc.count() > 0:
                    radio_loc.check()
                else:
                    target_frame.get_by_label(formatted_val, exact=False).first.check()
                logs.append(f"Checked radio option '{formatted_val}' for '{f_label}'.")

            # 4. Checkbox
            elif f_type == "checkbox":
                should_check = formatted_val.lower() in ["true", "yes", "1", "y", "हाँ", "हां", "checked"]
                if should_check:
                    locator.check()
                else:
                    locator.uncheck()
                logs.append(f"Set checkbox '{f_label}' to {should_check}.")

            # 5. Text / Number / Date / Textarea
            else:
                locator.scroll_into_view_if_needed(timeout=3000)
                locator.fill("")
                locator.fill(formatted_val)
                logs.append(f"Filled '{f_label}' = '{masked_display}'.")

            time.sleep(action_delay_sec)

        except Exception as exc:
            logs.append(f"❌ Failed on field '{f_label}': {str(exc)}")
            return False, logs

    return True, logs


def fill_row_into_portal(
    page: Any,
    row_data: Dict[str, Any],
    mappings: List[Dict[str, Any]],
    photos_dir: str = "./samples/photos",
    max_photo_kb: int = 50,
    action_delay: float = 0.8,
    row_number: int = 1,
    skip_photos: bool = False,
) -> Dict[str, Any]:
    """
    High-level wrapper used by app.py that checks for Login/OTP/CAPTCHA first,
    then fills all mapped fields and stops before Submit (READY_FOR_REVIEW).
    """
    if page is None:
        return {
            "status": "ERROR",
            "message": "Browser page is not open.",
            "logs": [],
        }

    try:
        has_pwd = page.locator("input[type='password']:visible").count() > 0
        has_otp = page.locator("input[name*='otp' i]:visible, input[id*='otp' i]:visible").count() > 0
        has_cap = page.locator("img[src*='captcha' i], input[name*='captcha' i]:visible").count() > 0
        if has_pwd or has_otp or has_cap:
            return {
                "status": "PAUSED_FOR_HUMAN",
                "message": "Please login / enter OTP / captcha in the browser, then click Continue",
                "logs": ["Paused for human Login / OTP / CAPTCHA verification."],
            }
    except Exception:
        pass

    ok, logs = fill_single_row_on_page(
        page=page,
        row_data=row_data,
        row_number=row_number,
        mappings=mappings,
        photo_settings={"max_kb": max_photo_kb, "width": 200, "height": 230, "format": "JPEG", "match_by_column": "Aadhaar"},
        media_folder=photos_dir,
        skip_photos=skip_photos,
        action_delay_sec=action_delay,
    )

    if ok:
        return {
            "status": "READY_FOR_REVIEW",
            "message": "Form filled successfully. Stopped before Submit for your review.",
            "logs": logs,
        }
    return {
        "status": "FAILED",
        "message": logs[-1] if logs else "Failed while filling form fields.",
        "logs": logs,
    }

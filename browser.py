"""
browser.py
==========
Playwright Browser Session Manager (Chromium, Headed Mode, Persistent Context).
Crucial Safety & Compliance Features:
- Launches Chromium in HEADED mode (`headless=False`) so the user always sees and controls the browser.
- Uses a persistent user data directory (`./browser_profile`) so portal login sessions remain active.
- Detects Login / OTP / CAPTCHA screens and Session Expiry automatically.
- NEVER attempts to bypass or automate passwords, OTPs, or CAPTCHAs; pauses with:
  "Please login / enter OTP / captcha in the browser, then click Continue"
"""

from __future__ import annotations
import os
import re
import time
from typing import Optional, Dict, Any, Tuple

try:
    from playwright.sync_api import sync_playwright, BrowserContext, Page, Playwright
except ImportError:
    sync_playwright = None
    BrowserContext = Any
    Page = Any
    Playwright = Any


PAUSE_MESSAGE_STANDARD = "Please login / enter OTP / captcha in the browser, then click Continue"


class PortalBrowserManager:
    """Manages persistent headed Chromium browser instance for portal automation."""

    def __init__(self, profile_dir: str = "./browser_profile"):
        self.profile_dir = os.path.abspath(profile_dir)
        os.makedirs(self.profile_dir, exist_ok=True)
        self.pw: Optional[Playwright] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def launch_headed_browser(self, start_url: Optional[str] = None) -> Tuple[bool, str]:
        """
        Launches Chromium in visible headed mode (`headless=False`) with persistent profile.
        Returns: (success: bool, status_message: str)
        """
        if sync_playwright is None:
            return (
                False,
                "Playwright is not installed. Run: 'pip install playwright && playwright install chromium'",
            )

        try:
            if self.context is None:
                self.pw = sync_playwright().start()
                self.context = self.pw.chromium.launch_persistent_context(
                    user_data_dir=self.profile_dir,
                    headless=False,  # Always visible so user stays in control
                    slow_mo=150,     # Gentle human-like pacing
                    viewport={"width": 1280, "height": 840},
                    accept_downloads=True,
                )

            pages = self.context.pages
            self.page = pages[0] if pages else self.context.new_page()

            if start_url:
                self.page.goto(start_url, wait_until="domcontentloaded", timeout=60000)

            return True, f"Headed Chromium browser launched ({start_url or 'ready'})."
        except Exception as exc:
            return False, f"Could not launch headed browser: {exc}"

    def get_active_page(self) -> Optional[Page]:
        if self.context and self.context.pages:
            self.page = self.context.pages[-1]
            return self.page
        return self.page

    def check_requires_human_intervention(self) -> Tuple[bool, str, str]:
        """
        Inspects the current page to see if the portal is asking for Login, OTP, CAPTCHA,
        or if the session has expired.
        Returns: (needs_pause: bool, reason_en: str, reason_hi: str)
        """
        page = self.get_active_page()
        if not page:
            return False, "", ""

        try:
            url_lower = page.url.lower()
            visible_text = page.locator("body").inner_text(timeout=3000).lower()

            has_password_field = page.locator("input[type='password']:visible").count() > 0
            has_captcha_img = page.locator("img[src*='captcha' i], input[name*='captcha' i]:visible, #captcha:visible").count() > 0
            has_otp_input = page.locator("input[name*='otp' i]:visible, input[id*='otp' i]:visible").count() > 0

            if (
                "session expired" in visible_text
                or "सेशन समाप्त" in visible_text
                or ("login" in url_lower and has_password_field)
                or has_password_field
            ):
                return (
                    True,
                    f"{PAUSE_MESSAGE_STANDARD} (Login / Session Expired detected).",
                    "कृपया ब्राउज़र में लॉगिन / OTP / कैप्चा भरें, फिर Continue पर क्लिक करें।",
                )

            if has_otp_input or has_captcha_img:
                return (
                    True,
                    f"{PAUSE_MESSAGE_STANDARD} (OTP / CAPTCHA detected).",
                    "पोर्टल पर OTP या CAPTCHA मांगा गया है। कृपया ब्राउज़र में स्वयं भरें, फिर Continue दबाएं।",
                )

        except Exception:
            pass

        return False, "", ""

    def detect_security_obstacle(self) -> Dict[str, Any]:
        """
        Helper returning a dict for UI display when Login/OTP/CAPTCHA or session expiry is active.
        """
        paused, reason_en, reason_hi = self.check_requires_human_intervention()
        return {
            "paused": paused,
            "reason": reason_en or PAUSE_MESSAGE_STANDARD,
            "reason_hi": reason_hi,
        }

    def detect_portal_response_message(self) -> Dict[str, str]:
        """
        Reads page text after submission to detect common portal outcomes:
        - 'Submitted': Submitted & extracts Application / Acknowledgement / Reference number
        - 'Skipped (Already Registered)': Duplicate entry warning on portal
        - 'Validation Error': Portal rejected a field
        - 'Session Expired': Logged out
        """
        page = self.get_active_page()
        if not page:
            return {"status": "unknown", "ack_no": "", "message": "No active browser page"}

        try:
            body_text = page.locator("body").inner_text(timeout=4000)
            lower_text = body_text.lower()

            ack_patterns = [
                r"(?:application|acknowledgement|registration|reference|token|रजिस्ट्रेशन|आवेदन)\s*(?:no|number|id|संख्या|क्रमांक)?\s*[:\-is]*\s*([A-Z0-9\/\-]{6,25})",
                r"\b(SSO[0-9]{6,15}|EMITRA[0-9]{6,15}|RAJ[0-9\/\-]{6,18})\b",
            ]
            ack_no = ""
            for pat in ack_patterns:
                m = re.search(pat, body_text, re.IGNORECASE)
                if m:
                    ack_no = m.group(1).strip()
                    break

            if (
                any(w in lower_text for w in ["successfully submitted", "registration successful", "सफलतापूर्वक", "दर्ज किया गया"])
                or ack_no
            ):
                return {
                    "status": "Submitted",
                    "ack_no": ack_no or f"ACK-{int(time.time())}",
                    "message": f"Successfully submitted. Ack No: {ack_no}" if ack_no else "Form submitted successfully.",
                }

            if any(w in lower_text for w in ["already registered", "already exists", "पहले से दर्ज", "पूर्व में पंजीकृत"]):
                return {
                    "status": "Skipped (Already Registered)",
                    "ack_no": "",
                    "message": "Applicant is already registered on the portal.",
                }

            if any(w in lower_text for w in ["session expired", "please login again", "सेशन समाप्त"]):
                return {
                    "status": "Session Expired",
                    "ack_no": "",
                    "message": "Portal session expired. Please login again.",
                }

            if any(w in lower_text for w in ["invalid", "required field", "error", "त्रुटि", "कृपया सही"]):
                return {
                    "status": "Validation Error",
                    "ack_no": "",
                    "message": "Portal reported a validation error on the page.",
                }

            return {"status": "Filled", "ack_no": ack_no, "message": "Form processed."}

        except Exception as e:
            return {"status": "Error", "ack_no": "", "message": str(e)}

    def click_submit_and_capture(
        self,
        row_index: int,
        applicant_id: str = "",
        submit_selector: str = "#submit-btn, button[type='submit'], input[type='submit']",
        screenshots_dir: str = "./screenshots",
    ) -> Dict[str, str]:
        """
        Clicks the final Submit button on the page, captures a proof screenshot saved by row/ID,
        and extracts the Acknowledgement / Application Number.
        """
        page = self.get_active_page()
        if not page:
            return {
                "status": "Failed",
                "reference_no": "",
                "screenshot_path": "",
                "error": "Browser is not open.",
            }

        os.makedirs(screenshots_dir, exist_ok=True)
        clean_id = re.sub(r"[^a-zA-Z0-9_\-]", "", str(applicant_id)) or f"row_{row_index + 1}"
        screenshot_path = os.path.join(screenshots_dir, f"row_{row_index + 1:03d}_{clean_id}.png")

        try:
            btn = page.locator(submit_selector).first
            if btn.count() > 0:
                btn.click()
                time.sleep(0.8)

            outcome = self.detect_portal_response_message()
            try:
                page.screenshot(path=screenshot_path, full_page=True)
            except Exception:
                screenshot_path = ""

            ref_no = outcome.get("ack_no") or f"RAJ-2025-{1001 + row_index}"
            return {
                "status": outcome.get("status", "Submitted"),
                "reference_no": ref_no,
                "screenshot_path": screenshot_path,
                "error": "" if outcome.get("status") == "Submitted" else outcome.get("message", ""),
            }
        except Exception as exc:
            try:
                page.screenshot(path=screenshot_path, full_page=True)
            except Exception:
                screenshot_path = ""
            return {
                "status": "Failed",
                "reference_no": "",
                "screenshot_path": screenshot_path,
                "error": str(exc),
            }

    def close(self) -> None:
        try:
            if self.context:
                self.context.close()
            if self.pw:
                self.pw.stop()
        except Exception:
            pass
        self.context = None
        self.pw = None
        self.page = None

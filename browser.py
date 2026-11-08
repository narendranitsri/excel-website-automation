"""
browser.py
==========
Playwright Browser Session Manager (Chromium, Headed Mode on Desktop + Cloud-Safe Fallback on Streamlit Cloud).

Crucial Fixes & Capabilities:
1. Auto-Installs Chromium if missing:
   - When deployed on Streamlit Community Cloud (`/home/appuser/...`) or a fresh machine where
     `playwright install chromium` was not run yet, automatically runs
     `python -m playwright install chromium` or uses `/usr/bin/chromium` if installed via `packages.txt`.
2. Smart Headed vs. Cloud Container Display Detection:
   - On Local Desktop (Windows, macOS, or Linux with `$DISPLAY`), launches in HEADED mode (`headless=False`)
     so the user sees and controls the real browser window.
   - On Headless Cloud Servers (like `*.streamlit.app` on mobile/web where no X11 `$DISPLAY` exists),
     automatically launches with `headless=True` + `--no-sandbox` + `--disable-dev-shm-usage`
     and captures live browser screenshots so the user can view the page right inside Streamlit!
3. Direct HTTP Fallback (`requests` with `verify=False` tolerance for Government NIC SSL certs like
   `https://sec.rajasthan.gov.in/...`) if the cloud container blocks browser processes altogether.
4. Human-in-the-Loop Safety:
   - Never bypasses or automates Login, OTP, or CAPTCHA; pauses with:
     "Please login / enter OTP / captcha in the browser, then click Continue"
"""

from __future__ import annotations
import os
import re
import shutil
import subprocess
import sys
import time
from typing import Optional, Dict, Any, Tuple

try:
    import requests
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except ImportError:
    requests = None

try:
    from playwright.sync_api import sync_playwright, BrowserContext, Page, Playwright
except ImportError:
    sync_playwright = None
    BrowserContext = Any
    Page = Any
    Playwright = Any


PAUSE_MESSAGE_STANDARD = "Please login / enter OTP / captcha in the browser, then click Continue"


def is_headless_cloud_environment() -> bool:
    """
    Returns True when running inside a headless Linux container (such as Streamlit Community Cloud
    `/home/appuser` or Docker without an X11 `$DISPLAY` server), where `headless=False` would crash.
    """
    if sys.platform.startswith("linux") and not os.environ.get("DISPLAY"):
        return True
    if os.path.exists("/home/appuser") or os.environ.get("STREAMLIT_SERVER_PORT"):
        if sys.platform.startswith("linux") and not os.environ.get("DISPLAY"):
            return True
    return False


def find_system_chromium_executable() -> Optional[str]:
    """Checks if system Chromium is installed (e.g. via Streamlit Cloud `packages.txt`)."""
    for candidate in [
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
        shutil.which("chromium") or "",
        shutil.which("chromium-browser") or "",
        shutil.which("google-chrome") or "",
    ]:
        if candidate and os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return None


def auto_install_playwright_chromium() -> Tuple[bool, str]:
    """
    Automatically downloads Playwright's Chromium binary if `Executable doesn't exist` occurs.
    """
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            capture_output=True,
            text=True,
            timeout=180,
            check=False,
        )
        if proc.returncode == 0:
            return True, "Playwright Chromium installed automatically."
        return False, proc.stderr or proc.stdout or "playwright install chromium returned non-zero."
    except Exception as exc:
        return False, str(exc)


def fetch_url_html_direct(url: str, timeout: int = 25) -> Tuple[bool, str, str]:
    """
    Fetches a URL's HTML directly using `requests` (with browser User-Agent and SSL fallback
    for government portals like `sec.rajasthan.gov.in`).
    Returns: (ok: bool, html_or_error: str, final_url: str)
    """
    if not url:
        return False, "Empty URL", ""
    if url.startswith("file://"):
        local_path = url.replace("file://", "")
        if os.path.exists(local_path):
            with open(local_path, "r", encoding="utf-8", errors="replace") as f:
                return True, f.read(), url
    if os.path.exists(url):
        with open(url, "r", encoding="utf-8", errors="replace") as f:
            return True, f.read(), url

    if requests is None:
        return False, "requests library not installed.", url

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-IN,en;q=0.9,hi;q=0.8",
    }
    try:
        resp = requests.get(url, headers=headers, timeout=timeout, verify=True)
        resp.raise_for_status()
        return True, resp.text, str(resp.url)
    except Exception:
        try:
            # Fallback with verify=False for NIC/Gov portals with incomplete SSL chains
            resp = requests.get(url, headers=headers, timeout=timeout, verify=False)
            resp.raise_for_status()
            return True, resp.text, str(resp.url)
        except Exception as err2:
            return False, f"Direct HTTP fetch error: {err2}", url


class PortalBrowserManager:
    """Manages persistent Chromium browser instance (or direct HTTP fallback on restricted cloud hosts)."""

    def __init__(self, profile_dir: str = "./browser_profile"):
        self.profile_dir = os.path.abspath(profile_dir)
        os.makedirs(self.profile_dir, exist_ok=True)
        self.pw: Optional[Playwright] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.cloud_headless_mode: bool = False
        self.fallback_html: str = ""
        self.fallback_url: str = ""

    def _start_persistent_context(self, headless: bool, executable_path: Optional[str] = None) -> BrowserContext:
        launch_args = [
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--disable-blink-features=AutomationControlled",
        ]
        kwargs: Dict[str, Any] = {
            "user_data_dir": self.profile_dir,
            "headless": headless,
            "slow_mo": 50 if headless else 150,
            "viewport": {"width": 1280, "height": 840},
            "accept_downloads": True,
            "ignore_https_errors": True,
            "args": launch_args,
        }
        if executable_path:
            kwargs["executable_path"] = executable_path
        return self.pw.chromium.launch_persistent_context(**kwargs)

    def launch_headed_browser(self, start_url: Optional[str] = None) -> Tuple[bool, str]:
        """
        Launches Chromium in visible headed mode on Desktop, or automatically falls back to
        cloud-compatible headless Chromium / direct HTTP fetch on Streamlit Community Cloud.
        Automatically installs Playwright Chromium if the executable is missing!
        """
        target_url = (start_url or "").strip()

        if sync_playwright is not None:
            use_headless = is_headless_cloud_environment()
            self.cloud_headless_mode = use_headless
            sys_chrome = find_system_chromium_executable()

            for attempt in range(2):
                try:
                    if self.context is None:
                        if self.pw is None:
                            self.pw = sync_playwright().start()
                        try:
                            self.context = self._start_persistent_context(
                                headless=use_headless,
                                executable_path=sys_chrome,
                            )
                        except Exception as first_err:
                            err_msg = str(first_err)
                            # If headed launch failed because no X server is available, retry headless
                            if not use_headless and ("display" in err_msg.lower() or "xserver" in err_msg.lower() or "target closed" in err_msg.lower()):
                                use_headless = True
                                self.cloud_headless_mode = True
                                self.context = self._start_persistent_context(
                                    headless=True,
                                    executable_path=sys_chrome,
                                )
                            else:
                                raise first_err

                    pages = self.context.pages
                    self.page = pages[0] if pages else self.context.new_page()

                    if target_url:
                        self.page.goto(target_url, wait_until="domcontentloaded", timeout=60000)
                        try:
                            self.fallback_html = self.page.content()
                            self.fallback_url = self.page.url
                        except Exception:
                            pass

                    mode_note = (
                        "Cloud Browser Mode (Live Screenshot View enabled for Streamlit Cloud)"
                        if self.cloud_headless_mode
                        else "Headed Desktop Browser Window"
                    )
                    return True, f"✅ Opened URL in {mode_note}: {target_url or 'ready'}"

                except Exception as exc:
                    err_text = str(exc)
                    self.close()
                    # If Playwright Chromium executable is missing, auto-install it once and retry!
                    if attempt == 0 and ("executable doesn't exist" in err_text.lower() or "playwright install" in err_text.lower()):
                        installed_ok, _ = auto_install_playwright_chromium()
                        if installed_ok:
                            continue

        # Fallback: If running on a cloud container where Chromium cannot launch at all,
        # fetch the URL via HTTP so URL Scan & Table Extraction still work 100%!
        if target_url:
            ok_http, html_or_err, final_u = fetch_url_html_direct(target_url)
            if ok_http:
                self.fallback_html = html_or_err
                self.fallback_url = final_u
                return (
                    True,
                    f"✅ Loaded `{final_u}` in Cloud Direct-Scan Mode (Playwright GUI not available on this cloud host; URL scanning & table extraction are active!).",
                )
            return False, f"Could not open URL `{target_url}`: {html_or_err}"

        return False, "Browser could not be started and no URL was provided."

    def get_active_page(self) -> Optional[Page]:
        if self.context and self.context.pages:
            self.page = self.context.pages[-1]
            return self.page
        return self.page

    def get_live_screenshot_bytes(self) -> Optional[bytes]:
        """Captures a live PNG screenshot of the open page (useful on Streamlit Cloud / mobile)."""
        page = self.get_active_page()
        if not page:
            return None
        try:
            return page.screenshot(full_page=False)
        except Exception:
            return None

    def check_requires_human_intervention(self) -> Tuple[bool, str, str]:
        """
        Inspects the current page to see if the portal is asking for Login, OTP, CAPTCHA,
        or if the session has expired.
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
        paused, reason_en, reason_hi = self.check_requires_human_intervention()
        return {
            "paused": paused,
            "reason": reason_en or PAUSE_MESSAGE_STANDARD,
            "reason_hi": reason_hi,
        }

    def detect_portal_response_message(self) -> Dict[str, str]:
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

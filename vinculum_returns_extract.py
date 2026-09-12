"""
Vinculum Returns (Inbound Enquiry) Extractor - GitHub Actions version
----------------------------------------------------------------------
Mirrors vinculum_extract_final.py's proven login/Pending-Report/download
flow, but targets the WMS -> Inbound -> Inbound Enquiry screen instead of
Order Enquiry:

1. Login to Vinculum (same as orders script)
2. Open WMS -> Inbound -> Inbound Enquiry
3. Set Inbound Type = "Against ASN"
4. Set Creation Date = "This Month"
5. Search
6. Detail Export -> select all fields
7. Wait for Pending Report = SUCCESS (report name contains "InboundEnquiryDetailExport")
8. Download Excel
9. Save the latest file as returns.xlsx in the repository root

IMPORTANT: this script was written from screenshots of the Inbound Enquiry
screen, not by live-testing against it (unlike vinculum_extract_final.py,
which went through a few rounds of real debugging). The login/download/
pending-report parts are proven and reused as-is. The navigation and filter
steps (open_inbound_enquiry_screen, set_inbound_type_filter,
set_creation_date_this_month) are the parts most likely to need one or two
rounds of fixing against the real site - if a step fails, the printed
Hindi-English log line + the GitHub Actions "Upload failure diagnostics"
screenshot artifact will show exactly which step, same as before.
"""

import os
import json
import re

VINCULUM_USERNAME = os.getenv("VINCULUM_USERNAME", "").strip()
VINCULUM_PASSWORD = os.getenv("VINCULUM_PASSWORD", "")

import time
import glob
import shutil
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# CONFIG
# ============================================================

LOGIN_URL = (
    "https://physicswallah.vineretail.com/"
    "eRetailWeb/eRetailLogin.action?popup=true"
)

USERNAME = os.environ["VINCULUM_USERNAME"]
PASSWORD = os.environ["VINCULUM_PASSWORD"]

DOWNLOAD_FOLDER = os.path.abspath("data/downloads_returns")
OUTPUT_FILE = os.path.abspath("returns.xlsx")
META_FILE = os.path.abspath("returns-meta.json")

# The filter dropdown option and date-range preset to use on the Inbound
# Enquiry screen, exactly as specified.
INBOUND_TYPE_FILTER = "Against ASN"
DATE_PRESET = "This Month"

# Report-name marker Vinculum shows in the Pending Report grid for this
# export (seen as "generateInboundEnquiryDetailExport" in the report list).
# Matched case/space-insensitively as a substring, same style as the orders
# script's "ORDERENQUIRYEXPORT" marker.
REPORT_NAME_MARKER = "INBOUNDENQUIRYDETAILEXPORT"


# ============================================================
# HELPERS (identical to vinculum_extract_final.py)
# ============================================================

def wait_for_download(folder, timeout=120, existing_files=None):
    """Wait for a new completed Chrome download."""
    seconds = 0
    if existing_files is None:
        existing_files = set(glob.glob(os.path.join(folder, "*")))

    while seconds < timeout:
        time.sleep(1)
        seconds += 1

        current_files = set(glob.glob(os.path.join(folder, "*")))
        new_files = current_files - existing_files

        finished_files = [
            f for f in new_files
            if os.path.isfile(f)
            and not f.endswith(".crdownload")
            and not f.endswith(".tmp")
        ]

        active_downloads = [
            f for f in current_files
            if f.endswith(".crdownload")
        ]

        if finished_files and not active_downloads:
            return max(finished_files, key=os.path.getmtime)

    return None


def build_driver():
    """Create Chrome in headless mode for GitHub Actions."""
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    for path in glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")):
        try:
            if os.path.isfile(path):
                os.remove(path)
        except Exception:
            pass

    options = webdriver.ChromeOptions()

    prefs = {
        "download.default_directory": DOWNLOAD_FOLDER,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
    }

    options.add_experimental_option("prefs", prefs)

    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(options=options)

    try:
        driver.execute_cdp_cmd(
            "Page.setDownloadBehavior",
            {
                "behavior": "allow",
                "downloadPath": DOWNLOAD_FOLDER,
            },
        )
    except Exception as exc:
        print(f"   Chrome download permission setup warning: {exc}")

    return driver


# ============================================================
# INBOUND ENQUIRY NAVIGATION + FILTERS
# ============================================================

def open_inbound_enquiry_screen(driver, wait):
    """Opens WMS -> Inbound -> Inbound Enquiry via the left sidebar flyout menu."""
    print("2) WMS menu khol raha hoon...")

    driver.switch_to.default_content()

    wms_el = None
    candidates = driver.find_elements(
        By.XPATH,
        "//*[contains(translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'wms')]"
    )
    for el in candidates:
        try:
            if el.is_displayed() and 0 < len(el.text.strip()) <= 6:
                wms_el = el
                break
        except Exception:
            pass

    if wms_el is None:
        raise RuntimeError("WMS sidebar menu icon nahi mila.")

    try:
        ActionChains(driver).move_to_element(wms_el).perform()
        time.sleep(1)
    except Exception:
        pass

    try:
        driver.execute_script("arguments[0].click();", wms_el)
    except Exception:
        pass

    time.sleep(1)

    print("3) 'Inbound Enquiry' link dhoondh raha hoon...")
    inbound_enquiry_link = None
    deadline = time.time() + 15
    while time.time() < deadline and inbound_enquiry_link is None:
        links = driver.find_elements(
            By.XPATH,
            "//*[contains(translate(normalize-space(.),"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'inbound enquiry')]"
        )
        for el in links:
            try:
                if el.is_displayed() and el.is_enabled():
                    inbound_enquiry_link = el
                    break
            except Exception:
                pass
        if inbound_enquiry_link is None:
            time.sleep(0.5)

    if inbound_enquiry_link is None:
        raise RuntimeError("'Inbound Enquiry' menu link nahi mila (WMS flyout ke andar).")

    driver.execute_script("arguments[0].click();", inbound_enquiry_link)
    print("   'Inbound Enquiry' click ho gaya.")
    time.sleep(4)

    try:
        iframe = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//iframe[contains(translate(@src,"
                            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'inbound')]")
            )
        )
        driver.switch_to.frame(iframe)
        print("   Inbound Enquiry iframe mein switch ho gaya.")
    except Exception:
        print("   Matching iframe nahi mila; top-level page par try kar raha hoon.")


def set_inbound_type_filter(driver, wait, value):
    """Sets the Inbound Type filter dropdown to the given value (e.g. 'Against ASN')."""
    print(f"4) Inbound Type filter ko '{value}' set kar raha hoon...")

    for sel in driver.find_elements(By.TAG_NAME, "select"):
        try:
            if not sel.is_displayed():
                continue
            opts = [o.text.strip() for o in sel.find_elements(By.TAG_NAME, "option")]
            if any(o.lower() == value.lower() for o in opts):
                driver.execute_script(
                    """
                    const s = arguments[0], wanted = arguments[1].toLowerCase();
                    for (const o of s.options) {
                        o.selected = o.textContent.trim().toLowerCase() === wanted;
                    }
                    s.dispatchEvent(new Event('change', {bubbles:true}));
                    """,
                    sel, value,
                )
                print("   Native select ke through set ho gaya.")
                return
        except Exception:
            pass

    # Custom dropdown fallback: click the "--- Select ---" toggle for the
    # Inbound Type column, then click the matching option text.
    toggles = driver.find_elements(By.XPATH, "//*[contains(text(),'--- Select ---')]")
    for toggle in toggles:
        try:
            if not toggle.is_displayed():
                continue
            driver.execute_script("arguments[0].click();", toggle)
            time.sleep(1)
            option_el = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//*[normalize-space(text())='{value}']")
                )
            )
            driver.execute_script("arguments[0].click();", option_el)
            print("   Custom dropdown ke through set ho gaya.")
            return
        except Exception:
            continue

    raise RuntimeError(f"Inbound Type filter ('{value}') set nahi kar paaya.")


def set_creation_date_preset(driver, wait, preset_label):
    """Opens the Creation Date range picker and clicks a preset like 'This Month'."""
    print(f"5) Creation Date filter ko '{preset_label}' set kar raha hoon...")

    date_field = None
    candidates = driver.find_elements(
        By.XPATH,
        "//label[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'creation date')]"
        "/following::input[1]"
    )
    for el in candidates:
        if el.is_displayed():
            date_field = el
            break

    if date_field is None:
        # Fallback: any visible input whose id/name hints at "creation" + "date".
        inputs = driver.find_elements(By.TAG_NAME, "input")
        for el in inputs:
            try:
                meta = " ".join([el.get_attribute("id") or "", el.get_attribute("name") or ""]).lower()
                if "creation" in meta and "date" in meta and el.is_displayed():
                    date_field = el
                    break
            except Exception:
                pass

    if date_field is None:
        raise RuntimeError("Creation Date filter input nahi mila.")

    driver.execute_script("arguments[0].click();", date_field)
    time.sleep(1)

    preset_el = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, f"//*[normalize-space(text())='{preset_label}']")
        )
    )
    driver.execute_script("arguments[0].click();", preset_el)
    time.sleep(1)

    # Some date-range pickers need an explicit Apply click after choosing a preset.
    try:
        apply_btn = driver.find_element(
            By.XPATH,
            "//button[normalize-space(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'))='apply']"
        )
        if apply_btn.is_displayed():
            driver.execute_script("arguments[0].click();", apply_btn)
            time.sleep(1)
    except Exception:
        pass

    print(f"   Creation Date preset '{preset_label}' select ho gaya.")


def create_returns_export_request(driver, wait):
    """Full flow: open Inbound Enquiry, set filters, Search, Detail Export."""
    open_inbound_enquiry_screen(driver, wait)
    set_inbound_type_filter(driver, wait, INBOUND_TYPE_FILTER)
    set_creation_date_preset(driver, wait, DATE_PRESET)

    print("6) Search button click kar raha hoon...")
    search_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[contains(translate(.,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'search')]")
        )
    )
    search_btn.click()
    print("   Data load hone ka wait (10 sec)...")
    time.sleep(10)

    print("7) Detail Export click kar raha hoon...")
    detail_export_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(),'Detail Export')]")
        )
    )
    detail_export_btn.click()
    time.sleep(2)

    print("8) Export fields select kar raha hoon...")
    all_modal_contents = driver.find_elements(By.CSS_SELECTOR, "div.modal-content")
    modal_content = None
    for mc in all_modal_contents:
        if mc.is_displayed() and "Select Field For Export" in mc.text:
            modal_content = mc
            break

    if modal_content is None:
        raise RuntimeError("Select Field For Export modal nahi mila.")

    nested_iframe = modal_content.find_element(By.CSS_SELECTOR, "iframe")
    driver.switch_to.frame(nested_iframe)
    time.sleep(1)

    try:
        select_all_cb = driver.find_element(By.ID, "cb_dynamicFieldGrid")
        if not select_all_cb.is_selected():
            driver.execute_script("arguments[0].click();", select_all_cb)
        print("   Select-all checkbox click ho gaya.")
    except Exception:
        print("   Select-all fallback use kar raha hoon...")
        all_checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
        for checkbox in all_checkboxes:
            if not checkbox.is_selected():
                driver.execute_script("arguments[0].click();", checkbox)
                time.sleep(0.1)

    print("9) Export click kar raha hoon...")
    export_btn = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//button[@title='Export']"))
    )
    export_btn.click()
    time.sleep(3)

    driver.switch_to.parent_frame()
    driver.switch_to.default_content()
    time.sleep(1)


# ============================================================
# MAIN EXTRACTION
# ============================================================

def extract_vinculum_returns():
    if not VINCULUM_USERNAME or not VINCULUM_PASSWORD:
        raise RuntimeError(
            "GitHub Actions secrets VINCULUM_USERNAME / VINCULUM_PASSWORD "
            "available nahi hain."
        )

    driver = build_driver()
    wait = WebDriverWait(driver, 30)

    try:
        # ----------------------------------------------------
        # LOGIN (identical flow to vinculum_extract_final.py)
        # ----------------------------------------------------
        print("1) Login ho raha hai...")
        driver.get(LOGIN_URL)

        def first_visible(selectors):
            for by, value in selectors:
                try:
                    for el in driver.find_elements(by, value):
                        if el.is_displayed() and el.is_enabled():
                            return el
                except Exception:
                    pass
            return None

        wait.until(lambda d: (
            first_visible([
                (By.ID, "username"), (By.NAME, "username"),
                (By.ID, "j_username"), (By.NAME, "j_username"),
                (By.CSS_SELECTOR, "input[type='text']"),
                (By.CSS_SELECTOR, "input[type='email']")
            ]) is not None
            or "selcompanylocationbs.action" in d.current_url.lower()
        ))

        username_el = first_visible([
            (By.ID, "username"), (By.NAME, "username"),
            (By.ID, "j_username"), (By.NAME, "j_username"),
            (By.CSS_SELECTOR, "input[type='text']"),
            (By.CSS_SELECTOR, "input[type='email']")
        ])
        password_el = first_visible([
            (By.ID, "password"), (By.NAME, "password"),
            (By.ID, "j_password"), (By.NAME, "j_password"),
            (By.CSS_SELECTOR, "input[type='password']")
        ])

        if username_el and password_el:
            username_el.clear()
            username_el.send_keys(VINCULUM_USERNAME)
            password_el.clear()
            password_el.send_keys(VINCULUM_PASSWORD)

            login_btn = first_visible([
                (By.ID, "loginButton"), (By.ID, "login"),
                (By.NAME, "login"),
                (By.CSS_SELECTOR, "button[type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']"),
                (By.XPATH, "//button[contains(translate(normalize-space(.),'LOGIN','login'),'login')]"),
                (By.XPATH, "//input[contains(translate(@value,'LOGIN','login'),'login')]")
            ])
            if not login_btn:
                raise RuntimeError("Login button nahi mila.")
            driver.execute_script("arguments[0].click();", login_btn)
        else:
            print("   Existing Vinculum session detected.")

        login_deadline = time.time() + 30
        while time.time() < login_deadline:
            try:
                alert = driver.switch_to.alert
                alert.accept()
                time.sleep(1)
                continue
            except Exception:
                pass

            dialogs = driver.find_elements(By.CSS_SELECTOR, "[role='dialog'], .modal, .ui-dialog, .modal-dialog")
            for dialog in dialogs:
                try:
                    if not dialog.is_displayed():
                        continue
                    buttons = dialog.find_elements(By.XPATH, ".//button | .//input[@type='button'] | .//input[@type='submit'] | .//a")
                    for btn in buttons:
                        if not btn.is_displayed() or not btn.is_enabled():
                            continue
                        label = " ".join(filter(None, [
                            btn.text, btn.get_attribute("value"),
                            btn.get_attribute("aria-label"), btn.get_attribute("title")
                        ])).strip().lower()
                        if any(k in label for k in ["continue", "ok", "yes", "proceed", "logout other", "terminate", "close"]):
                            driver.execute_script("arguments[0].click();", btn)
                            time.sleep(1)
                            break
                except Exception:
                    pass

            try:
                url_now = driver.current_url.lower()
                body = driver.find_element(By.TAG_NAME, "body").text.lower()
                if "selcompanylocationbs.action" in url_now or "pending report" in body or "order enquiry" in body:
                    print("   Login SUCCESS.")
                    break
            except Exception:
                pass

            time.sleep(1)

        current_report_id = None

        # ----------------------------------------------------
        # CREATE RETURNS EXPORT REQUEST
        # ----------------------------------------------------
        create_returns_export_request(driver, wait)

        # ----------------------------------------------------
        # PENDING REPORT
        # ----------------------------------------------------
        print("10) Pending Report iframe dhoondh raha hoon...")

        def find_and_switch_to_pending_iframe():
            driver.switch_to.default_content()
            time.sleep(0.5)
            all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
            for fr in all_iframes:
                if not fr.is_displayed():
                    continue
                try:
                    driver.switch_to.frame(fr)
                    body_text = driver.find_element(By.TAG_NAME, "body").text
                    if "Pending Report" in body_text or "Report ID" in body_text:
                        return True
                    driver.switch_to.default_content()
                except Exception:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
            return False

        found = find_and_switch_to_pending_iframe()
        if not found:
            time.sleep(3)
            found = find_and_switch_to_pending_iframe()
        if not found:
            raise RuntimeError("Pending Report iframe nahi mila.")

        time.sleep(2)

        STATUS_WORDS = {"SUCCESS", "ERROR", "WIP", "PENDING", "FAILED", "PROCESSING"}

        def parse_export_row(row):
            cells = row.find_elements(By.TAG_NAME, "td")
            texts = [c.text.strip() for c in cells]
            status_text = next((t.upper() for t in texts if t.upper() in STATUS_WORDS), None)
            report_id = next((t for t in texts if t.isdigit() and len(t) >= 4), None)
            error_msg = ""
            for t in texts:
                if not t or t.upper() in STATUS_WORDS or t.isdigit():
                    continue
                if REPORT_NAME_MARKER in t.upper().replace(" ", ""):
                    continue
                if "/" in t or ":" in t:
                    continue
                if len(t) > 8:
                    error_msg = t
                    break
            return texts, report_id, status_text, error_msg

        def read_latest_returns_export():
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow")
            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                joined = "".join(c.text.strip() for c in cells).upper().replace(" ", "")
                if REPORT_NAME_MARKER in joined:
                    texts, report_id, status_text, error_msg = parse_export_row(row)
                    return row, {"texts": texts, "report_id": report_id, "status": status_text, "error_msg": error_msg}
            print(f"   (debug) tr.jqgrow rows mile: {len(rows)}")
            return None, None

        status_ready = False
        attempt = 0
        export_retry_count = 0
        MAX_EXPORT_RETRIES = 10

        while not status_ready:
            attempt += 1
            try:
                driver.switch_to.default_content()
                find_and_switch_to_pending_iframe()
                row, info = read_latest_returns_export()

                if row is None:
                    print(f"   Attempt {attempt} - returns export row abhi nahi mila.")
                else:
                    report_id = info["report_id"]
                    status_text = info["status"] or ""
                    error_msg = info["error_msg"]
                    print(f"   Attempt {attempt} - Report {report_id} - Status: {status_text}")

                    if status_text == "SUCCESS":
                        current_report_id = report_id
                        status_ready = True
                        print(f"   SUCCESS mila - Report ID: {report_id}")
                        break

                    if status_text == "ERROR":
                        export_retry_count += 1
                        detail = error_msg or "Generic Business Error"
                        print(f"   Report {report_id} ERROR: {detail}. Fresh export retry {export_retry_count}/{MAX_EXPORT_RETRIES}...")
                        if export_retry_count > MAX_EXPORT_RETRIES:
                            raise RuntimeError(f"Fresh export retries exhausted. Last Report ID: {report_id}; Error: {detail}")
                        driver.switch_to.default_content()
                        time.sleep(1)
                        create_returns_export_request(driver, wait)
                        attempt = 0
                        time.sleep(2)
                        continue

            except RuntimeError:
                raise
            except Exception as exc:
                print(f"   Status read issue: {type(exc).__name__}: {exc}")

            try:
                find_and_switch_to_pending_iframe()
                driver.execute_script("if (typeof refreshGrid === 'function') { refreshGrid(); }")
                print("   Pending Report refresh kiya (10 sec interval).")
            except Exception as exc:
                print(f"   Refresh issue: {type(exc).__name__}: {exc}")

            time.sleep(10)

        # ----------------------------------------------------
        # DOWNLOAD
        # ----------------------------------------------------
        print("11) Download click kar raha hoon...")

        files_before_download = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))

        rows = driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow")
        export_rows = []
        for candidate in rows:
            cells = candidate.find_elements(By.TAG_NAME, "td")
            joined = "".join(c.text.strip() for c in cells).upper().replace(" ", "")
            if REPORT_NAME_MARKER not in joined:
                continue
            _, r_id, r_status, _ = parse_export_row(candidate)
            export_rows.append((candidate, r_id, r_status))

        row = None
        matched_report_id = None
        for candidate, r_id, r_status in export_rows:
            if r_id == str(current_report_id) and r_status == "SUCCESS":
                row = candidate
                matched_report_id = r_id
                break
        if row is None:
            for candidate, r_id, r_status in export_rows:
                if r_status == "SUCCESS":
                    row = candidate
                    matched_report_id = r_id
                    break

        if row is None:
            raise RuntimeError(f"SUCCESS returns export row download ke liye nahi mila. Report ID: {current_report_id}")

        print(f"   SUCCESS row confirmed - Report ID: {matched_report_id}")

        download_controls = row.find_elements(
            By.XPATH,
            ".//*[@onclick[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'download')]]"
        )
        if not download_controls:
            download_controls = row.find_elements(
                By.XPATH,
                ".//*[@onclick and contains(translate(@onclick, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'download')]"
            )
        if not download_controls:
            candidates = row.find_elements(By.CSS_SELECTOR, "label, a, button, input, img, i, span")
            for el in candidates:
                try:
                    meta = " ".join([
                        el.get_attribute("onclick") or "", el.get_attribute("title") or "",
                        el.get_attribute("aria-label") or "", el.get_attribute("alt") or "",
                        el.get_attribute("class") or "",
                    ]).lower()
                    if "download" in meta or "downloadreport" in meta:
                        download_controls.append(el)
                except Exception:
                    pass

        if not download_controls:
            raise RuntimeError("SUCCESS report mein Vinculum download icon/control nahi mila.")

        download_control = download_controls[0]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", download_control)
        time.sleep(0.5)

        try:
            WebDriverWait(driver, 10).until(lambda d: download_control.is_displayed() and download_control.is_enabled())
        except Exception:
            pass

        try:
            download_control.click()
        except Exception:
            driver.execute_script("arguments[0].click();", download_control)

        print(f"   Download initiated for Report ID: {matched_report_id}")

        print("12) Excel download hone ka wait...")
        downloaded_file = wait_for_download(DOWNLOAD_FOLDER, timeout=120, existing_files=files_before_download)

        if not downloaded_file:
            current_files = sorted(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")), key=os.path.getmtime, reverse=True)
            print("   Download folder files:", [os.path.basename(f) for f in current_files[:10]])
            raise RuntimeError("Excel download timeout ho gaya.")

        shutil.copy2(downloaded_file, OUTPUT_FILE)

        generated_at = datetime.utcnow().isoformat() + "Z"
        with open(META_FILE, "w") as f:
            json.dump({"generatedAt": generated_at}, f)

        print(f"SUCCESS: Latest Returns data saved as: {OUTPUT_FILE}")
        print(f"Meta file likhi gayi: {META_FILE} (generatedAt: {generated_at})")
        print(f"Downloaded source file: {os.path.basename(downloaded_file)}")

    except Exception as e:
        print("========================================")
        print("VINCULUM RETURNS EXTRACTION FAILED")
        print(f"Error: {e}")
        print("========================================")
        raise

    finally:
        driver.quit()


if __name__ == "__main__":
    extract_vinculum_returns()

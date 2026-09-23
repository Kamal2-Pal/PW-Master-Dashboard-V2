"""
Vinculum Returns (Inbound Enquiry) Extractor - GitHub Actions V5
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
    """Opens Inbound Enquiry by directly invoking the site's own openScreen()
    JS function - confirmed via inspect element to be exactly what the real
    'Inbound Enquiry' menu link's onclick does:

        onclick="openScreen('Inbound Enquiry', 'inboundEnquiryBS',
                             'fa fa-arrow-circle-right'); return false;"

    This replaces the previous approach of hovering/clicking through the WMS
    sidebar flyout menu, which repeatedly failed to reliably reveal the menu
    in headless Chrome. Calling the function directly is exactly what a real
    click does, without needing to find/hover the right sidebar icon at all.
    """
    print("2) Inbound Enquiry screen khol raha hoon (openScreen() JS call)...")
    driver.switch_to.default_content()

    try:
        driver.execute_script(
            "openScreen('Inbound Enquiry', 'inboundEnquiryBS', 'fa fa-arrow-circle-right');"
        )
        print("   openScreen() call ho gaya.")
    except Exception as exc:
        raise RuntimeError(f"openScreen() call fail hui: {exc}")

    time.sleep(2)

    # Find which document context (top-level page, or one of possibly several
    # iframes) actually contains the Inbound Enquiry screen's filter
    # controls. Polls repeatedly (checking top-level, then every iframe,
    # including one level of nesting) until the markers show up or timeout.
    def screen_markers_present():
        try:
            return bool(driver.find_elements(By.ID, "gs_displayInboundType")) or \
                   bool(driver.find_elements(By.ID, "gs_createdDate"))
        except Exception:
            return False

    found_context = False
    deadline = time.time() + 25
    attempt = 0
    while time.time() < deadline and not found_context:
        attempt += 1
        driver.switch_to.default_content()
        if screen_markers_present():
            found_context = True
            print(f"   (attempt {attempt}) Inbound Enquiry screen top-level page par mil gaya.")
            break

        all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for fr in all_iframes:
            try:
                if not fr.is_displayed():
                    continue
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)
                if screen_markers_present():
                    found_context = True
                    print(f"   (attempt {attempt}) Inbound Enquiry screen iframe ke andar mil gaya ({len(all_iframes)} iframe(s) the).")
                    break

                nested_iframes = driver.find_elements(By.TAG_NAME, "iframe")
                for nfr in nested_iframes:
                    try:
                        if not nfr.is_displayed():
                            continue
                        driver.switch_to.frame(nfr)
                        if screen_markers_present():
                            found_context = True
                            print(f"   (attempt {attempt}) Inbound Enquiry screen NESTED iframe ke andar mil gaya.")
                            break
                        driver.switch_to.parent_frame()
                    except Exception:
                        try:
                            driver.switch_to.parent_frame()
                        except Exception:
                            pass
                if found_context:
                    break
            except Exception:
                continue

        if not found_context:
            time.sleep(1)

    if not found_context:
        driver.switch_to.default_content()
        print(f"   {attempt} attempts ke baad bhi screen markers nahi mile; top-level page par hi aage badh raha hoon.")
        try:
            print(f"   [debug] Current URL: {driver.current_url}")
            print(f"   [debug] Open windows/tabs: {len(driver.window_handles)}")
            print(f"   [debug] iframe count (top-level): {len(driver.find_elements(By.TAG_NAME, 'iframe'))}")
            body_text = driver.find_element(By.TAG_NAME, "body").text
            print(f"   [debug] Page body snippet (first 400 chars): {body_text[:400]!r}")
        except Exception as exc:
            print(f"   [debug] Diagnostics collection failed: {exc}")


def set_inbound_type_filter(driver, wait, value):
    """Sets the Inbound Type filter dropdown to the given value (e.g. 'Against ASN')."""
    print(f"4) Inbound Type filter ko '{value}' set kar raha hoon...")

    # Primary: exact ID confirmed via inspect element (id="gs_displayInboundType").
    try:
        sel = wait.until(EC.presence_of_element_located((By.ID, "gs_displayInboundType")))
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
        selected_text = driver.execute_script(
            "return arguments[0].options[arguments[0].selectedIndex].textContent.trim();", sel
        )
        if selected_text.strip().lower() == value.lower():
            print(f"   #gs_displayInboundType ke through set ho gaya (selected: '{selected_text}').")
            return
        print(f"   #gs_displayInboundType mila lekin selection confirm nahi hui (got '{selected_text}') - fallback try kar raha hoon...")
    except Exception as exc:
        print(f"   #gs_displayInboundType se set nahi hua ({exc}) - fallback try kar raha hoon...")

    # Fallback 1: any native <select> on the page containing this option text.
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
                print("   Native select ke through set ho gaya (fallback).")
                return
        except Exception:
            pass

    # Fallback 2: custom dropdown - click the "--- Select ---" toggle for the
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
            print("   Custom dropdown ke through set ho gaya (fallback).")
            return
        except Exception:
            continue

    raise RuntimeError(f"Inbound Type filter ('{value}') set nahi kar paaya.")


def set_creation_date_preset(driver, wait, preset_label):
    """Opens the Creation Date range picker and clicks a preset like 'This Month'."""
    print(f"5) Creation Date filter ko '{preset_label}' set kar raha hoon...")

    date_field = None
    # Primary: exact ID confirmed via inspect element (id="gs_createdDate").
    try:
        el = wait.until(EC.presence_of_element_located((By.ID, "gs_createdDate")))
        if el.is_displayed():
            date_field = el
            print("   #gs_createdDate field mil gaya.")
    except Exception:
        pass

    if date_field is None:
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

    print("6) Search button click kar raha hoon (#SearchBtn)...")
    search_btn = wait.until(EC.element_to_be_clickable((By.ID, "SearchBtn")))
    search_btn.click()
    print("   Data load hone ka wait (10 sec)...")
    time.sleep(10)

    print("7) Detail Export click kar raha hoon (detailExportClickNew())...")

    # IMPORTANT:
    # Inspect Element shows id="downloadButton", but this id can be repeated
    # in the grid. Do NOT blindly use By.ID("downloadButton"), because
    # Selenium may click a different row/button. Identify the actual
    # Detail Export button by its onclick handler + visible label.
    detail_export_btn = None
    detail_deadline = time.time() + 30

    while time.time() < detail_deadline and detail_export_btn is None:
        candidates = driver.find_elements(
            By.XPATH,
            "//button[contains(@onclick,'detailExportClickNew') and "
            ".//label[contains(normalize-space(.),'Detail Export')]]"
        )

        for el in candidates:
            try:
                if el.is_displayed() and el.is_enabled():
                    detail_export_btn = el
                    break
            except Exception:
                continue

        if detail_export_btn is None:
            # Fallback to the confirmed id, but still require the button to
            # visibly contain the Detail Export label.
            for el in driver.find_elements(By.ID, "downloadButton"):
                try:
                    label_text = (el.text or "").strip().lower()
                    onclick = (el.get_attribute("onclick") or "").lower()
                    if (
                        el.is_displayed()
                        and el.is_enabled()
                        and "detailexportclicknew" in onclick
                        and "detail export" in label_text
                    ):
                        detail_export_btn = el
                        break
                except Exception:
                    continue

        if detail_export_btn is None:
            time.sleep(1)

    if detail_export_btn is None:
        raise RuntimeError(
            "Detail Export button nahi mila. "
            "Confirmed onclick='detailExportClickNew()' + 'Detail Export' label "
            "wala button locate nahi ho paaya."
        )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        detail_export_btn
    )
    time.sleep(0.5)
    try:
        driver.execute_script("arguments[0].click();", detail_export_btn)
    except Exception:
        detail_export_btn.click()

    print("   Detail Export button click ho gaya.")
    time.sleep(2)

    print("8) Export fields select kar raha hoon...")

    # IMPORTANT: Inspect Element confirmed the real export dialog controls:
    #   Select-all checkbox -> input#cb_dynamicFieldGrid
    #   Export button       -> button[title="Export"]
    # The dialog itself is not reliably represented as a Bootstrap
    # .modal-content in GitHub Actions, so do NOT wait for a modal wrapper.
    # Instead, search directly for the confirmed control across every
    # window/tab and nested iframe.

    def find_element_recursive(locator, depth=0, max_depth=6):
        """Find a visible element in the current window/frame tree."""
        try:
            elements = driver.find_elements(*locator)
            for el in elements:
                try:
                    if el.is_displayed():
                        return el
                except Exception:
                    pass
        except Exception:
            pass

        if depth >= max_depth:
            return None

        try:
            frames = driver.find_elements(By.TAG_NAME, "iframe")
        except Exception:
            return None

        for fr in frames:
            try:
                # Do not require iframe.is_displayed(). In headless Chrome,
                # an iframe containing a dynamically opened export dialog can
                # report as not displayed even though its DOM is accessible.
                driver.switch_to.frame(fr)
                found = find_element_recursive(locator, depth + 1, max_depth)
                if found is not None:
                    return found
                driver.switch_to.parent_frame()
            except Exception:
                try:
                    driver.switch_to.parent_frame()
                except Exception:
                    driver.switch_to.default_content()
        return None

    select_all_cb = None
    export_window = None
    deadline = time.time() + 90
    attempt = 0

    while time.time() < deadline and select_all_cb is None:
        attempt += 1

        for handle in list(driver.window_handles):
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()

                # Exact ID confirmed from Inspect Element.
                select_all_cb = find_element_recursive(
                    (By.ID, "cb_dynamicFieldGrid")
                )

                if select_all_cb is not None:
                    export_window = handle
                    print(
                        f"   (attempt {attempt}) Select-all checkbox "
                        "#cb_dynamicFieldGrid mil gaya."
                    )
                    break
            except Exception:
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass

        if select_all_cb is None:
            time.sleep(1)

    if select_all_cb is None:
        driver.switch_to.default_content()
        raise RuntimeError(
            "Detail Export ke baad #cb_dynamicFieldGrid checkbox 90 sec "
            "tak nahi mila. Inspect Element ke confirmed selector ke "
            "according dialog load nahi hua."
        )

    if export_window:
        driver.switch_to.window(export_window)

    # IMPORTANT: The export dialog can re-render its field grid immediately
    # after it appears. The Selenium WebElement returned by the recursive
    # search can therefore become stale between "find" and "click".
    # Do NOT call select_all_cb.is_selected() on that old WebElement.
    # Re-query the live DOM inside the CURRENT iframe and click the current
    # checkbox in the same JavaScript execution.
    try:
        clicked = driver.execute_script("""
            const cb = document.querySelector("#cb_dynamicFieldGrid");
            if (!cb) return false;
            if (!cb.checked) {
                cb.click();
            }
            return !!cb.checked;
        """)
        if not clicked:
            # One short retry in case the grid was re-rendering at this exact
            # moment. This still uses a fresh DOM lookup, never the stale
            # Selenium element captured above.
            time.sleep(1)
            clicked = driver.execute_script("""
                const cb = document.querySelector("#cb_dynamicFieldGrid");
                if (!cb) return false;
                if (!cb.checked) {
                    cb.click();
                }
                return !!cb.checked;
            """)
        if not clicked:
            raise RuntimeError(
                "Live DOM mein #cb_dynamicFieldGrid mila nahi ya checked nahi hua."
            )
        print("   Select-all checkbox click ho gaya.")
    except Exception as exc:
        raise RuntimeError(
            "#cb_dynamicFieldGrid mila, lekin fresh DOM lookup se select-all click nahi ho paya."
        ) from exc

    print("9) Export click kar raha hoon...")

    # Exact Export button selector confirmed from Inspect Element:
    # <button ... title="Export" onclick="exportData();">
    export_btn = find_element_recursive(
        (By.CSS_SELECTOR, "button[title='Export']")
    )

    if export_btn is None:
        # Text fallback only if the exact inspected selector is unavailable.
        export_btn = find_element_recursive(
            (
                By.XPATH,
                "//button[contains(translate(normalize-space(.),"
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
                "'export')]"
            )
        )

    if export_btn is None:
        raise RuntimeError(
            "Export modal/field grid mila, lekin exact "
            "button[title='Export'] nahi mila."
        )

    driver.execute_script("arguments[0].click();", export_btn)
    print("   Export button click ho gaya (exportData()).")
    time.sleep(3)

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
        # PENDING REPORT - WAIT FOR SUCCESS, THEN DOWNLOAD
        # ----------------------------------------------------
        # IMPORTANT: create the export request ONLY ONCE.
        # We first capture the existing report IDs, then create one fresh
        # request. This lets us reliably identify the exact new report.
        print("10) Pending Report open/monitor kar raha hoon...")

        def switch_to_pending_report(open_screen=False):
            """Locate the real Pending Report grid using the same proven
            iframe logic as vinculum_extract_final.py.

            The reference extractor successfully reaches Pending Report by
            searching the page's iframes for the actual report-grid content.
            We keep that proven behavior here, while also allowing a
            top-level fallback for the DOM layout seen in manual Inspect.
            """
            driver.switch_to.default_content()

            if open_screen:
                try:
                    driver.execute_script(
                        'openScreen("Pending Report", '
                        '"pendingDisplay","fa fa-fw fa-truck");'
                    )
                    print("   Pending Report openScreen() call ho gaya.")
                    time.sleep(2)
                except Exception as exc:
                    print(f"   Pending Report openScreen() warning: {exc}")

            def content_present():
                try:
                    body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
                    rows = driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow")
                    return (
                        ("pending report" in body_text or "report id" in body_text
                         or "report names" in body_text)
                        and bool(rows)
                    )
                except Exception:
                    return False

            deadline = time.time() + 20

            while time.time() < deadline:
                # First check the current/top-level document.
                driver.switch_to.default_content()
                if content_present():
                    print("   Pending Report top-level document mein mil gaya.")
                    return True

                # Proven reference flow: inspect every visible iframe.
                try:
                    frames = driver.find_elements(By.TAG_NAME, "iframe")
                except Exception:
                    frames = []

                for fr in frames:
                    try:
                        # Do not depend on iframe visibility alone; headless
                        # Chrome can expose the document before is_displayed()
                        # becomes true.
                        driver.switch_to.default_content()
                        driver.switch_to.frame(fr)

                        if content_present():
                            print("   Pending Report iframe mein mil gaya.")
                            return True

                        nested = driver.find_elements(By.TAG_NAME, "iframe")
                        for nfr in nested:
                            try:
                                driver.switch_to.frame(nfr)
                                if content_present():
                                    print("   Pending Report nested iframe mein mil gaya.")
                                    return True
                                driver.switch_to.parent_frame()
                            except Exception:
                                try:
                                    driver.switch_to.parent_frame()
                                except Exception:
                                    pass
                    except Exception:
                        try:
                            driver.switch_to.default_content()
                        except Exception:
                            pass

                driver.switch_to.default_content()
                time.sleep(1)

            driver.switch_to.default_content()
            try:
                body = driver.find_element(By.TAG_NAME, "body").text
                print(f"   [debug] Pending Report search timeout. Body snippet: {body[:1000]!r}")
                print(f"   [debug] Top-level iframe count: {len(driver.find_elements(By.TAG_NAME, 'iframe'))}")
            except Exception as exc:
                print(f"   [debug] Pending Report diagnostics failed: {exc}")
            return False

        def refresh_pending_report():
            if not switch_to_pending_report():
                return False
            try:
                refreshed = driver.execute_script(
                    "if (typeof refreshGrid === 'function') { refreshGrid(); return true; } return false;"
                )
                if refreshed:
                    print("   Pending Report ka actual Refresh (refreshGrid()) click hua.")
                    return True

                # If the screen function is scoped differently, click the
                # inspected refresh control directly as a fallback.
                refresh_links = driver.find_elements(
                    By.CSS_SELECTOR, "#refreshBtn a[onclick*='refreshGrid']"
                )
                if refresh_links:
                    driver.execute_script("arguments[0].click();", refresh_links[0])
                    print("   Pending Report refresh link click hua.")
                    return True
                return False
            finally:
                driver.switch_to.default_content()

        def get_return_rows():
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow")
            result = []
            for row in rows:
                try:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    texts = [c.text.strip() for c in cells]
                    joined = " ".join(texts).upper().replace(" ", "")
                    # Report grid may show either the generated function name
                    # or the human-readable report name.
                    if (
                        "INBOUNDENQ" not in joined
                        and "GENERATEINBOUND" not in joined
                        and "INBOUNDENQUIRY" not in joined
                    ):
                        continue
                    status_text = next(
                        (t.upper() for t in texts if t.upper() in {"SUCCESS", "ERROR", "WIP", "PENDING", "FAILED", "PROCESSING"}),
                        ""
                    )
                    report_id = next((t for t in texts if t.isdigit() and len(t) >= 4), None)
                    result.append({"row": row, "texts": texts, "report_id": report_id, "status": status_text})
                except Exception:
                    continue
            return result

        if not switch_to_pending_report(open_screen=True):
            time.sleep(3)
            if not switch_to_pending_report():
                raise RuntimeError("Pending Report screen/grid nahi mila.")
        driver.switch_to.default_content()

        switch_to_pending_report()
        before_ids = {r["report_id"] for r in get_return_rows() if r["report_id"]}
        print(f"   Existing return report IDs captured: {sorted(before_ids)[-10:]}")
        driver.switch_to.default_content()

        create_returns_export_request(driver, wait)
        print("   Fresh return export request create ho gaya (single request).")
        print("   Jab tak status SUCCESS nahi hota, download button nahi dabega.")

        current_report_id = None
        status_ready = False
        max_wait_seconds = 600
        poll_seconds = 10
        started = time.time()

        while time.time() - started < max_wait_seconds:
            try:
                refresh_pending_report()
            except Exception as exc:
                print(f"   Refresh issue: {type(exc).__name__}: {exc}")

            time.sleep(1)
            if not switch_to_pending_report():
                print("   Pending Report frame abhi available nahi hai; next poll...")
                time.sleep(poll_seconds)
                continue

            rows = get_return_rows()
            driver.switch_to.default_content()
            new_rows = [r for r in rows if r["report_id"] and r["report_id"] not in before_ids]
            candidate = new_rows[0] if new_rows else (rows[0] if rows else None)

            if candidate is None:
                print("   Return export row abhi nahi mila; 10 sec baad refresh karunga.")
                time.sleep(poll_seconds)
                continue

            current_report_id = candidate["report_id"]
            status_text = candidate["status"]
            print(f"   Report {current_report_id} - Status: {status_text or 'UNKNOWN'}")

            if status_text in {"ERROR", "FAILED"}:
                raise RuntimeError(
                    f"Return export Report ID {current_report_id} ERROR/FAILED hua. Row: {' | '.join(candidate['texts'])}"
                )

            if status_text == "SUCCESS":
                status_ready = True
                print(f"   SUCCESS mila - Report ID: {current_report_id}")
                break

            print("   Status SUCCESS nahi hua, isliye download button abhi nahi dabega.")
            time.sleep(poll_seconds)

        if not status_ready:
            raise RuntimeError(
                f"Return export {current_report_id or 'new report'} {max_wait_seconds} seconds mein SUCCESS nahi hua."
            )

        # ----------------------------------------------------
        # DOWNLOAD - ONLY AFTER SUCCESS
        # ----------------------------------------------------
        print("11) SUCCESS confirm ho gaya. Ab isi report ka download click kar raha hoon...")
        files_before_download = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))

        if not switch_to_pending_report():
            raise RuntimeError("SUCCESS ke baad Pending Report screen/grid nahi mila.")

        rows = get_return_rows()
        target = next(
            (r for r in rows if r["report_id"] == str(current_report_id) and r["status"] == "SUCCESS"),
            None,
        )
        if target is None:
            driver.switch_to.default_content()
            raise RuntimeError(f"SUCCESS Report ID {current_report_id} ka exact row download ke liye nahi mila.")

        row = target["row"]
        print(f"   Exact SUCCESS row confirmed - Report ID: {current_report_id}")

        # Inspect Element confirmed: label onclick="javascript:downloadReport("REPORT_ID", ...); return false;"
        download_controls = row.find_elements(By.CSS_SELECTOR, "label[onclick*='downloadReport']")
        if not download_controls:
            download_controls = row.find_elements(
                By.XPATH,
                ".//label[contains(translate(@onclick, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'downloadreport')]"
            )
        if not download_controls:
            driver.switch_to.default_content()
            raise RuntimeError(f"SUCCESS Report ID {current_report_id} mein exact downloadReport control nahi mila.")

        download_control = download_controls[0]
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", download_control)
        time.sleep(0.5)
        try:
            driver.execute_script("arguments[0].click();", download_control)
        except Exception:
            download_control.click()

        print(f"   Download button click ho gaya - Report ID: {current_report_id}")
        driver.switch_to.default_content()
        print("12) Excel download hone ka wait...")
        downloaded_file = wait_for_download(DOWNLOAD_FOLDER, timeout=120, existing_files=files_before_download)
        if not downloaded_file:
            current_files = sorted(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")), key=os.path.getmtime, reverse=True)
            print("   Download folder files:", [os.path.basename(f) for f in current_files[:10]])
            raise RuntimeError("Excel download timeout ho gaya.")
        print(f"   Download complete: {os.path.basename(downloaded_file)}")

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

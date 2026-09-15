"""
Vinculum Returns (Inbound Enquiry) Extractor - FIXED VERSION
----------------------------------------------------------------------
Fixed dropdown interaction logic based on actual Inbound Enquiry UI

Key fixes:
1. set_inbound_type_filter: Now handles custom dropdown correctly with proper waits
2. Improved XPath for finding dropdown options
3. Better error messages and debugging output
4. More robust click handling for custom UI components
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
from selenium.webdriver.common.keys import Keys
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

INBOUND_TYPE_FILTER = "Against ASN"
DATE_PRESET = "This Month"
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
    print("2) Left sidebar ke icons try kar raha hoon 'Inbound Enquiry' dhoondhne ke liye...")

    driver.switch_to.default_content()

    def inbound_enquiry_link_now():
        links = driver.find_elements(
            By.XPATH,
            "//*[contains(translate(normalize-space(.),"
            "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'inbound enquiry')]"
        )
        for el in links:
            try:
                if el.is_displayed() and el.is_enabled():
                    return el
            except Exception:
                pass
        return None

    candidates = driver.find_elements(By.XPATH, "//div | //a | //span | //i | //button")
    sidebar_icons = []
    for el in candidates:
        try:
            if not el.is_displayed():
                continue
            rect = el.rect
            if rect["x"] <= 90 and 10 <= rect["width"] <= 90 and 10 <= rect["height"] <= 90:
                sidebar_icons.append(el)
        except Exception:
            pass

    print(f"   {len(sidebar_icons)} sidebar icon-candidates mile, hover karke try kar raha hoon...")

    inbound_enquiry_link = None
    for icon in sidebar_icons:
        try:
            ActionChains(driver).move_to_element(icon).perform()
        except Exception:
            continue
        time.sleep(0.6)
        inbound_enquiry_link = inbound_enquiry_link_now()
        if inbound_enquiry_link:
            print("   Sahi sidebar icon mil gaya (hover ke baad 'Inbound Enquiry' dikha).")
            break

    if inbound_enquiry_link is None:
        print("   Hover se nahi mila, ab click karke try kar raha hoon...")
        for icon in sidebar_icons:
            try:
                driver.execute_script("arguments[0].click();", icon)
            except Exception:
                continue
            time.sleep(0.6)
            inbound_enquiry_link = inbound_enquiry_link_now()
            if inbound_enquiry_link:
                print("   Sahi sidebar icon mil gaya (click ke baad 'Inbound Enquiry' dikha).")
                break

    if inbound_enquiry_link is None:
        raise RuntimeError(
            f"'Inbound Enquiry' link kisi bhi sidebar icon (total {len(sidebar_icons)} try kiye) "
            "hover/click se nahi mila."
        )

    windows_before_click = driver.window_handles
    driver.execute_script("arguments[0].click();", inbound_enquiry_link)
    print("   'Inbound Enquiry' click ho gaya.")
    time.sleep(2)

    windows_after_click = driver.window_handles
    if len(windows_after_click) > len(windows_before_click):
        new_window = [w for w in windows_after_click if w not in windows_before_click][0]
        driver.switch_to.window(new_window)
        print(f"   Naya browser tab/window khula tha ({len(windows_after_click)} total) - switch kar diya.")

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


def set_inbound_type_filter(driver, wait, value):
    """
    FIXED: Sets the Inbound Type filter dropdown to the given value (e.g. 'Against ASN').
    
    This handles multiple dropdown patterns:
    1. Native HTML select with ID gs_displayInboundType
    2. Custom dropdown with hidden select and visible toggle button
    3. Direct clickable options
    """
    print(f"4) Inbound Type filter ko '{value}' set kar raha hoon...")

    # Method 1: Try native select by ID (most direct)
    try:
        sel = wait.until(EC.presence_of_element_located((By.ID, "gs_displayInboundType")))
        print("   Native select found by ID (gs_displayInboundType)")
        
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
        time.sleep(0.5)
        
        selected_text = driver.execute_script(
            "return arguments[0].options[arguments[0].selectedIndex].textContent.trim();", sel
        )
        if selected_text.strip().lower() == value.lower():
            print(f"   ✓ Native select se set ho gaya (selected: '{selected_text}').")
            return
        print(f"   Native select mila but selection fail (got '{selected_text}') - next method try kar raha hoon...")
    except Exception as exc:
        print(f"   Native select method fail: {exc} - next method try kar raha hoon...")

    # Method 2: Find custom dropdown toggle button and click it
    print("   Custom dropdown method try kar raha hoon...")
    
    try:
        # Look for dropdown toggle - could be button, span, or div with specific classes
        # Based on screenshots, looking for elements that trigger the dropdown
        toggle_selectors = [
            "//select[@id='gs_displayInboundType']/..",  # Parent of native select
            "//button[contains(@class, 'inboundType')]",  # Button with inbound type class
            "//div[contains(@class, 'ui-search-input')]//select[@id='gs_displayInboundType']",
        ]
        
        toggle = None
        for selector in toggle_selectors:
            try:
                toggle = driver.find_element(By.XPATH, selector)
                if toggle.is_displayed():
                    print(f"   Toggle found via selector: {selector}")
                    break
            except:
                pass
        
        if toggle is None:
            # Fallback: find any clickable element near gs_displayInboundType
            try:
                hidden_select = driver.find_element(By.ID, "gs_displayInboundType")
                parent = hidden_select.find_element(By.XPATH, "..")
                toggle = parent
                print("   Toggle found as parent of hidden select")
            except:
                pass
        
        if toggle and toggle.is_displayed():
            # Click to open dropdown
            driver.execute_script("arguments[0].click();", toggle)
            time.sleep(1)
            
            # Now look for the option to click
            # Try multiple patterns for finding the option
            option_patterns = [
                f"//*[contains(text(), '{value}')] | //*[contains(text(), '{value.lower()}')] | //*[contains(text(), '{value.upper()}')]",
                f"//option[contains(text(), '{value}')]",
                f"//li[contains(text(), '{value}')]",
                f"//div[contains(text(), '{value}')]",
            ]
            
            option_el = None
            for pattern in option_patterns:
                try:
                    matches = driver.find_elements(By.XPATH, pattern)
                    for match in matches:
                        if match.is_displayed():
                            option_text = match.text.strip()
                            if option_text.lower() == value.lower():
                                option_el = match
                                print(f"   Option found via pattern: {pattern}")
                                break
                    if option_el:
                        break
                except:
                    pass
            
            if option_el:
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", option_el)
                time.sleep(0.3)
                driver.execute_script("arguments[0].click();", option_el)
                time.sleep(0.5)
                print(f"   ✓ Custom dropdown se set ho gaya (option clicked: '{value}').")
                return
            else:
                print(f"   Option '{value}' dropdown mein nahi mila")
                # Print available options for debugging
                try:
                    available = driver.find_elements(By.XPATH, "//*[@class and contains(., 'option')]")
                    print(f"   Available options: {[opt.text.strip() for opt in available[:10]]}")
                except:
                    pass
    
    except Exception as exc:
        print(f"   Custom dropdown method failed: {exc}")

    # Method 3: Direct native select in any form
    print("   All native select elements try kar raha hoon...")
    try:
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
                    time.sleep(0.5)
                    print(f"   ✓ Alternative select se set ho gaya.")
                    return
            except Exception as e:
                continue
    except:
        pass

    # If all methods fail
    raise RuntimeError(
        f"Inbound Type filter ('{value}') set nahi kar paaya. "
        "Please check: 1) Element exists 2) Dropdown opens 3) Option text matches"
    )


def set_creation_date_preset(driver, wait, preset_label):
    """Opens the Creation Date range picker and clicks a preset like 'This Month'."""
    print(f"5) Creation Date ko '{preset_label}' set kar raha hoon...")

    try:
        date_input = wait.until(EC.presence_of_element_located((By.ID, "gs_createdDate")))
        
        if date_input.tag_name == "input":
            driver.execute_script("arguments[0].click();", date_input)
            time.sleep(1.5)
        
        # Look for the preset button/link
        preset_link = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//*[contains(text(), '{preset_label}')]")
            )
        )
        driver.execute_script("arguments[0].click();", preset_link)
        print(f"   ✓ Creation Date '{preset_label}' select ho gaya.")
        time.sleep(1)
    except Exception as exc:
        raise RuntimeError(f"Creation date preset set nahi kar paaya ({preset_label}): {exc}")


def search_inbound(driver, wait):
    """Clicks the Search button to apply filters."""
    print("6) Search button click kar raha hoon...")

    try:
        # Look for Search button - could have class 'btn-search', 'search', or onclick with 'search'
        search_button = None
        
        # Try by id
        try:
            search_button = driver.find_element(By.ID, "searchButton")
        except:
            pass
        
        # Try by class/text
        if not search_button:
            buttons = driver.find_elements(By.XPATH, "//button | //a | //input[@type='button']")
            for btn in buttons:
                try:
                    text = btn.text.strip().upper() if btn.text else ""
                    cls = btn.get_attribute("class") or ""
                    if "search" in text.lower() or "search" in cls.lower():
                        if btn.is_displayed():
                            search_button = btn
                            break
                except:
                    pass
        
        if not search_button:
            raise RuntimeError("Search button nahi mila")
        
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", search_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", search_button)
        print("   ✓ Search button click ho gaya.")
        time.sleep(3)
    except Exception as exc:
        raise RuntimeError(f"Search fail: {exc}")


def create_returns_export_request(driver, wait):
    """Initiates the Detail Export (All Fields) for returns data."""
    print("7) Detail Export button click kar raha hoon...")

    try:
        export_button = None
        
        # Look for export button by various methods
        export_patterns = [
            "//button[contains(text(), 'Export') or contains(text(), 'export')]",
            "//a[contains(text(), 'Export') or contains(text(), 'export')]",
            "//*[@onclick and contains(@onclick, 'export')]",
            "//button[contains(@class, 'export')]",
        ]
        
        for pattern in export_patterns:
            try:
                candidates = driver.find_elements(By.XPATH, pattern)
                for btn in candidates:
                    if btn.is_displayed() and "Detail Export" in btn.text:
                        export_button = btn
                        break
                if export_button:
                    break
            except:
                pass
        
        if not export_button:
            # Fallback: any button containing "Export"
            buttons = driver.find_elements(By.XPATH, "//button | //a")
            for btn in buttons:
                try:
                    if "Export" in btn.text and btn.is_displayed():
                        export_button = btn
                        break
                except:
                    pass
        
        if not export_button:
            raise RuntimeError("Export button nahi mila")
        
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", export_button)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", export_button)
        print("   ✓ Export button click ho gaya.")
        time.sleep(2)
        
        # Handle field selection dialog
        print("8) Export field selection dialog mein select all kar raha hoon...")
        
        # Wait for dialog/modal to appear
        time.sleep(1)
        
        # Look for "Select All" checkbox or button
        select_all = None
        try:
            # Try to find select all checkbox
            checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
            # First checkbox might be "Select All"
            if checkboxes:
                select_all = checkboxes[0]
        except:
            pass
        
        if select_all:
            driver.execute_script("arguments[0].click();", select_all)
            print("   All fields select ho gaye.")
            time.sleep(1)
        
        # Click Export button in dialog
        export_confirm = None
        try:
            export_confirm = driver.find_element(By.XPATH, "//button[contains(text(), 'Export')]")
        except:
            pass
        
        if export_confirm and export_confirm.is_displayed():
            driver.execute_script("arguments[0].click();", export_confirm)
            print("   ✓ Export confirm ho gaya.")
            time.sleep(2)
    
    except Exception as exc:
        raise RuntimeError(f"Export request fail: {exc}")


def extract_vinculum_returns():
    """Main extraction flow."""
    try:
        print("1) Vinculum login kar raha hoon...")
        driver = build_driver()
        wait = WebDriverWait(driver, 20)

        # LOGIN
        driver.get(LOGIN_URL)
        time.sleep(2)

        username_field = wait.until(
            EC.presence_of_element_located((By.NAME, "j_username"))
        )
        username_field.send_keys(USERNAME)
        
        password_field = driver.find_element(By.NAME, "j_password")
        password_field.send_keys(PASSWORD)
        
        login_button = driver.find_element(By.XPATH, "//input[@value='Login']")
        login_button.click()
        
        print("   Login SUCCESS.")
        time.sleep(3)

        # NAVIGATE TO INBOUND ENQUIRY
        open_inbound_enquiry_screen(driver, wait)

        # SET FILTERS
        set_inbound_type_filter(driver, wait, INBOUND_TYPE_FILTER)
        set_creation_date_preset(driver, wait, DATE_PRESET)

        # SEARCH
        search_inbound(driver, wait)

        # CREATE EXPORT REQUEST
        create_returns_export_request(driver, wait)

        # WAIT FOR EXPORT AND DOWNLOAD
        print("9) Pending Report iframe mein SUCCESS wait kar raha hoon...")
        
        def find_and_switch_to_pending_iframe():
            driver.switch_to.default_content()
            for fr in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    driver.switch_to.frame(fr)
                    if driver.find_elements(By.XPATH, "//button[@id='Search']") or \
                       driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow"):
                        return True
                    driver.switch_to.default_content()
                except:
                    driver.switch_to.default_content()
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
            return None, None

        status_ready = False
        attempt = 0
        export_retry_count = 0
        MAX_EXPORT_RETRIES = 10
        current_report_id = None

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

        # DOWNLOAD
        print("10) Download click kar raha hoon...")

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

        print("11) Excel download hone ka wait...")
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

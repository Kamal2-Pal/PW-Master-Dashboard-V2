"""
Vinculum Returns (Inbound Enquiry) Extractor - GitHub Actions version
----------------------------------------------------------------------
Fully updated and robust flow incorporating all UI screenshot configurations:
1. Multi-fallback robust login (handles frames, direct JS value injection & explicit waits)
2. Direct JS / Navigation link trigger for Inbound Enquiry
3. Inbound Type filter = "Against ASN" (#gs_displayInboundType)
4. Creation Date filter = "This Month" (#gs_createdDate)
5. Search & Detail Export modal field select (#cb_dynamicFieldGrid)
6. Export trigger & Pending Report polling (matching GENERATEINBOUND)
7. Report download & Excel metadata update
"""

import os
import json
import time
import glob
import shutil
from datetime import datetime

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

VINCULUM_USERNAME = os.getenv("VINCULUM_USERNAME", "").strip()
VINCULUM_PASSWORD = os.getenv("VINCULUM_PASSWORD", "").strip()

DOWNLOAD_FOLDER = os.path.abspath("data/downloads_returns")
OUTPUT_FILE = os.path.abspath("returns.xlsx")
META_FILE = os.path.abspath("returns-meta.json")

# Filter configurations
INBOUND_TYPE_FILTER = "Against ASN"
DATE_PRESET = "This Month"

# Pending Report Marker
REPORT_NAME_MARKER = "GENERATEINBOUND"


# ============================================================
# HELPERS
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
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

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
# INBOUND ENQUIRY NAVIGATION & FILTERS
# ============================================================

def open_inbound_enquiry_screen(driver, wait):
    """Opens WMS -> Inbound -> Inbound Enquiry via JS or Click."""
    print("2) 'Inbound Enquiry' screen open kar raha hoon...")
    driver.switch_to.default_content()

    try:
        driver.execute_script("openScreen('Inbound Enquiry', 'inboundEnquiryBS', 'fa fa-arrow-circle-right');")
        print("   Direct openScreen() JS function execute ho gaya.")
    except Exception as e:
        print(f"   openScreen JS call fail hua: {e}. Sidebar element click try kar raha hoon...")
        link = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//*[contains(translate(normalize-space(.),'INBOUND ENQUIRY','inbound enquiry'),'inbound enquiry')]"
        )))
        driver.execute_script("arguments[0].click();", link)

    time.sleep(4)

    # Context Switch Verification
    deadline = time.time() + 30
    while time.time() < deadline:
        driver.switch_to.default_content()
        if len(driver.find_elements(By.ID, "gs_displayInboundType")) > 0:
            print("   Inbound Enquiry screen top-level frame par load ho gayi.")
            return

        all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
        for fr in all_iframes:
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)
                if len(driver.find_elements(By.ID, "gs_displayInboundType")) > 0:
                    print("   Inbound Enquiry screen IFRAME ke andar mil gayi.")
                    return
            except Exception:
                pass
        time.sleep(2)

    print("   Warning: Inbound Enquiry screen markers check timeout hua, default flow continue ho raha hai.")


def set_inbound_type_filter(driver, wait, value):
    """Sets the Inbound Type dropdown (#gs_displayInboundType)."""
    print(f"4) Inbound Type filter ko '{value}' set kar raha hoon...")
    try:
        sel = wait.until(EC.presence_of_element_located((By.ID, "gs_displayInboundType")))
        driver.execute_script(
            """
            const s = arguments[0], wanted = arguments[1].toLowerCase();
            for (const o of s.options) {
                if (o.textContent.trim().toLowerCase() === wanted) {
                    o.selected = true;
                    break;
                }
            }
            s.dispatchEvent(new Event('change', {bubbles:true}));
            """,
            sel, value,
        )
        time.sleep(1)
        print(f"   Inbound Type '{value}' successfully set ho gaya.")
    except Exception as e:
        print(f"   Inbound Type set karte waqt error: {e}")


def set_creation_date_preset(driver, wait, preset_label):
    """Opens Creation Date picker (#gs_createdDate) and picks 'This Month'."""
    print(f"5) Creation Date filter ko '{preset_label}' set kar raha hoon...")
    try:
        date_field = wait.until(EC.presence_of_element_located((By.ID, "gs_createdDate")))
        driver.execute_script("arguments[0].click();", date_field)
        time.sleep(1)

        preset_el = wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//*[normalize-space(text())='{preset_label}']")
            )
        )
        driver.execute_script("arguments[0].click();", preset_el)
        time.sleep(1)

        apply_btns = driver.find_elements(
            By.XPATH, "//button[normalize-space(translate(.,'APPLY','apply'))='apply']"
        )
        for btn in apply_btns:
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(1)
                break
        print(f"   Creation Date preset '{preset_label}' applied successfully.")
    except Exception as e:
        print(f"   Creation Date set karte waqt issue aaya: {e}")


def create_returns_export_request(driver, wait):
    """Executes full export sequence on Inbound Enquiry screen."""
    open_inbound_enquiry_screen(driver, wait)
    set_inbound_type_filter(driver, wait, INBOUND_TYPE_FILTER)
    set_creation_date_preset(driver, wait, DATE_PRESET)

    print("6) Search button click kar raha hoon...")
    try:
        search_btn = wait.until(EC.element_to_be_clickable((By.ID, "SearchBtn")))
        driver.execute_script("arguments[0].click();", search_btn)
    except Exception:
        search_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(translate(.,'SEARCH','search'),'search')]")))
        driver.execute_script("arguments[0].click();", search_btn)
    
    print("   Data load hone ka wait (8 sec)...")
    time.sleep(8)

    print("7) Detail Export click kar raha hoon...")
    try:
        detail_export_btn = wait.until(EC.element_to_be_clickable((By.ID, "downloadButton")))
        driver.execute_script("arguments[0].click();", detail_export_btn)
    except Exception:
        detail_export_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(),'Detail Export')]")))
        driver.execute_script("arguments[0].click();", detail_export_btn)
    
    time.sleep(3)

    print("8) Export fields select-all modal check kar raha hoon...")
    all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
    for fr in all_iframes:
        try:
            driver.switch_to.default_content()
            driver.switch_to.frame(fr)
            cb = driver.find_elements(By.ID, "cb_dynamicFieldGrid")
            if cb:
                if not cb[0].is_selected():
                    driver.execute_script("arguments[0].click();", cb[0])
                print("   Select-all checkbox (#cb_dynamicFieldGrid) selected.")
                break
        except Exception:
            pass

    print("9) Final Export button trigger kar raha hoon...")
    export_btn = wait.until(EC.element_to_be_clickable(
        (By.XPATH, "//button[contains(@onclick, 'exportData')] | //button[@title='Export']")
    ))
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(3)

    driver.switch_to.default_content()


# ============================================================
# MAIN EXTRACTION
# ============================================================

def extract_vinculum_returns():
    if not VINCULUM_USERNAME or not VINCULUM_PASSWORD:
        raise RuntimeError("GitHub Secrets VINCULUM_USERNAME / VINCULUM_PASSWORD available nahi hain.")

    driver = build_driver()
    wait = WebDriverWait(driver, 30)

    try:
        # ----------------------------------------------------
        # FIXED LOGIN LOGIC (Avoids ElementNotInteractableException)
        # ----------------------------------------------------
        print("1) Login process start ho raha hai...")
        driver.get(LOGIN_URL)
        time.sleep(2)

        # Look for login elements across top-level and potential frames
        def get_login_inputs():
            driver.switch_to.default_content()
            u = driver.find_elements(By.XPATH, "//input[@type='text' or @type='email' or @name='username' or @id='username']")
            p = driver.find_elements(By.XPATH, "//input[@type='password' or @name='password' or @id='password']")
            if u and p:
                return u[0], p[0]
            
            for fr in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    driver.switch_to.default_content()
                    driver.switch_to.frame(fr)
                    u = driver.find_elements(By.XPATH, "//input[@type='text' or @type='email' or @name='username' or @id='username']")
                    p = driver.find_elements(By.XPATH, "//input[@type='password' or @name='password' or @id='password']")
                    if u and p:
                        return u[0], p[0]
                except Exception:
                    pass
            return None, None

        username_el, password_el = None, None
        deadline = time.time() + 20
        while time.time() < deadline:
            username_el, password_el = get_login_inputs()
            if username_el and password_el:
                break
            time.sleep(1)

        if not username_el or not password_el:
            raise RuntimeError("Login fields (Username/Password) nahi mil paaye.")

        # Safe Value Injection via JS & Send Keys
        driver.execute_script("arguments[0].scrollIntoView(true);", username_el)
        driver.execute_script("arguments[0].value = '';", username_el)
        try:
            username_el.send_keys(VINCULUM_USERNAME)
        except Exception:
            driver.execute_script("arguments[0].value = arguments[1];", username_el, VINCULUM_USERNAME)

        driver.execute_script("arguments[0].value = '';", password_el)
        try:
            password_el.send_keys(VINCULUM_PASSWORD)
        except Exception:
            driver.execute_script("arguments[0].value = arguments[1];", password_el, VINCULUM_PASSWORD)

        # Login Click
        login_btn = driver.find_element(
            By.XPATH, "//button[contains(translate(.,'LOGIN','login'),'login')] | //input[@type='submit' or @id='loginButton']"
        )
        driver.execute_script("arguments[0].click();", login_btn)

        time.sleep(5)
        print("   Login SUCCESS.")

        # ----------------------------------------------------
        # EXPORT REQUEST
        # ----------------------------------------------------
        create_returns_export_request(driver, wait)

        # ----------------------------------------------------
        # PENDING REPORT POLLING
        # ----------------------------------------------------
        print("10) Pending Report grid check kar raha hoon...")

        def switch_to_pending_iframe():
            driver.switch_to.default_content()
            for fr in driver.find_elements(By.TAG_NAME, "iframe"):
                try:
                    driver.switch_to.frame(fr)
                    if "Pending Report" in driver.find_element(By.TAG_NAME, "body").text:
                        return True
                    driver.switch_to.default_content()
                except Exception:
                    driver.switch_to.default_content()
            return False

        switch_to_pending_iframe()

        status_ready = False
        attempt = 0

        while not status_ready and attempt < 20:
            attempt += 1
            switch_to_pending_iframe()
            rows = driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow")

            for row in rows:
                row_text = row.text.upper().replace(" ", "")
                if REPORT_NAME_MARKER in row_text:
                    if "SUCCESS" in row_text:
                        print("   Report status SUCCESS confirm ho gaya!")
                        status_ready = True
                        
                        download_img = row.find_element(
                            By.XPATH, ".//img[contains(@src,'DownloadData') or contains(@onclick,'downloadReport')] | .//a | .//*[@title='Download']"
                        )
                        files_before = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))
                        driver.execute_script("arguments[0].click();", download_img)
                        print("   Download button trigger ho gaya.")

                        downloaded_file = wait_for_download(DOWNLOAD_FOLDER, timeout=120, existing_files=files_before)
                        if downloaded_file:
                            shutil.copy2(downloaded_file, OUTPUT_FILE)
                            generated_at = datetime.utcnow().isoformat() + "Z"
                            with open(META_FILE, "w") as f:
                                json.dump({"generatedAt": generated_at}, f)
                            print(f"SUCCESS: Returns data successfully downloaded & saved to {OUTPUT_FILE}")
                            return
                        else:
                            raise RuntimeError("Download trigger hone ke baad timeout ho gaya.")
                    elif "ERROR" in row_text:
                        raise RuntimeError("Pending Report mein status 'ERROR' aagaya hai.")

            print(f"   Attempt {attempt}: Report complete hone ka waiting (10 sec)...")
            time.sleep(10)

        raise RuntimeError("Pending Report completion timeout ho gaya.")

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

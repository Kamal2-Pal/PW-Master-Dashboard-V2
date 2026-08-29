"""
Vinculum Data Extractor - GitHub Actions version
-------------------------------------------------
Automatically:
1. Login to Vinculum
2. Handle "already logged in" confirmation
3. Open Order Enquiry
4. Select the last 7 calendar days
5. Search
6. Detail Export -> select all fields
7. Wait for Pending Report = SUCCESS
8. Download Excel
9. Save the latest file as data.xlsx in the repository root
"""

import os

# Credentials are supplied by GitHub Actions environment variables.
VINCULUM_USERNAME = os.getenv("VINCULUM_USERNAME", "").strip()
VINCULUM_PASSWORD = os.getenv("VINCULUM_PASSWORD", "")

import time
import glob
import shutil
from datetime import datetime, timedelta

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ============================================================
# CONFIG
# ============================================================

LOGIN_URL = (
    "https://physicswallah.vineretail.com/"
    "eRetailWeb/eRetailLogin.action?popup=true"
)

# Credentials come from GitHub Repository Secrets.
USERNAME = os.environ["VINCULUM_USERNAME"]
PASSWORD = os.environ["VINCULUM_PASSWORD"]
DATE_MODE = os.environ.get("VINCULUM_DATE_MODE", "LAST_7_DAYS").upper()

if DATE_MODE not in {"LAST_7_DAYS", "YESTERDAY"}:
    raise RuntimeError(
        f"Unsupported VINCULUM_DATE_MODE={DATE_MODE!r}. "
        "Use LAST_7_DAYS or YESTERDAY."
    )

USERNAME_FIELD_ID = "userName"
PASSWORD_FIELD_ID = "password"
LOGIN_BUTTON_SELECTOR = "input[onclick*='doLoginJS']"

# GitHub runner folders
DOWNLOAD_FOLDER = os.path.abspath("data/downloads")
OUTPUT_FILE = os.path.abspath("data.xlsx")

# ============================================================
# ORDER ENQUIRY FILTERS
# ============================================================

CHANNELS = [
    "M10 - B2B_BOS",
    "M09 - B2B_DC",
    "M16 - B2B_EI",
    "M21 - B2B_INV_PWH",
    "M08 - B2B_TI",
    "PWH - Sikanderabad_FC1",
]

DATE_DAYS = 7


# ============================================================
# HELPERS
# ============================================================

def wait_for_download(folder, timeout=120):
    """Wait for a new completed download."""
    seconds = 0
    existing_files = set(glob.glob(os.path.join(folder, "*")))

    while seconds < timeout:
        time.sleep(1)
        seconds += 1

        current_files = set(glob.glob(os.path.join(folder, "*")))
        new_files = current_files - existing_files

        finished_files = [
            f for f in new_files
            if os.path.isfile(f) and not f.endswith(".crdownload")
        ]
        still_downloading = [
            f for f in new_files
            if f.endswith(".crdownload")
        ]

        if finished_files and not still_downloading:
            return max(finished_files, key=os.path.getmtime)

    return None


def build_driver():
    """Create Chrome in headless mode for GitHub Actions."""
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

    # Clean previous downloads.
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

    # GitHub-hosted Ubuntu runner
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    # Selenium Manager automatically manages ChromeDriver.
    return webdriver.Chrome(options=options)


def accept_existing_session_alert(driver):
    """
    Vinculum may show a native browser alert after login:
    'You are already logged in from IP ... Do you really want
    to end previous session and start new Login session?'

    Accept it so the GitHub runner can continue.
    """
    try:
        alert = WebDriverWait(driver, 15).until(EC.alert_is_present())
        print("Vinculum session alert found:")
        print(alert.text)

        alert.accept()

        print("Previous session confirmation accepted.")
        time.sleep(4)

        # A second alert is unlikely, but handle it if Vinculum sends one.
        try:
            second_alert = WebDriverWait(driver, 3).until(
                EC.alert_is_present()
            )
            print("Second Vinculum alert found:")
            print(second_alert.text)
            second_alert.accept()
            time.sleep(2)
        except Exception:
            pass

    except Exception:
        print("No 'already logged in' alert found. Continuing.")


def select_vinculum_channels(driver, channels):
    """
    Select all requested channels.
    First tries a native <select>/<option> control, then a custom
    multiselect by opening the Channel control and clicking exact labels.
    Fails loudly if all requested channels cannot be confirmed.
    """
    wanted = [c.strip().lower() for c in channels]

    def norm(v):
        return " ".join((v or "").split()).strip().lower()

    # Native select / multi-select
    best = None
    best_count = 0
    for sel in driver.find_elements(By.TAG_NAME, "select"):
        try:
            if not sel.is_displayed():
                continue
            options = sel.find_elements(By.TAG_NAME, "option")
            option_texts = {norm(o.text) for o in options}
            count = sum(1 for c in wanted if c in option_texts)
            if count > best_count:
                best_count = count
                best = sel
        except Exception:
            pass

    if best is not None and best_count >= 2:
        driver.execute_script("""
            const s = arguments[0];
            const wanted = arguments[1];
            for (const o of s.options) {
                o.selected = wanted.includes((o.textContent || "").trim().toLowerCase());
            }
            s.dispatchEvent(new Event("input", {bubbles:true}));
            s.dispatchEvent(new Event("change", {bubbles:true}));
        """, best, wanted)
        time.sleep(1)

        selected = driver.execute_script("""
            return Array.from(arguments[0].selectedOptions)
              .map(o => (o.textContent || "").trim().toLowerCase());
        """, best)

        missing = [channels[i] for i, c in enumerate(wanted) if c not in set(selected)]
        if not missing:
            print(f"   All {len(channels)} requested channels selected.")
            return

        print("   Native selector found, but missing:", missing)

    # Custom multiselect: try to open a visible Channel control.
    controls = driver.find_elements(
        By.XPATH,
        "//*[self::button or self::div or self::span or self::label or self::input]"
        "[contains(translate(concat(normalize-space(.),' ',@aria-label,' ',@title),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'channel')]"
    )

    for el in controls:
        try:
            if not el.is_displayed():
                continue
            txt = norm(el.text)
            aria = norm(el.get_attribute("aria-label"))
            title = norm(el.get_attribute("title"))
            if txt == "channel" or aria == "channel" or title == "channel"                or "select channel" in txt or "select channel" in aria or "select channel" in title:
                driver.execute_script("arguments[0].click();", el)
                time.sleep(0.5)
                break
        except Exception:
            pass

    clicked = set()
    for channel in channels:
        candidates = driver.find_elements(
            By.XPATH,
            f"//*[self::option or self::li or self::label or self::span or self::div or self::td or self::a]"
            f"[normalize-space(.)={repr(channel)}]"
        )
        for el in candidates:
            try:
                if not el.is_displayed():
                    continue
                if len((el.text or "").strip()) > len(channel) + 20:
                    continue
                driver.execute_script("arguments[0].click();", el)
                clicked.add(channel)
                time.sleep(0.15)
                break
            except Exception:
                pass

    if len(clicked) != len(channels):
        missing = [c for c in channels if c not in clicked]
        raise RuntimeError(
            "Required channel filter could not be confirmed. "
            f"Missing: {missing}"
        )

    print(f"   All {len(clicked)} requested channels selected.")


def create_order_export_request(driver, wait):
    """Create a fresh OrderEnquiryExport request using the configured 7-day/channel filters."""
    # 2. ORDER ENQUIRY
    # ----------------------------------------------------
    print("2) Order Enquiry screen khol raha hoon...")

    driver.execute_script(
        'openScreen("Order Enquiry", '
        '"orderEnquiryBS","fa fa-arrow-circle-right");'
    )

    time.sleep(5)

    # ----------------------------------------------------
    # 2b. ORDER ENQUIRY IFRAME
    # ----------------------------------------------------
    print("3) Order Enquiry iframe mein switch kar raha hoon...")

    try:
        iframe = wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "iframe[src*='orderEnquiry']")
            )
        )
        driver.switch_to.frame(iframe)
        print("   Iframe mein switch ho gaya.")

    except Exception:
        print(
            "   Matching iframe nahi mila; "
            "top-level page par try kar raha hoon."
        )

    # ----------------------------------------------------
    # 2c. LAST 7 DAYS
    # ----------------------------------------------------
    if DATE_MODE == "YESTERDAY":
        print("4) Date filter ko YESTERDAY set kar raha hoon...")
        today = datetime.now()
        start_date = today - timedelta(days=1)
        end_date = start_date
    else:
        print(f"4) Date filter ko LAST {DATE_DAYS} DAYS set kar raha hoon...")
        today = datetime.now()
        start_date = today - timedelta(days=DATE_DAYS - 1)
        end_date = today

    start_date_str = start_date.strftime("%d/%m/%Y")
    end_date_str = end_date.strftime("%d/%m/%Y")

    print(f"   Order Create Date: {start_date_str} to {end_date_str}")

    driver.execute_script(
        """
        var picker = $('#gs_orderDate').data('daterangepicker');
        if (!picker) throw new Error("Order date daterangepicker not found.");
        picker.setStartDate(arguments[0]);
        picker.setEndDate(arguments[1]);
        $('#gs_orderDate').trigger('apply.daterangepicker', picker);
        """,
        start_date_str,
        end_date_str,
    )

    time.sleep(1)

    # ----------------------------------------------------
    # 2d. CHANNEL FILTER
    # ----------------------------------------------------
    print("5) Required B2B channels select kar raha hoon...")
    select_vinculum_channels(driver, CHANNELS)

    # ----------------------------------------------------
    # 2e. SEARCH
    # ----------------------------------------------------
    print("6) Search button click kar raha hoon...")

    search_btn = wait.until(
        EC.element_to_be_clickable(
            (By.ID, "searchOrderButton")
        )
    )

    search_btn.click()

    print("   Data load hone ka wait (15 sec)...")
    time.sleep(15)

    # ----------------------------------------------------
    # SAFETY CHECK
    # ----------------------------------------------------
    try:
        error_box = driver.find_element(By.ID, "messageLabel")

        if error_box.is_displayed() and error_box.text.strip():
            raise RuntimeError(
                f"Vinculum error: {error_box.text.strip()}"
            )

    except RuntimeError:
        raise

    except Exception:
        pass

    # ----------------------------------------------------
    # 3. DETAIL EXPORT
    # ----------------------------------------------------
    print("7) Detail Export click kar raha hoon...")

    detail_export_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//*[contains(text(),'Detail Export')]")
        )
    )

    detail_export_btn.click()
    time.sleep(2)

    # ----------------------------------------------------
    # 4. EXPORT FIELD MODAL
    # ----------------------------------------------------
    print("8) Export fields select kar raha hoon...")

    all_modal_contents = driver.find_elements(
        By.CSS_SELECTOR, "div.modal-content"
    )

    modal_content = None

    for mc in all_modal_contents:
        if mc.is_displayed() and "Select Field For Export" in mc.text:
            modal_content = mc
            break

    if modal_content is None:
        raise RuntimeError(
            "Select Field For Export modal nahi mila."
        )

    nested_iframe = modal_content.find_element(
        By.CSS_SELECTOR, "iframe"
    )

    driver.switch_to.frame(nested_iframe)
    time.sleep(1)

    # ----------------------------------------------------
    # SELECT ALL
    # ----------------------------------------------------
    try:
        select_all_cb = driver.find_element(
            By.ID, "cb_dynamicFieldGrid"
        )

        if not select_all_cb.is_selected():
            driver.execute_script(
                "arguments[0].click();",
                select_all_cb
            )

        print("   Select-all checkbox click ho gaya.")

    except Exception:
        print("   Select-all fallback use kar raha hoon...")

        all_checkboxes = driver.find_elements(
            By.CSS_SELECTOR,
            "input[type='checkbox']"
        )

        print(
            f"   {len(all_checkboxes)} checkboxes mile."
        )

        for checkbox in all_checkboxes:
            if not checkbox.is_selected():
                driver.execute_script(
                    "arguments[0].click();",
                    checkbox
                )
                time.sleep(0.1)

    # ----------------------------------------------------
    # EXPORT
    # ----------------------------------------------------
    print("9) Export click kar raha hoon...")

    export_btn = wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//button[@title='Export']")
        )
    )

    export_btn.click()
    time.sleep(3)

    # Back to Order Enquiry parent iframe
    driver.switch_to.parent_frame()


    driver.switch_to.default_content()
    time.sleep(1)


# ============================================================
# MAIN EXTRACTION
# ============================================================

def extract_vinculum_data():

    if not VINCULUM_USERNAME or not VINCULUM_PASSWORD:
        raise RuntimeError(
            "GitHub Actions secrets VINCULUM_USERNAME / VINCULUM_PASSWORD "
            "available nahi hain. Repository Settings > Secrets and variables > Actions check karein."
        )
    driver = build_driver()
    wait = WebDriverWait(driver, 30)

    try:
        # ----------------------------------------------------
        # 1. LOGIN
        # ----------------------------------------------------
        print("1) Login ho raha hai...")
        print("1) Vinculum login page open kar raha hoon...")
        driver.get(LOGIN_URL)

        wait = WebDriverWait(driver, 30)

        def first_visible(selectors):
            for by, value in selectors:
                try:
                    for el in driver.find_elements(by, value):
                        if el.is_displayed() and el.is_enabled():
                            return el
                except Exception:
                    pass
            return None

        # Wait for login form or an already authenticated page.
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
            print("2) Username/password fill kar raha hoon...")
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

            print("3) Login button click kar raha hoon...")
            driver.execute_script("arguments[0].click();", login_btn)
        else:
            print("2) Existing Vinculum session detected.")

        print("4) Login ke baad page/popup check kar raha hoon...")
        login_deadline = time.time() + 30

        while time.time() < login_deadline:
            # Browser JavaScript alert
            try:
                alert = driver.switch_to.alert
                print("   Browser alert:", alert.text[:200])
                alert.accept()
                time.sleep(1)
                continue
            except Exception:
                pass

            # Visible HTML dialogs
            dialogs = driver.find_elements(
                By.CSS_SELECTOR,
                "[role='dialog'], .modal, .ui-dialog, .modal-dialog"
            )
            for dialog in dialogs:
                try:
                    if not dialog.is_displayed():
                        continue
                    buttons = dialog.find_elements(
                        By.XPATH,
                        ".//button | .//input[@type='button'] | .//input[@type='submit'] | .//a"
                    )
                    for btn in buttons:
                        if not btn.is_displayed() or not btn.is_enabled():
                            continue
                        label = " ".join(filter(None, [
                            btn.text,
                            btn.get_attribute("value"),
                            btn.get_attribute("aria-label"),
                            btn.get_attribute("title")
                        ])).strip().lower()
                        if any(k in label for k in [
                            "continue", "ok", "yes", "proceed",
                            "logout other", "terminate", "close"
                        ]):
                            print("   Popup action:", label[:80])
                            driver.execute_script("arguments[0].click();", btn)
                            time.sleep(1)
                            break
                except Exception:
                    pass

            try:
                url_now = driver.current_url.lower()
                body = driver.find_element(By.TAG_NAME, "body").text.lower()
                if (
                    "selcompanylocationbs.action" in url_now
                    or "pending report" in body
                    or "order enquiry" in body
                ):
                    print("   Login SUCCESS — Vinculum application page detected.")
                    break
            except Exception:
                pass

            time.sleep(1)

        # Do not silently continue if the login page is still active.
        try:
            final_url = driver.current_url
            final_body = driver.find_element(By.TAG_NAME, "body").text.lower()
        except Exception:
            final_url = driver.current_url
            final_body = ""

        if (
            "login" in final_url.lower()
            and "selcompanylocationbs.action" not in final_url.lower()
        ) or "invalid username" in final_body or "invalid password" in final_body:
            raise RuntimeError(
                "Vinculum login successful nahi hua. "
                f"Current URL: {final_url}"
            )

        print("   Login stage complete:", final_url)

        # ----------------------------------------------------
        # CREATE ORDER EXPORT REQUEST
        # ----------------------------------------------------
        print("2) Order Enquiry + filters + Detail Export start kar raha hoon...")
        current_report_id = create_order_export_request(driver, wait)

        # ----------------------------------------------------
        # 6. PENDING REPORT IFRAME
        # ----------------------------------------------------
        print("10) Pending Report iframe dhoondh raha hoon...")

        driver.switch_to.default_content()
        time.sleep(2)

        all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
        target_iframe = None

        for fr in all_iframes:
            src = (fr.get_attribute("src") or "").lower()

            if fr.is_displayed() and (
                "pending" in src or "export" in src
            ):
                target_iframe = fr
                break

        if target_iframe is None:
            displayed = [
                fr for fr in all_iframes
                if fr.is_displayed()
            ]

            if displayed:
                target_iframe = displayed[-1]

        if target_iframe is None:
            raise RuntimeError(
                "Pending Report iframe nahi mila."
            )

        driver.switch_to.frame(target_iframe)
        time.sleep(3)

        # ----------------------------------------------------
        # 7. WAIT FOR SUCCESS / RETRY ON ERROR
        # ----------------------------------------------------
        print("11) Report status ka wait kar raha hoon...")
        print("   Pending Report har 10 sec refresh hoga.")
        print("   SUCCESS milte hi download hoga; ERROR par automation fresh export karegi.")

        status_ready = False
        attempt = 0
        export_retry_count = 0
        MAX_EXPORT_RETRIES = 10

        def switch_to_pending_iframe():
            driver.switch_to.default_content()
            time.sleep(0.5)

            all_iframes = driver.find_elements(By.TAG_NAME, "iframe")
            target = None

            for fr in all_iframes:
                src = (fr.get_attribute("src") or "").lower()
                if fr.is_displayed() and ("pending" in src or "export" in src):
                    target = fr
                    break

            if target is None:
                displayed = [fr for fr in all_iframes if fr.is_displayed()]
                if displayed:
                    target = displayed[-1]

            if target is None:
                raise RuntimeError("Pending Report iframe nahi mila.")

            driver.switch_to.frame(target)

        def read_latest_order_export():
            """Read newest OrderEnquiryExport row."""
            rows = driver.find_elements(
                By.CSS_SELECTOR,
                "table.table-bordered tbody tr"
            )

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")
                if len(cells) < 7:
                    continue

                values = [c.text.strip() for c in cells]
                if "ORDERENQUIRYEXPORT" in values[3].upper():
                    return row, values

            return None, None

        while not status_ready:
            attempt += 1

            try:
                switch_to_pending_iframe()
                row, values = read_latest_order_export()

                if row is None:
                    print(f"   Attempt {attempt} - OrderEnquiryExport row abhi nahi mila.")
                else:
                    report_id = values[0]
                    status_text = values[1].strip().upper()
                    error_msg = values[6].strip()

                    print(
                        f"   Attempt {attempt} - Report {report_id} - "
                        f"Status: {status_text}"
                    )

                    if status_text == "SUCCESS":
                        current_report_id = report_id
                        status_ready = True
                        print(f"   SUCCESS mila - Report ID: {report_id}")
                        break

                    if status_text == "ERROR":
                        export_retry_count += 1
                        detail = error_msg or "Generic Business Error"
                        print(
                            f"   Report {report_id} ERROR: {detail}. "
                            f"Fresh export retry {export_retry_count}/{MAX_EXPORT_RETRIES}..."
                        )

                        if export_retry_count > MAX_EXPORT_RETRIES:
                            raise RuntimeError(
                                "Fresh export retries exhausted. "
                                f"Last Report ID: {report_id}; Error: {detail}"
                            )

                        # Create a completely new export request instead of stopping.
                        driver.switch_to.default_content()
                        time.sleep(1)
                        current_report_id = create_order_export_request(driver, wait)
                        attempt = 0
                        time.sleep(2)
                        continue

            except RuntimeError:
                raise
            except Exception as exc:
                print(
                    f"   Status read issue: {type(exc).__name__}: {exc}"
                )

            # Refresh Pending Report every 10 seconds.
            try:
                driver.switch_to.default_content()
                driver.execute_script(
                    "if (typeof refreshGrid === 'function') "
                    "{ refreshGrid(); } else { location.reload(); }"
                )
                print("   Pending Report refresh kiya (10 sec interval).")
            except Exception as exc:
                print(
                    f"   Refresh issue: {type(exc).__name__}: {exc}"
                )

            time.sleep(10)

        # ----------------------------------------------------
        # 8. DOWNLOAD
        # ----------------------------------------------------
        print("12) Download click kar raha hoon...")

        # Find the SUCCESS row again and click its actual download control.
        row, values = read_latest_order_export()
        if row is None:
            raise RuntimeError("SUCCESS report row download ke liye nahi mila.")

        download_controls = row.find_elements(
            By.CSS_SELECTOR,
            "label[onclick*='downloadReport']"
        )

        if not download_controls:
            download_controls = row.find_elements(
                By.CSS_SELECTOR,
                "label"
            )

        if not download_controls:
            raise RuntimeError(
                "SUCCESS report mein downloadReport wala control nahi mila."
            )

        driver.execute_script(
            "arguments[0].click();",
            download_controls[0]
        )
        print(f"   Download initiated for Report ID: {values[0]}")

        # ----------------------------------------------------
        # 9. WAIT FOR DOWNLOAD
        # ----------------------------------------------------
        print("13) Excel download hone ka wait...")

        downloaded_file = wait_for_download(
            DOWNLOAD_FOLDER,
            timeout=120
        )

        if not downloaded_file:
            raise RuntimeError(
                "Excel download timeout ho gaya."
            )

        # ----------------------------------------------------
        # 10. SAVE AS data.xlsx
        # ----------------------------------------------------
        shutil.copy2(downloaded_file, OUTPUT_FILE)

        print(
            f"SUCCESS: Latest Vinculum data saved as: "
            f"{OUTPUT_FILE}"
        )

        print(
            f"Downloaded source file: "
            f"{os.path.basename(downloaded_file)}"
        )

    except Exception as e:
        print("========================================")
        print("VINCULUM EXTRACTION FAILED")
        print(f"Error: {e}")
        print("========================================")
        raise

    finally:
        driver.quit()


if __name__ == "__main__":
    extract_vinculum_data()

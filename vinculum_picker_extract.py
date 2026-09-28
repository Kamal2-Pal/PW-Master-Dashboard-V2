"""
Vinculum Picker Report Extractor + Daily Archive
------------------------------------------------
Flow:
1. Login to Vinculum.
2. Open Pick Pack Report.
3. Select B2B Orders.
4. Select date range = last 3 calendar days.
5. Select Picked Report.
6. Click Print.
7. Wait for the generated download/report button; download automatically.
8. Save the raw 3-day report.
9. Keep ONLY yesterday's rows in picker_history.csv (upsert/deduplicate).
10. Keep today's rows separately in picker_live.csv for the dashboard.

IMPORTANT:
- This script does not modify the existing B2B/Returns files.
- The exact report columns are based on the uploaded real Picker Report sample.
- If the Vinculum UI selector changes, the script saves diagnostics and raises a clear error.
"""

import csv
import glob
import hashlib
import json
import os
import re
import shutil
import time
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

REPORT_URL = (
    "https://physicswallah.vineretail.com/"
    "eRetailWeb/selCompanyLocationBS.action"
)

USERNAME = os.environ["VINCULUM_USERNAME"]
PASSWORD = os.environ["VINCULUM_PASSWORD"]

DOWNLOAD_FOLDER = os.path.abspath("data/downloads_picker")
RAW_FILE = os.path.abspath("data/picker_raw_last3days.csv")
LIVE_FILE = os.path.abspath("data/picker_live.csv")
HISTORY_FILE = os.path.abspath("data/picker_history.csv")
META_FILE = os.path.abspath("data/picker-meta.json")
DIAGNOSTIC_FILE = os.path.abspath("picker_failure.html")

DOWNLOAD_TIMEOUT = 180
REPORT_WAIT_TIMEOUT = 180


EXPECTED_COLUMNS = [
    "Order No",
    "Extern Order No",
    "Order Date",
    "Allocation Date",
    "Order Type",
    "SKU Code",
    "SKU Description",
    "Size",
    "Pick Date",
    "Pick User",
    "Bin",
    "Qty",
    "Pack Size ",
    "Brand Code",
    "Primary Supplier",
    "Supplier SKU",
    "LPN",
    "Pack User",
    "Customer Name",
    "Picklist No.",
    "Source WH",
]


# ============================================================
# BASIC HELPERS
# ============================================================

def norm(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def visible_text(el):
    try:
        return norm(el.text)
    except Exception:
        return ""


def all_text(el):
    parts = [
        visible_text(el),
        norm(el.get_attribute("value")),
        norm(el.get_attribute("title")),
        norm(el.get_attribute("aria-label")),
        norm(el.get_attribute("name")),
        norm(el.get_attribute("id")),
        norm(el.get_attribute("class")),
        norm(el.get_attribute("onclick")),
    ]
    return " ".join(x for x in parts if x).strip()


def click_js(driver, element):
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center', inline:'center'});",
        element,
    )
    time.sleep(0.4)
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click();", element)


def first_visible(driver, locators):
    for by, selector in locators:
        try:
            for el in driver.find_elements(by, selector):
                if el.is_displayed() and el.is_enabled():
                    return el
        except Exception:
            pass
    return None


def save_diagnostic(driver, reason):
    try:
        with open(DIAGNOSTIC_FILE, "w", encoding="utf-8") as f:
            f.write("<!-- " + reason + " -->\n")
            f.write(driver.page_source)
    except Exception:
        pass


def clean_download_folder():
    os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)
    for p in glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")):
        try:
            if os.path.isfile(p):
                os.remove(p)
        except Exception:
            pass


def wait_for_new_download(before, timeout=DOWNLOAD_TIMEOUT):
    deadline = time.time() + timeout

    while time.time() < deadline:
        current = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))
        new_files = current - before

        # Do not accept Chrome's temporary partial download.
        finished = [
            p for p in new_files
            if os.path.isfile(p)
            and not p.endswith(".crdownload")
            and not p.endswith(".tmp")
        ]
        active = [
            p for p in current
            if p.endswith(".crdownload")
        ]

        if finished and not active:
            return max(finished, key=os.path.getmtime)

        time.sleep(1)

    return None


# ============================================================
# CHROME
# ============================================================

def build_driver():
    clean_download_folder()

    options = webdriver.ChromeOptions()
    options.add_experimental_option(
        "prefs",
        {
            "download.default_directory": DOWNLOAD_FOLDER,
            "download.prompt_for_download": False,
            "directory_upgrade": True,
            "safebrowsing.enabled": True,
        },
    )
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
    except Exception:
        pass

    return driver


# ============================================================
# LOGIN
# ============================================================

def handle_alerts(driver):
    try:
        while True:
            alert = driver.switch_to.alert
            alert.accept()
            time.sleep(0.7)
    except Exception:
        pass


def handle_common_dialogs(driver):
    for dialog in driver.find_elements(
        By.CSS_SELECTOR,
        "[role='dialog'], .modal, .ui-dialog, .modal-dialog"
    ):
        try:
            if not dialog.is_displayed():
                continue

            for btn in dialog.find_elements(
                By.XPATH,
                ".//button | .//input[@type='button'] | "
                ".//input[@type='submit'] | .//a"
            ):
                if not btn.is_displayed() or not btn.is_enabled():
                    continue

                label = all_text(btn).lower()
                if any(
                    word in label
                    for word in [
                        "continue", "ok", "yes", "proceed",
                        "logout other", "terminate",
                    ]
                ):
                    click_js(driver, btn)
                    time.sleep(1)
                    break
        except Exception:
            pass


def login(driver, wait):
    driver.get(LOGIN_URL)
    time.sleep(2)

    user = first_visible(
        driver,
        [
            (By.ID, "userName"),
            (By.NAME, "userName"),
            (By.CSS_SELECTOR, "input[name='username']"),
            (By.CSS_SELECTOR, "input[type='text']"),
        ],
    )
    pwd = first_visible(
        driver,
        [
            (By.ID, "password"),
            (By.NAME, "password"),
            (By.CSS_SELECTOR, "input[type='password']"),
        ],
    )

    if user and pwd:
        user.clear()
        user.send_keys(USERNAME)
        pwd.clear()
        pwd.send_keys(PASSWORD)

        login_btn = first_visible(
            driver,
            [
                (By.CSS_SELECTOR, "input[onclick*='doLoginJS']"),
                (By.ID, "loginButton"),
                (By.ID, "login"),
                (By.NAME, "login"),
                (By.CSS_SELECTOR, "button[type='submit']"),
                (By.CSS_SELECTOR, "input[type='submit']"),
            ],
        )

        if not login_btn:
            raise RuntimeError("Vinculum login button nahi mila.")

        click_js(driver, login_btn)

    deadline = time.time() + 45
    while time.time() < deadline:
        handle_alerts(driver)
        handle_common_dialogs(driver)

        try:
            url = driver.current_url.lower()
            body = driver.find_element(By.TAG_NAME, "body").text.lower()

            if (
                "selcompanylocationbs.action" in url
                or "pick pack" in body
                or "reports" in body
            ):
                print("Login SUCCESS.")
                return
        except Exception:
            pass

        time.sleep(1)

    raise RuntimeError("Vinculum login/app page timeout.")


# ============================================================
# PICK PACK REPORT
# ============================================================

def open_pick_pack_report(driver):
    """
    Open Pick Pack Report without navigating away from the logged-in
    application page.

    The previous version used driver.get(REPORT_URL) and then searched only
    for visible text "Pick Pack Report". In headless Vinculum, the report menu
    can be rendered as an icon/link whose visible text is not available.
    We therefore inspect the actual DOM for the site's own menu action
    (onclick/href/data attributes) containing Pick/Pack and click that element.
    """
    print("   Pick Pack Report menu locate kar raha hoon...")
    driver.switch_to.default_content()

    def report_screen_present():
        """Return True when the Pick Pack Report screen has actually loaded."""
        try:
            body = driver.find_element(By.TAG_NAME, "body").text.lower()
            markers = [
                "picked report",
                "packed report",
                "pick user",
                "picklist",
            ]
            if sum(1 for m in markers if m in body) >= 2:
                return True
        except Exception:
            pass

        # Common report controls; checked without assuming one exact ID.
        try:
            labels = driver.find_elements(By.TAG_NAME, "label")
            label_text = " ".join(
                visible_text(x).lower() for x in labels if x.is_displayed()
            )
            if "picked report" in label_text or "packed report" in label_text:
                return True
        except Exception:
            pass

        return False

    def click_actual_pick_pack_menu():
        """
        Search the current document for the real Pick Pack menu element.
        Priority:
        1) visible text/title/aria-label
        2) onclick containing openScreen + pick/pack
        3) href/data attributes containing pick/pack
        """
        # 1) Visible/menu metadata.
        elements = driver.find_elements(
            By.XPATH,
            "//*[self::a or self::li or self::span or self::div or self::label or self::button]"
        )
        for el in elements:
            try:
                if not el.is_displayed() or not el.is_enabled():
                    continue

                meta = " ".join([
                    visible_text(el),
                    norm(el.get_attribute("title")),
                    norm(el.get_attribute("aria-label")),
                    norm(el.get_attribute("data-original-title")),
                    norm(el.get_attribute("href")),
                    norm(el.get_attribute("onclick")),
                    norm(el.get_attribute("data-url")),
                    norm(el.get_attribute("data-screen")),
                ]).lower()

                if "pick" in meta and "pack" in meta:
                    click_js(driver, el)
                    return True
            except Exception:
                pass

        # 2) Directly inspect onclick attributes. This catches icon-only
        # Vinculum menu entries where the visible text is absent.
        try:
            onclick_elements = driver.find_elements(
                By.XPATH,
                "//*[@onclick]"
            )
            for el in onclick_elements:
                try:
                    if not el.is_displayed() or not el.is_enabled():
                        continue
                    onclick = norm(el.get_attribute("onclick")).lower()
                    if "pick" in onclick and "pack" in onclick:
                        click_js(driver, el)
                        return True
                except Exception:
                    pass
        except Exception:
            pass

        return False

    # Do NOT driver.get(REPORT_URL) here. Login already brought us into the
    # authenticated application; navigating directly can discard the menu
    # context/session state that contains the report action.
    deadline = time.time() + 45

    while time.time() < deadline:
        handle_alerts(driver)
        handle_common_dialogs(driver)

        # Top-level document.
        driver.switch_to.default_content()

        if report_screen_present():
            print("Pick Pack Report already open hai.")
            return

        if click_actual_pick_pack_menu():
            print("Pick Pack Report menu click ho gaya.")
            time.sleep(2)

            if report_screen_present():
                print("Pick Pack Report open ho gaya.")
                return

        # If the application/report is inside an iframe, search there too.
        # IMPORTANT: when the report is found inside the iframe, DO NOT switch
        # back to default_content(). The next step (B2B Orders) is inside the
        # same report iframe.
        driver.switch_to.default_content()
        frames = driver.find_elements(By.TAG_NAME, "iframe")

        for fr in frames:
            found_in_frame = False
            try:
                driver.switch_to.default_content()
                driver.switch_to.frame(fr)

                if report_screen_present():
                    print("Pick Pack Report iframe ke andar open hai.")
                    found_in_frame = True
                    return

                if click_actual_pick_pack_menu():
                    print("Pick Pack Report iframe menu click ho gaya.")
                    time.sleep(2)
                    if report_screen_present():
                        print("Pick Pack Report iframe se open ho gaya.")
                        found_in_frame = True
                        return
            except Exception:
                pass
            finally:
                if not found_in_frame:
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass

        # Last fallback: use the site's real openScreen function only when
        # its function exists AND the DOM exposes a Pick/Pack menu action.
        # We extract the actual argument list from that menu instead of
        # inventing a screen name.
        try:
            driver.switch_to.default_content()
            menu_data = driver.execute_script("""
                const els = Array.from(document.querySelectorAll('[onclick]'));
                for (const e of els) {
                    const o = (e.getAttribute('onclick') || '').toLowerCase();
                    const meta = [
                        e.innerText || '',
                        e.getAttribute('title') || '',
                        e.getAttribute('aria-label') || '',
                        o
                    ].join(' ').toLowerCase();
                    if (meta.includes('pick') && meta.includes('pack') &&
                        o.includes('openscreen')) {
                        return e.getAttribute('onclick');
                    }
                }
                return '';
            """)

            if menu_data:
                print("Actual Pick Pack openScreen action mila.")
                driver.execute_script(menu_data)
                time.sleep(2)
                if report_screen_present():
                    print("Pick Pack Report actual openScreen() se open ho gaya.")
                    return
        except Exception:
            pass

        time.sleep(1)

    save_diagnostic(driver, "Pick Pack Report screen not found after DOM/menu inspection")
    raise RuntimeError("Pick Pack Report screen nahi mila.")


def click_label_by_text(driver, text, exact=False):
    wanted = norm(text).lower()

    # Prefer labels because radio buttons in this screen are label-driven.
    labels = driver.find_elements(By.TAG_NAME, "label")
    for label in labels:
        try:
            if not label.is_displayed():
                continue
            t = visible_text(label).lower()
            ok = (t == wanted) if exact else (wanted in t)
            if ok:
                click_js(driver, label)
                return True
        except Exception:
            pass

    # Fallback: text element.
    elements = driver.find_elements(
        By.XPATH,
        "//*[contains(translate(normalize-space(.),"
        "'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),"
        f"'{wanted}')]",
    )
    for el in elements:
        try:
            if el.is_displayed() and el.is_enabled():
                click_js(driver, el)
                return True
        except Exception:
            pass

    return False


def click_option_across_frames(driver, text):
    """Click a report option in the current document or any child iframe.

    Vinculum opens the Pick Pack Report inside an iframe. The previous run
    proved that the report iframe was open, but the next step searched only
    the top-level document, so B2B Orders was not found. This helper searches
    the active document first and then child iframes, leaving Selenium in the
    frame where the option was actually clicked.
    """
    if click_label_by_text(driver, text):
        return True

    try:
        frames = driver.find_elements(By.TAG_NAME, "iframe")
    except Exception:
        frames = []

    for fr in frames:
        try:
            driver.switch_to.frame(fr)
            if click_label_by_text(driver, text):
                return True

            # One additional level is enough for the nested Vinculum report
            # frame seen in the live run.
            child_frames = driver.find_elements(By.TAG_NAME, "iframe")
            for child in child_frames:
                try:
                    driver.switch_to.frame(child)
                    if click_label_by_text(driver, text):
                        return True
                    driver.switch_to.parent_frame()
                except Exception:
                    try:
                        driver.switch_to.parent_frame()
                    except Exception:
                        pass
        except Exception:
            pass
        finally:
            # If the click succeeded, we intentionally stay in that frame.
            try:
                current = driver.execute_script("return window.frameElement;")
                if current is not None:
                    # Do not reset here; caller needs the successful frame.
                    pass
            except Exception:
                pass

        # Reset only when this frame did not contain the option.
        try:
            driver.switch_to.default_content()
        except Exception:
            pass

    return False


def select_b2b_orders(driver):
    if not click_option_across_frames(driver, "B2B Orders"):
        save_diagnostic(driver, "B2B Orders radio not found in report frame")
        raise RuntimeError("B2B Orders select nahi hua.")

    time.sleep(0.5)
    print("B2B Orders selected.")


def select_picked_report(driver):
    if not click_option_across_frames(driver, "Picked Report"):
        save_diagnostic(driver, "Picked Report radio not found in report frame")
        raise RuntimeError("Picked Report select nahi hua.")

    time.sleep(0.5)
    print("Picked Report selected.")


def find_date_input(driver):
    # First use input[type=date] if present.
    candidates = driver.find_elements(By.CSS_SELECTOR, "input[type='date']")
    for el in candidates:
        if el.is_displayed():
            return el

    # The inspected screen shows an input.form-control.active near the Date label.
    inputs = driver.find_elements(
        By.CSS_SELECTOR,
        "input.form-control, input.form-control.input-normal, input"
    )

    for el in inputs:
        try:
            if not el.is_displayed():
                continue

            typ = (el.get_attribute("type") or "").lower()
            if typ in {"radio", "checkbox", "button", "submit", "hidden"}:
                continue

            # Ignore Company/site select-related fields by nearby text/value.
            meta = all_text(el).lower()
            if "company" in meta or "site" in meta:
                continue

            return el
        except Exception:
            pass

    return None


def set_last_3_days(driver):
    # User's requested rule: today + previous 2 calendar days.
    today = datetime.now()
    start = today - timedelta(days=2)
    end = today

    # Screen shows one Date field. In this Vinculum screen that field can
    # be a daterangepicker; try its jQuery daterangepicker first.
    start_str = start.strftime("%d/%m/%Y")
    end_str = end.strftime("%d/%m/%Y")

    try:
        result = driver.execute_script(
            """
            var els = Array.from(document.querySelectorAll('input'));
            var target = els.find(function(e) {
                if (!e || e.offsetParent === null) return false;
                var t = (e.type || '').toLowerCase();
                if (['radio','checkbox','button','submit','hidden'].indexOf(t) >= 0) return false;
                return true;
            });
            if (!target) return 'NO_INPUT';

            var picker = window.jQuery && jQuery(target).data('daterangepicker');
            if (picker) {
                picker.setStartDate(arguments[0]);
                picker.setEndDate(arguments[1]);
                jQuery(target).trigger('apply.daterangepicker', picker);
                return 'DATERANGE_SET';
            }

            target.value = arguments[0] + ' - ' + arguments[1];
            target.dispatchEvent(new Event('input', {bubbles:true}));
            target.dispatchEvent(new Event('change', {bubbles:true}));
            return 'TEXT_SET';
            """,
            start_str,
            end_str,
        )
        if result in {"DATERANGE_SET", "TEXT_SET"}:
            print(f"Date filter set: {start_str} to {end_str}")
            return
    except Exception:
        pass

    inp = find_date_input(driver)
    if not inp:
        save_diagnostic(driver, "Date input not found")
        raise RuntimeError("Date input nahi mila.")

    try:
        inp.clear()
    except Exception:
        pass

    inp.send_keys(f"{start_str} - {end_str}")
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
        inp,
    )

    print(f"Date filter set: {start_str} to {end_str}")


# ============================================================
# PRINT / DOWNLOAD
# ============================================================

def find_print_button(driver):
    # Screenshot/inspect confirms: button#downloadButtonDetail ... fa-print
    candidates = driver.find_elements(By.ID, "downloadButtonDetail")
    for el in candidates:
        try:
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            pass

    candidates = driver.find_elements(
        By.XPATH,
        "//button[contains(@onclick,'Print') or "
        "contains(translate(normalize-space(.),'PRINT','print'),'print')]"
    )
    for el in candidates:
        try:
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            pass

    # Last fallback: visible button with print icon.
    for el in driver.find_elements(By.CSS_SELECTOR, "button, a, input"):
        try:
            if not el.is_displayed() or not el.is_enabled():
                continue
            meta = all_text(el).lower()
            if "print" in meta:
                return el
        except Exception:
            pass

    return None


def click_print_and_wait_for_download(driver):
    print_btn = find_print_button(driver)
    if not print_btn:
        save_diagnostic(driver, "Print button not found")
        raise RuntimeError("Print button nahi mila.")

    before = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))
    click_js(driver, print_btn)
    print("Print click ho gaya. Vinculum report generate hone ka wait...")

    # User observed 1–2 minutes. Do not stop early.
    downloaded = wait_for_new_download(before, timeout=REPORT_WAIT_TIMEOUT)

    if downloaded:
        print("Picker report download complete:", os.path.basename(downloaded))
        return downloaded

    # Some versions expose a download button after report generation instead
    # of starting the browser download immediately. Poll the DOM for it.
    deadline = time.time() + REPORT_WAIT_TIMEOUT
    while time.time() < deadline:
        candidates = driver.find_elements(
            By.XPATH,
            "//*[contains(translate(@id,'DOWNLOAD','download'),'download') "
            "or contains(translate(@class,'DOWNLOAD','download'),'download') "
            "or contains(translate(normalize-space(.),'DOWNLOAD','download'),'download')]"
        )
        for el in candidates:
            try:
                if not el.is_displayed() or not el.is_enabled():
                    continue
                meta = all_text(el).lower()
                if "download" in meta:
                    before = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))
                    click_js(driver, el)
                    downloaded = wait_for_new_download(
                        before, timeout=30
                    )
                    if downloaded:
                        print("Download button click se report mil gaya:",
                              os.path.basename(downloaded))
                        return downloaded
            except Exception:
                pass

        time.sleep(2)

    save_diagnostic(driver, "Picker report download timeout")
    raise RuntimeError(
        f"Print ke baad {REPORT_WAIT_TIMEOUT} sec mein Picker report download nahi hui."
    )


# ============================================================
# CSV NORMALIZATION / ARCHIVE
# ============================================================

def detect_delimiter(path):
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        sample = f.read(8192)
    try:
        return csv.Sniffer().sniff(sample, delimiters=",\t;|").delimiter
    except Exception:
        return ","


def read_report(path):
    delimiter = detect_delimiter(path)

    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=delimiter)
        rows = list(reader)
        fieldnames = reader.fieldnames or []

    normalized = {norm(x).lower(): x for x in fieldnames}

    missing = [
        col for col in EXPECTED_COLUMNS
        if col.lower() not in normalized
    ]
    if missing:
        raise RuntimeError(
            "Picker report columns mismatch. Missing: " + ", ".join(missing)
        )

    # Map exact expected spelling to actual CSV field spelling.
    rows2 = []
    for row in rows:
        out = {}
        for col in EXPECTED_COLUMNS:
            actual = normalized[col.lower()]
            out[col] = row.get(actual, "")
        rows2.append(out)

    return rows2


def parse_date_only(value):
    s = norm(value)
    # Actual sample format: 26/09/2026 12:40 43
    m = re.search(r"(\d{2})/(\d{2})/(\d{4})", s)
    if not m:
        return None
    return f"{m.group(3)}-{m.group(2)}-{m.group(1)}"


def row_key(row):
    # Stable line-level key. Same report row should not be archived twice.
    raw = "|".join([
        norm(row.get("Order No")),
        norm(row.get("SKU Code")),
        norm(row.get("Pick User")),
        norm(row.get("Pick Date")),
        norm(row.get("LPN")),
        norm(row.get("Picklist No.")),
        norm(row.get("Qty")),
    ])
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def write_csv(path, rows):
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=EXPECTED_COLUMNS,
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(rows)


def update_live_and_history(raw_rows):
    today = datetime.now().date()
    yesterday = today - timedelta(days=1)

    live_rows = []
    yesterday_rows = []

    for row in raw_rows:
        pick_date = parse_date_only(row.get("Pick Date"))
        if not pick_date:
            continue

        if pick_date == today.isoformat():
            live_rows.append(row)

        if pick_date == yesterday.isoformat():
            yesterday_rows.append(row)

    # Current date remains live only.
    write_csv(LIVE_FILE, live_rows)

    # History contains everything previously archived + ONLY yesterday
    # from today's 3-day source. No older dates from the source are added.
    existing = []
    if os.path.exists(HISTORY_FILE):
        try:
            existing = read_report(HISTORY_FILE)
        except Exception:
            existing = []

    merged = {}
    for row in existing:
        k = row_key(row)
        merged[k] = row

    for row in yesterday_rows:
        merged[row_key(row)] = row

    history_rows = list(merged.values())
    history_rows.sort(
        key=lambda r: (
            parse_date_only(r.get("Pick Date")) or "",
            norm(r.get("Order No")),
        )
    )

    write_csv(HISTORY_FILE, history_rows)

    meta = {
        "generatedAt": datetime.utcnow().isoformat() + "Z",
        "today": today.isoformat(),
        "yesterdayArchived": yesterday.isoformat(),
        "sourceRows": len(raw_rows),
        "liveRows": len(live_rows),
        "yesterdayRows": len(yesterday_rows),
        "historyRows": len(history_rows),
    }
    with open(META_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)

    print(
        f"Archive complete: source={len(raw_rows)}, "
        f"yesterday added/updated={len(yesterday_rows)}, "
        f"live today={len(live_rows)}, history total={len(history_rows)}"
    )


# ============================================================
# MAIN
# ============================================================

def main():
    driver = build_driver()

    try:
        wait = WebDriverWait(driver, 30)

        print("1) Vinculum login...")
        login(driver, wait)

        print("2) Pick Pack Report open...")
        open_pick_pack_report(driver)

        print("3) B2B Orders...")
        select_b2b_orders(driver)

        print("4) Last 3 days...")
        set_last_3_days(driver)

        print("5) Picked Report...")
        select_picked_report(driver)

        print("6) Print + automatic download...")
        downloaded = click_print_and_wait_for_download(driver)

        # Normalize the downloaded file to a stable raw CSV name.
        shutil.copy2(downloaded, RAW_FILE)

        print("7) Downloaded report read/process...")
        rows = read_report(RAW_FILE)

        print("8) Live + yesterday-only archive update...")
        update_live_and_history(rows)

        print("SUCCESS: Picker extraction complete.")

    except Exception as exc:
        try:
            save_diagnostic(driver, str(exc))
        except Exception:
            pass
        print("========================================")
        print("VINCULUM PICKER EXTRACTION FAILED")
        print("Error:", exc)
        print("========================================")
        raise

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()

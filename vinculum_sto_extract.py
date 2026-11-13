"""Export the recent Sikanderabad STO dispatch report from Vinculum.

The existing order enquiry export remains separate. This script runs after it
in GitHub Actions and writes sto-data.xlsx for the dashboard's STO summary.
"""
import glob
import json
import os
import re
import shutil
import time
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from vinculum_extract_final import (
    LOGIN_URL,
    USERNAME,
    PASSWORD,
    DOWNLOAD_FOLDER,
    build_driver,
    wait_for_download,
    accept_existing_session_alert,
)

OUTPUT_FILE = os.path.abspath("sto-data.xlsx")
META_FILE = os.path.abspath("sto-meta.json")
ORIGIN_SITE = "Sikanderabad_FC1"
STATUS_WORDS = {"SUCCESS", "ERROR", "WIP", "PENDING", "FAILED", "PROCESSING"}


def first_visible(driver, selectors):
    for by, selector in selectors:
        try:
            for element in driver.find_elements(by, selector):
                if element.is_displayed() and element.is_enabled():
                    return element
        except Exception:
            pass
    return None


def login(driver, wait):
    driver.get(LOGIN_URL)
    username_selectors = [
        (By.ID, "username"), (By.NAME, "username"),
        (By.ID, "j_username"), (By.NAME, "j_username"),
        (By.CSS_SELECTOR, "input[type='text']"),
        (By.CSS_SELECTOR, "input[type='email']"),
    ]
    wait.until(lambda d:
        first_visible(d, username_selectors) is not None
        or "selcompanylocationbs.action" in d.current_url.lower()
    )
    user = first_visible(driver, username_selectors)
    password = first_visible(driver, [
        (By.ID, "password"), (By.NAME, "password"),
        (By.ID, "j_password"), (By.NAME, "j_password"),
        (By.CSS_SELECTOR, "input[type='password']"),
    ])
    if user and password:
        user.clear()
        user.send_keys(USERNAME)
        password.clear()
        password.send_keys(PASSWORD)
        button = first_visible(driver, [
            (By.ID, "loginButton"), (By.ID, "login"),
            (By.NAME, "login"), (By.CSS_SELECTOR, "button[type='submit']"),
            (By.CSS_SELECTOR, "input[type='submit']"),
        ])
        if not button:
            raise RuntimeError("Vinculum login button nahi mila.")
        driver.execute_script("arguments[0].click();", button)

    accept_existing_session_alert(driver)
    deadline = time.time() + 40
    while time.time() < deadline:
        try:
            alert = driver.switch_to.alert
            alert.accept()
        except Exception:
            pass
        try:
            if "selcompanylocationbs.action" in driver.current_url.lower():
                return
        except Exception:
            pass
        time.sleep(1)
    raise RuntimeError("Vinculum login ke baad dashboard page nahi mila.")


def switch_to_sto_frame(driver, wait):
    driver.switch_to.default_content()
    try:
        frame = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, "iframe[src*='stoOrderEnquiry']")
        ))
        driver.switch_to.frame(frame)
        return
    except Exception:
        pass

    deadline = time.time() + 30
    while time.time() < deadline:
        driver.switch_to.default_content()
        for frame in driver.find_elements(By.TAG_NAME, "iframe"):
            if not frame.is_displayed():
                continue
            try:
                driver.switch_to.frame(frame)
                body = driver.find_element(By.TAG_NAME, "body").text.lower()
                if "sto order enquiry" in body or "origin site" in body:
                    return
            except Exception:
                pass
            driver.switch_to.default_content()
        time.sleep(1)
    raise RuntimeError("STO Order Enquiry iframe nahi mila.")


def set_sto_filters(driver):
    today = datetime.now(ZoneInfo("Asia/Kolkata")).date()
    start = today - timedelta(days=14)
    start_text = start.strftime("%d/%m/%Y")
    end_text = today.strftime("%d/%m/%Y")

    set_range = driver.execute_script("""
      const inputs = Array.from(document.querySelectorAll('input'));
      let target = document.querySelector('#gs_stoDate');
      let picker = target && window.jQuery ? jQuery(target).data('daterangepicker') : null;
      if (!picker && window.jQuery) {
        for (const el of inputs) {
          const p = jQuery(el).data('daterangepicker');
          if (p) { target = el; picker = p; break; }
        }
      }
      if (!picker) return false;
      picker.setStartDate(arguments[0]);
      picker.setEndDate(arguments[1]);
      jQuery(target).trigger('apply.daterangepicker', picker);
      return true;
    """, start_text, end_text)
    if not set_range:
        raise RuntimeError("STO Date range filter nahi mila.")
    print(f"STO Date selected: {start_text} to {end_text}")

    origin_select = None
    for select in driver.find_elements(By.TAG_NAME, "select"):
        if not select.is_displayed():
            continue
        try:
            selector_meta = " ".join([
                select.get_attribute("id") or "",
                select.get_attribute("name") or "",
                select.get_attribute("aria-label") or "",
                select.get_attribute("title") or "",
            ]).lower()
            # Do not mistake Destination Site's dropdown for Origin Site just
            # because it also contains Sikanderabad as a possible destination.
            if "origin" not in selector_meta:
                continue
            options = select.find_elements(By.TAG_NAME, "option")
            if any(
                re.sub(r"[^a-z0-9]", "", o.text.lower()).endswith(
                    re.sub(r"[^a-z0-9]", "", ORIGIN_SITE.lower())
                )
                for o in options
            ):
                origin_select = select
                break
        except Exception:
            pass
    selected = False
    if origin_select is not None:
        selected = driver.execute_script("""
          const s = arguments[0], wanted = arguments[1].toLowerCase().replace(/[^a-z0-9]/g, '');
          const o = Array.from(s.options).find(x => x.text.trim().toLowerCase().replace(/[^a-z0-9]/g, '').endsWith(wanted));
          if (!o) return false;
          s.value = o.value;
          o.selected = true;
          s.dispatchEvent(new Event('input', {bubbles:true}));
          s.dispatchEvent(new Event('change', {bubbles:true}));
          return Array.from(s.selectedOptions).some(x => x.text.trim().toLowerCase().replace(/[^a-z0-9]/g, '').endsWith(wanted));
        """, origin_select, ORIGIN_SITE)

    if not selected:
        origin_input = first_visible(driver, [
            (By.CSS_SELECTOR, "input[id*='origin'], input[name*='origin'], input[placeholder*='Origin'], input[aria-label*='Origin']"),
        ])
        if not origin_input:
            for input_element in driver.find_elements(By.CSS_SELECTOR, "input"):
                try:
                    value = input_element.get_attribute("value") or ""
                    if re.sub(r"[^a-z0-9]", "", value.lower()).endswith("sikanderabadfc1"):
                        origin_input = input_element
                        break
                except Exception:
                    pass
        if not origin_input:
            headers = driver.find_elements(By.XPATH, "//th")
            for header in headers:
                if "origin site" not in (header.text or "").strip().lower():
                    continue
                try:
                    header_row = header.find_element(By.XPATH, "ancestor::tr[1]")
                    header_cells = header_row.find_elements(By.XPATH, "./th | ./td")
                    column_index = next(i for i, cell in enumerate(header_cells, start=1) if cell == header)
                    filter_row = header_row.find_element(By.XPATH, "following-sibling::tr[1]")
                    origin_input = filter_row.find_element(By.XPATH, f"./td[{column_index}]//input")
                    break
                except Exception:
                    pass
        if origin_input:
            value = origin_input.get_attribute("value") or ""
            if not re.sub(r"[^a-z0-9]", "", value.lower()).endswith("sikanderabadfc1"):
                origin_input.clear()
                origin_input.send_keys(ORIGIN_SITE)
                time.sleep(0.5)
                origin_input.send_keys(Keys.ARROWDOWN)
                origin_input.send_keys(Keys.ENTER)
                value = origin_input.get_attribute("value") or ""
            if re.sub(r"[^a-z0-9]", "", value.lower()).endswith("sikanderabadfc1"):
                selected = True
                try:
                    origin_input.send_keys(Keys.ENTER)
                except Exception:
                    pass
    if not selected:
        raise RuntimeError(f"Origin Site {ORIGIN_SITE} select confirm nahi hua.")
    print(f"Origin Site selected: {ORIGIN_SITE}")

    search = first_visible(driver, [
        (By.XPATH, "//*[self::button or self::input or self::a][contains(translate(normalize-space(.) ,'SEARCH','search'),'search') or contains(translate(@value,'SEARCH','search'),'search') or contains(translate(@title,'SEARCH','search'),'search')]"),
    ])
    if not search:
        raise RuntimeError("STO Order Enquiry Search button nahi mila.")
    driver.execute_script("arguments[0].click();", search)
    time.sleep(12)


def request_sto_detail_export(driver, wait):
    driver.switch_to.default_content()
    driver.execute_script(
        'openScreen("STO Order Enquiry", "stoOrderEnquiryBS", "fa fa-arrow-circle-right");'
    )
    time.sleep(5)
    switch_to_sto_frame(driver, wait)
    set_sto_filters(driver)

    detail = wait.until(EC.element_to_be_clickable((
        By.XPATH,
        "//*[self::button or self::a or self::input][contains(translate(normalize-space(.) ,'DETAIL EXPORT','detail export'),'detail export') or contains(translate(@value,'DETAIL EXPORT','detail export'),'detail export') or contains(translate(@title,'DETAIL EXPORT','detail export'),'detail export')]"
    )))
    driver.execute_script("arguments[0].click();", detail)
    time.sleep(2)

    modal = None
    for candidate in driver.find_elements(By.CSS_SELECTOR, "div.modal-content"):
        if candidate.is_displayed() and "Select Field For Export" in candidate.text:
            modal = candidate
            break
    if modal is None:
        raise RuntimeError("STO Detail Export fields modal nahi mila.")

    frame = modal.find_element(By.CSS_SELECTOR, "iframe")
    driver.switch_to.frame(frame)
    time.sleep(1)
    check_all = first_visible(driver, [(By.ID, "cb_dynamicFieldGrid")])
    if check_all and not check_all.is_selected():
        driver.execute_script("arguments[0].click();", check_all)
    elif not check_all:
        checkboxes = driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']")
        if not checkboxes:
            raise RuntimeError("STO export field checkboxes nahi mile.")
        for checkbox in checkboxes:
            if not checkbox.is_selected():
                driver.execute_script("arguments[0].click();", checkbox)

    export = wait.until(EC.element_to_be_clickable((
        By.XPATH, "//button[@title='Export']"
    )))
    driver.execute_script("arguments[0].click();", export)
    driver.switch_to.default_content()
    print("STO Detail Export request submitted.")


def switch_to_pending_report(driver):
    driver.switch_to.default_content()
    for frame in driver.find_elements(By.TAG_NAME, "iframe"):
        if not frame.is_displayed():
            continue
        try:
            driver.switch_to.frame(frame)
            body = driver.find_element(By.TAG_NAME, "body").text.lower()
            if "pending report" in body or "report id" in body:
                return True
        except Exception:
            pass
        driver.switch_to.default_content()
    return False


def parse_report_row(row):
    cells = row.find_elements(By.TAG_NAME, "td")
    texts = [c.text.strip() for c in cells]
    status = next((t.upper() for t in texts if t.upper() in STATUS_WORDS), "")
    report_id = next((t for t in texts if t.isdigit() and len(t) >= 4), "")
    joined = "".join(texts).upper().replace(" ", "")
    return texts, report_id, status, joined


def wait_for_sto_report(driver):
    deadline = time.time() + 20 * 60
    matched = None
    while time.time() < deadline:
        if not switch_to_pending_report(driver):
            time.sleep(2)
            continue
        rows = driver.find_elements(By.CSS_SELECTOR, "tr.jqgrow")
        sto_rows = []
        for row in rows:
            texts, report_id, status, joined = parse_report_row(row)
            if report_id and "STO" in joined and "EXPORT" in joined:
                sto_rows.append((row, report_id, status, texts))
        if sto_rows:
            row, report_id, status, texts = max(sto_rows, key=lambda item: int(item[1]))
            print(f"STO export report {report_id}: {status or 'processing'}")
            if status == "SUCCESS":
                matched = (row, report_id)
                break
            if status in {"ERROR", "FAILED"}:
                raise RuntimeError("STO export failed: " + " | ".join(texts))
        try:
            driver.execute_script("if (typeof refreshGrid === 'function') refreshGrid();")
        except Exception:
            pass
        time.sleep(10)

    if not matched:
        raise RuntimeError("STO Detail Export report 20 minutes mein SUCCESS nahi hua.")
    return matched


def download_sto_report(driver, row, report_id):
    files_before = set(glob.glob(os.path.join(DOWNLOAD_FOLDER, "*")))
    download_control = None
    for element in row.find_elements(
        By.XPATH, ".//a | .//button | .//img | .//i | .//span"
    ):
        meta = " ".join([
            element.text or "",
            element.get_attribute("title") or "",
            element.get_attribute("aria-label") or "",
            element.get_attribute("alt") or "",
            element.get_attribute("class") or "",
        ]).lower()
        if "download" in meta:
            download_control = element
            break
    if not download_control:
        raise RuntimeError(f"STO report {report_id} ka download control nahi mila.")
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", download_control)
    try:
        download_control.click()
    except Exception:
        driver.execute_script("arguments[0].click();", download_control)

    downloaded = wait_for_download(DOWNLOAD_FOLDER, timeout=180, existing_files=files_before)
    if not downloaded:
        raise RuntimeError("STO report Excel download timeout ho gaya.")
    shutil.copy2(downloaded, OUTPUT_FILE)
    with open(META_FILE, "w", encoding="utf-8") as meta:
        json.dump({"generatedAt": datetime.now(timezone.utc).isoformat()}, meta)
    print(f"SUCCESS: STO report {report_id} saved to {OUTPUT_FILE}")


def main():
    if not USERNAME or not PASSWORD:
        raise RuntimeError("VINCULUM_USERNAME/PASSWORD secrets available nahi hain.")
    driver = build_driver()
    wait = WebDriverWait(driver, 40)
    try:
        login(driver, wait)
        request_sto_detail_export(driver, wait)
        row, report_id = wait_for_sto_report(driver)
        download_sto_report(driver, row, report_id)
    except Exception:
        raise
    finally:
        driver.quit()


if __name__ == "__main__":
    main()

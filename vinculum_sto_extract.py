def _modal_in_current_context(driver):
    """Current frame me export modal dhundhein; (modal, inner_iframe) ya (None, None)."""
    for candidate in driver.find_elements(By.CSS_SELECTOR, "div.modal-content"):
        try:
            if not candidate.is_displayed():
                continue
        except Exception:
            continue
        text = (candidate.text or "").lower()
        inner = None
        try:
            inner = candidate.find_element(By.TAG_NAME, "iframe")
        except Exception:
            inner = None
        # Purana tight match ("select field for export") + relaxed match:
        # koi bhi visible modal jisme export/field likha ho ya jo iframe rakhta ho.
        if inner is not None or "export" in text or "field" in text:
            return candidate, inner
    for fr in driver.find_elements(By.TAG_NAME, "iframe"):
        try:
            if not fr.is_displayed():
                continue
        except Exception:
            continue
        meta = " ".join([
            fr.get_attribute("src") or "",
            fr.get_attribute("id") or "",
            fr.get_attribute("name") or "",
        ]).lower()
        if "export" in meta or "field" in meta:
            return fr, fr
    try:
        cb = driver.find_element(By.ID, "cb_dynamicFieldGrid")
        if cb.is_displayed():
            return cb, None
    except Exception:
        pass
    return None, None


def _enter_frame_path(driver, path):
    driver.switch_to.default_content()
    for idx in path:
        frames = driver.find_elements(By.TAG_NAME, "iframe")
        if idx >= len(frames):
            raise IndexError(f"frame index {idx} out of range")
        driver.switch_to.frame(frames[idx])


def find_export_modal(driver, timeout=90):
    """Top + 2 levels ke iframes me export modal poll karein, alerts clear karte hue."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        paths = [[]]
        try:
            driver.switch_to.default_content()
            top_count = len(driver.find_elements(By.TAG_NAME, "iframe"))
        except Exception:
            top_count = 0
        for i in range(top_count):
            paths.append([i])
            try:
                _enter_frame_path(driver, [i])
                nested_count = len(driver.find_elements(By.TAG_NAME, "iframe"))
            except Exception:
                nested_count = 0
            for j in range(nested_count):
                paths.append([i, j])
        for path in paths:
            try:
                _enter_frame_path(driver, path)
            except Exception:
                continue
            modal, inner = _modal_in_current_context(driver)
            if modal is not None:
                label = "top" if not path else "frame" + ".".join(str(p) for p in path)
                print(f"STO Detail Export modal mil gaya (context: {label}).")
                return modal, inner
        time.sleep(2)
    return None, None


def dump_modal_debug(driver):
    """Failure par screenshot + page-source save karein (workflow artifact me upload hoga)."""
    try:
        driver.switch_to.default_content()
        driver.save_screenshot("sto_modal_debug.png")
        with open("sto_modal_debug_top.html", "w", encoding="utf-8") as fh:
            fh.write(driver.page_source)
        for i in range(len(driver.find_elements(By.TAG_NAME, "iframe"))):
            try:
                _enter_frame_path(driver, [i])
                with open(f"sto_modal_debug_frame{i}.html", "w", encoding="utf-8") as fh:
                    fh.write(driver.page_source)
            except Exception:
                pass
        driver.switch_to.default_content()
        print("Debug files saved: sto_modal_debug.png / sto_modal_debug_*.html")
    except Exception as exc:
        print(f"Debug dump failed: {exc}")


def request_sto_detail_export(driver, wait):
    driver.switch_to.default_content()
    driver.execute_script(
        'openScreen("STO Order Enquiry", "stockTransferEnquiryActionBS", "fa fa-arrow-circle-right");'
    )
    time.sleep(5)
    switch_to_sto_frame(driver, wait)
    set_sto_filters(driver)
    detail = first_visible(driver, [(By.CSS_SELECTOR, "button[title='Detail Export']")])
    if not detail:
        detail = wait.until(EC.element_to_be_clickable((
            By.XPATH,
            "//*[self::button or self::a or self::input][contains(translate(normalize-space(.),'DETAIL EXPORT','detail export'),'detail export') or contains(translate(@value,'DETAIL EXPORT','detail export'),'detail export') or contains(translate(@title,'DETAIL EXPORT','detail export'),'detail export')]"
        )))
    detail_handler = (detail.get_attribute("onclick") or "").strip()
    print(f"STO Detail Export handler: {detail_handler or '(none on button)'}")

    def trigger_via_handler():
        driver.execute_script("doDetailExport();")

    def trigger_via_click():
        driver.execute_script("arguments[0].click();", detail)

    if "detailexport" in detail_handler.lower():
        first_trigger, second_trigger = trigger_via_handler, trigger_via_click
    else:
        first_trigger, second_trigger = trigger_via_click, trigger_via_handler

    first_trigger()
    modal, inner = find_export_modal(driver, timeout=45)
    if modal is None:
        print("Pehla trigger modal nahi khol paya; doosra trigger try kar rahe hain.")
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        try:
            second_trigger()
        except Exception as exc:
            print(f"Second trigger warning: {exc}")
        modal, inner = find_export_modal(driver, timeout=60)
    if modal is None:
        dump_modal_debug(driver)
        raise RuntimeError("STO Detail Export fields modal nahi mila (debug files artifact me upload honge).")

    if inner is not None and inner.tag_name.lower() == "iframe":
        driver.switch_to.frame(inner)
        time.sleep(1)

    check_all = first_visible(driver, [(By.ID, "cb_dynamicFieldGrid")])
    if check_all and not check_all.is_selected():
        driver.execute_script("arguments[0].click();", check_all)
    if not check_all:
        checkboxes = [cb for cb in driver.find_elements(By.CSS_SELECTOR, "input[type='checkbox']") if cb.is_displayed()]
        if not checkboxes:
            dump_modal_debug(driver)
            raise RuntimeError("STO export field checkboxes nahi mile.")
        for checkbox in checkboxes:
            if not checkbox.is_selected():
                driver.execute_script("arguments[0].click();", checkbox)
    export = first_visible(driver, [(By.CSS_SELECTOR, "button[title='Export']")])
    if not export:
        export = first_visible(driver, [(
            By.XPATH,
            "//button[contains(translate(normalize-space(.),'EXPORT','export'),'export')] | //a[contains(translate(normalize-space(.),'EXPORT','export'),'export')]"
        )])
    if not export:
        dump_modal_debug(driver)
        raise RuntimeError("STO export modal me Export button nahi mila.")
    driver.execute_script("arguments[0].click();", export)
    driver.switch_to.default_content()
    print("STO Detail Export request submitted.")

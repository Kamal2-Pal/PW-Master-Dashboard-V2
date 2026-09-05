"""
Sends an "Open Orders" summary email — Open <48H, At Risk, and Open >48H
combined — mirroring the exact SLA logic used in the dashboard (index.html
buildSlaRows()). Meant to run as a step in the same GitHub Actions workflow
that produces data.xlsx, right after extraction, so the numbers always match
what the dashboard shows.

Sends via Brevo's transactional email API (https://www.brevo.com) - free
tier, no domain/DNS verification needed, just a single verified sender email.

Required GitHub Secrets (Settings -> Secrets and variables -> Actions):
    BREVO_API_KEY       - API key from Brevo (SMTP & API -> API Keys)
    BREVO_SENDER_EMAIL   - the email address you verified as a Sender in Brevo
    ALERT_RECIPIENTS     - comma-separated list of recipient emails,
                            e.g. "person1@pw.live,person2@pw.live"

One-time Brevo setup:
    1. Sign up at https://app.brevo.com/signup
    2. Senders, Domains & Dedicated IPs -> Senders -> Add a Sender
       -> verify that email via the confirmation link Brevo sends it
    3. Profile icon (top-right) -> SMTP & API -> API Keys -> Generate a new API key
"""

import os
import sys
import glob
import json
import re
import urllib.request
import urllib.error
from datetime import datetime, date, timedelta, timezone

import openpyxl

IST = timezone(timedelta(hours=5, minutes=30))


def now_ist():
    """GitHub Actions runners run in UTC. The dashboard shows IST (browser's
    local timezone), so without this conversion 'now' was 5.5 hours behind
    real IST time - causing the email's 'As of ...' timestamp to be off by
    exactly 5h30m from when it actually arrived, and making elapsed-hours
    SLA calculations subtly wrong too."""
    return datetime.now(timezone.utc).astimezone(IST).replace(tzinfo=None)

ORDER_ID_FIELDS = ["Order No", "OrderNo", "Order Number", "Order_No", "Order_Number", "Order ID", "OrderID"]
DATE_FIELDS = ["Order create date", "OrderDate", "Order_Date", "Date"]
SHIP_DATE_FIELDS = ["Actual_ShipDate", "Ship_Date", "Shipped_Date"]
MANIFEST_DATE_FIELDS = ["Manifest Create Date", "Manifest_Create_Date", "ManifestCreateDate"]
WAREHOUSE_FIELDS = ["FulfillmentLocationName", "Warehouse", "Location"]
SHIPPED_STATUSES = {"shipped complete", "partially shipped", "delivered", "shipped & returned"}


def norm(v):
    return str(v).strip() if v is not None else ""


def get(row, names):
    for n in names:
        if row.get(n) not in (None, ""):
            return row[n]
    lower_map = {str(k).lower().replace(" ", "").replace("_", ""): k for k in row.keys()}
    for n in names:
        key = n.lower().replace(" ", "").replace("_", "")
        if key in lower_map and row.get(lower_map[key]) not in (None, ""):
            return row[lower_map[key]]
    return ""


def parse_date(v):
    if v in (None, ""):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    s = str(v).strip()
    # Match a YYYY-MM-DD (or YYYY/MM/DD) prefix and ignore anything after it
    # (time, fractional seconds like ".0", etc.) - mirrors the dashboard's
    # dateVal() regex approach, which is what was missing here. The old
    # exact-format strptime() list failed on real export values like
    # "2026-09-01 07:50:17.0" (trailing ".0"), silently turning every row's
    # date into None - which made every order get skipped and produced the
    # "0 open orders" bug.
    m = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if m:
        try:
            return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            pass
    m = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})", s)
    if m:
        try:
            return datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            pass
    try:
        return datetime.fromisoformat(s)
    except ValueError:
        return None


def order_id_of(row):
    return norm(get(row, ORDER_ID_FIELDS))


def read_xlsx_best_sheet(path):
    """Same idea as the dashboard's parseWorkbookBuffer(): pick whichever
    sheet in the workbook has the most data rows."""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    best_rows = []
    for ws in wb.worksheets:
        rows_iter = ws.iter_rows(values_only=True)
        try:
            headers = next(rows_iter)
        except StopIteration:
            continue
        headers = [str(h).strip() if h is not None else "" for h in headers]
        sheet_rows = []
        for values in rows_iter:
            record = {h: v for h, v in zip(headers, values) if h}
            sheet_rows.append(record)
        if len(sheet_rows) > len(best_rows):
            best_rows = sheet_rows
    return best_rows


def merge_and_dedup(history_rows, current_rows):
    """Mirrors mergeAndDeduplicate() in index.html: keep every current-window
    row as-is, only pull in history rows for orders that don't appear in the
    current window at all."""
    current_ids = {order_id_of(r) for r in current_rows if order_id_of(r)}
    history_only = [r for r in history_rows if order_id_of(r) and order_id_of(r) not in current_ids]
    return history_only + current_rows


def order_level(rows):
    """Mirrors orderLevel(): one row per unique Order No (first occurrence)."""
    seen = {}
    for r in rows:
        oid = order_id_of(r)
        if oid and oid not in seen:
            seen[oid] = r
    return list(seen.values())


def build_open_orders(rows):
    """Mirrors buildSlaRows(), filtered down to just the still-open buckets:
    Open <48H, At Risk, Open >48H."""
    now = now_ist()
    open_orders = []
    skip_cancelled_closed = 0
    skip_no_date = 0
    skip_has_ship_date = 0
    skip_shipped_status_no_date = 0
    for r in rows:
        od = parse_date(get(r, DATE_FIELDS))
        status = norm(get(r, ["Status"])).lower()

        if status in ("cancelled", "closed"):
            skip_cancelled_closed += 1
            continue
        if not od:
            skip_no_date += 1
            continue

        sd = parse_date(get(r, SHIP_DATE_FIELDS))
        is_shipped_status = status in SHIPPED_STATUSES

        if not sd and is_shipped_status:
            sd = parse_date(get(r, MANIFEST_DATE_FIELDS))

        if sd:
            skip_has_ship_date += 1
            continue  # already shipped (Within SLA or Breached) - not "open"
        if is_shipped_status:
            skip_shipped_status_no_date += 1
            continue  # status says shipped but no date at all - treated as shipped, not open

        hours = (now - od).total_seconds() / 3600
        if hours > 48:
            sla = "Open >48H"
        elif hours >= 42:
            sla = "At Risk"
        else:
            sla = "Open <48H"

        open_orders.append({
            "order_no": order_id_of(r),
            "warehouse": norm(get(r, WAREHOUSE_FIELDS)),
            "status": norm(get(r, ["Status"])),
            "order_create_date": od,
            "sla": sla,
            "hours": hours,
        })

    print(
        f"[diagnostic] input orders: {len(rows)} | "
        f"skipped(cancelled/closed): {skip_cancelled_closed} | "
        f"skipped(no order-create-date): {skip_no_date} | "
        f"skipped(has a ship date -> Within/Breached): {skip_has_ship_date} | "
        f"skipped(shipped-status, no date at all): {skip_shipped_status_no_date} | "
        f"-> open orders: {len(open_orders)}"
    )

    open_orders.sort(key=lambda x: x["hours"], reverse=True)
    return open_orders


def build_email_html(open_orders, max_rows=300):
    counts = {"Open <48H": 0, "At Risk": 0, "Open >48H": 0}
    for o in open_orders:
        counts[o["sla"]] += 1
    total = len(open_orders)

    summary_html = "".join(
        f'<td style="padding:10px 16px;border:1px solid #ddd;text-align:center">'
        f'<div style="font-size:11px;color:#666;text-transform:uppercase">{label}</div>'
        f'<div style="font-size:22px;font-weight:700">{count}</div></td>'
        for label, count in [("Open <48H", counts["Open <48H"]), ("At Risk", counts["At Risk"]),
                              ("Open >48H", counts["Open >48H"]), ("Total Open", total)]
    )

    rows_html = ""
    for o in open_orders[:max_rows]:
        color = "#c52f2f" if o["sla"] == "Open >48H" else ("#a46d00" if o["sla"] == "At Risk" else "#555")
        date_str = o["order_create_date"].strftime("%Y-%m-%d %H:%M") if o["order_create_date"] else ""
        rows_html += (
            f'<tr>'
            f'<td style="padding:6px 10px;border:1px solid #eee">{o["order_no"]}</td>'
            f'<td style="padding:6px 10px;border:1px solid #eee">{o["warehouse"]}</td>'
            f'<td style="padding:6px 10px;border:1px solid #eee">{date_str}</td>'
            f'<td style="padding:6px 10px;border:1px solid #eee;color:{color};font-weight:600">{o["sla"]}</td>'
            f'<td style="padding:6px 10px;border:1px solid #eee;text-align:right">{o["hours"]:.1f}</td>'
            f'</tr>'
        )

    truncated_note = ""
    if total > max_rows:
        truncated_note = (
            f'<p style="color:#888;font-size:12px">Showing top {max_rows} of {total} open orders '
            f'(sorted by longest open first). Full list is on the dashboard.</p>'
        )

    now_str = now_ist().strftime("%d %b %Y, %I:%M %p")

    return f"""
    <html><body style="font-family:Arial,sans-serif;color:#182230">
      <h2 style="margin-bottom:4px">Open Orders Summary</h2>
      <p style="color:#666;margin-top:0">As of {now_str}</p>
      <table style="border-collapse:collapse;margin-bottom:20px">
        <tr>{summary_html}</tr>
      </table>
      {truncated_note}
      <table style="border-collapse:collapse;width:100%;font-size:13px">
        <thead>
          <tr style="background:#f5f5f5">
            <th style="padding:8px 10px;border:1px solid #ddd;text-align:left">Order No</th>
            <th style="padding:8px 10px;border:1px solid #ddd;text-align:left">Warehouse</th>
            <th style="padding:8px 10px;border:1px solid #ddd;text-align:left">Order Create Date</th>
            <th style="padding:8px 10px;border:1px solid #ddd;text-align:left">SLA Status</th>
            <th style="padding:8px 10px;border:1px solid #ddd;text-align:right">Open Hours</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
    </body></html>
    """


def send_email(html_body, subject):
    api_key = os.environ["BREVO_API_KEY"]
    sender_email = os.environ["BREVO_SENDER_EMAIL"]
    recipients = [e.strip() for e in os.environ["ALERT_RECIPIENTS"].split(",") if e.strip()]

    if not recipients:
        print("No recipients configured (ALERT_RECIPIENTS is empty) - skipping send.")
        return

    payload = {
        "sender": {"email": sender_email, "name": "PW Dashboard Alerts"},
        "to": [{"email": r} for r in recipients],
        "subject": subject,
        "htmlContent": html_body,
    }

    req = urllib.request.Request(
        "https://api.brevo.com/v3/smtp/email",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "accept": "application/json",
            "api-key": api_key,
            "content-type": "application/json",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(req) as resp:
            print(f"Email sent to: {', '.join(recipients)} (status {resp.status})")
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        print(f"Brevo API error {e.code}: {body}")
        raise


def main():
    current_rows = read_xlsx_best_sheet("data.xlsx") if os.path.exists("data.xlsx") else []
    print(f"[diagnostic] data.xlsx rows read: {len(current_rows)}")
    if current_rows:
        print(f"[diagnostic] sample columns detected: {list(current_rows[0].keys())[:15]}")

    history_rows = []
    for path in sorted(glob.glob("history*.xlsx")):
        this_file_rows = read_xlsx_best_sheet(path)
        print(f"[diagnostic] {path} rows read: {len(this_file_rows)}")
        history_rows.extend(this_file_rows)

    raw = merge_and_dedup(history_rows, current_rows)
    print(f"[diagnostic] merged raw rows: {len(raw)}")
    if not raw:
        print("No data found in data.xlsx/history files - skipping email.")
        sys.exit(0)

    orders = order_level(raw)
    print(f"[diagnostic] unique orders after order_level(): {len(orders)}")
    open_orders = build_open_orders(orders)

    html_body = build_email_html(open_orders)
    total = len(open_orders)
    over48 = sum(1 for o in open_orders if o["sla"] == "Open >48H")
    subject = f"Open Orders Alert - {total} open ({over48} over 48H) - {now_ist().strftime('%d %b %I:%M %p')}"

    send_email(html_body, subject)


if __name__ == "__main__":
    main()

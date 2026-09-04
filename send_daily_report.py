from pathlib import Path
from datetime import datetime
from urllib.request import urlopen, Request
from html import escape

import pandas as pd


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT = BASE_DIR / "data.xlsx"
OUTPUT_HTML = BASE_DIR / "open_orders_email.html"

GITHUB_RAW_BASE = (
    "https://raw.githubusercontent.com/"
    "Kamal2-Pal/PW-Master-Dashboard-V2/main/"
)

HISTORY_FILE_SCAN_LIMIT = 20

SHIPPED_STATUSES = {
    "shipped complete",
    "partially shipped",
    "delivered",
    "shipped & returned",
}


# ============================================================
# GENERIC HELPERS
# ============================================================

def norm(value):
    return str(value).strip()


def norm_col(value):
    return (
        str(value)
        .strip()
        .lower()
        .replace("_", "")
        .replace(" ", "")
    )


def get_value(row, names):
    """
    Same flexible column matching approach used by dashboard logic.
    """
    if row is None:
        return ""

    keys = list(row.keys())

    for name in names:
        # Exact match
        if name in row:
            value = row[name]
            if value not in ("", None):
                return value

        # Normalized match
        target = norm_col(name)

        for key in keys:
            if norm_col(key) == target:
                value = row[key]
                if value not in ("", None):
                    return value

    return ""


def num(value):
    """
    Convert numeric-looking values safely.
    """
    if value is None:
        return 0.0

    try:
        text = str(value).replace(",", "").strip()

        if text == "":
            return 0.0

        result = float(text)

        if pd.isna(result):
            return 0.0

        return result

    except Exception:
        return 0.0


def fmt_num(value):
    try:
        value = float(value)

        if value.is_integer():
            return f"{value:,.0f}"

        return f"{value:,.2f}"

    except Exception:
        return str(value)


def fmt_date(value):
    try:
        parsed = pd.to_datetime(value, errors="coerce")

        if pd.isna(parsed):
            return str(value)

        return parsed.strftime("%d-%b-%Y")

    except Exception:
        return str(value)


def parse_date(value):
    """
    Flexible date parsing matching the dashboard's handling
    of Excel/date/string values.
    """
    if value is None or value == "":
        return None

    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.to_pydatetime()

    if isinstance(value, datetime):
        return value

    if isinstance(value, (int, float)):
        try:
            parsed = pd.to_datetime(
                value,
                unit="D",
                origin="1899-12-30",
                errors="coerce"
            )

            if pd.isna(parsed):
                return None

            return parsed.to_pydatetime()

        except Exception:
            return None

    text = str(value).strip()

    if not text:
        return None

    # YYYY-MM-DD / YYYY/MM/DD
    for fmt in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(text, fmt)

        except ValueError:
            pass

    parsed = pd.to_datetime(text, errors="coerce")

    if pd.isna(parsed):
        return None

    return parsed.to_pydatetime()


# ============================================================
# COLUMN MAPPING
# ============================================================

def find_column(columns, *names):
    mapping = {
        norm_col(column): column
        for column in columns
    }

    for name in names:
        key = norm_col(name)

        if key in mapping:
            return mapping[key]

    return None


# ============================================================
# EXCEL LOADING
# ============================================================

def read_excel_best_sheet(path):
    """
    Dashboard loads the sheet having the largest number of rows.
    We follow the same approach.
    """
    workbook = pd.ExcelFile(path)

    best_df = None
    best_len = -1

    for sheet in workbook.sheet_names:
        temp = pd.read_excel(
            path,
            sheet_name=sheet,
            dtype=object
        )

        if len(temp) > best_len:
            best_df = temp
            best_len = len(temp)

    if best_df is None:
        raise RuntimeError(f"No usable sheet found in {path.name}")

    return best_df


# ============================================================
# HISTORY DOWNLOAD
# ============================================================

def download_history_files():
    """
    Download the same history files used by dashboard:

        history.xlsx
        history-1.xlsx
        history-2.xlsx
        ...

    Stops at the first missing sequential file after history.xlsx.
    """

    downloaded = []

    for index in range(0, HISTORY_FILE_SCAN_LIMIT + 1):

        if index == 0:
            filename = "history.xlsx"
        else:
            filename = f"history-{index}.xlsx"

        local_path = BASE_DIR / filename
        url = GITHUB_RAW_BASE + filename

        try:
            request = Request(
                url,
                headers={
                    "User-Agent": "PW-B2B-Open-Orders-Report"
                },
            )

            with urlopen(request, timeout=60) as response:
                content = response.read()

            if not content:
                raise RuntimeError("Empty file received")

            local_path.write_bytes(content)

            downloaded.append(filename)

            print(f"History downloaded: {filename}")

        except Exception as exc:

            if filename == "history.xlsx":
                print(
                    f"history.xlsx not available: {exc}"
                )

            else:
                # Sequential scan stops at first missing file,
                # same logic as dashboard.
                break

    return downloaded


# ============================================================
# DASHBOARD-STYLE MERGE
# ============================================================

def order_id(row):
    return norm(
        get_value(
            row,
            [
                "Order No",
                "OrderNo",
                "Order Number",
                "Order_No",
                "Order_Number",
                "Order ID",
                "OrderID",
            ],
        )
    )


def merge_history_and_current(history_rows, current_rows):
    """
    Same concept as dashboard mergeAndDeduplicate():

    - Current export is treated as the fresh/live dataset.
    - History rows are added only when that Order No does NOT exist
      in current data.
    - This avoids double counting an order appearing in both places.
    """

    current_order_numbers = {
        order_id(row)
        for row in current_rows
        if order_id(row)
    }

    history_only = []

    for row in history_rows:
        oid = order_id(row)

        if oid and oid not in current_order_numbers:
            history_only.append(row)

    return history_only + current_rows


# ============================================================
# DASHBOARD-STYLE ORDER LEVEL
# ============================================================

def build_order_level(rows):
    """
    Same order-level aggregation used by dashboard.

    Multiple line items with same Order No are aggregated for:

        Order Qty
        Picked Qty
        Shipped Qty
        Open Qty
    """

    result = {}

    for row in rows:

        oid = order_id(row)

        if not oid:
            continue

        if oid not in result:

            result[oid] = {
                "order_no": oid,

                "order_date": get_value(
                    row,
                    [
                        "Order create date",
                        "OrderDate",
                        "Order_Date",
                        "Date",
                    ],
                ),

                "status": get_value(
                    row,
                    ["Status"]
                ),

                "warehouse": get_value(
                    row,
                    [
                        "FulfillmentLocationName",
                        "Warehouse",
                        "Location",
                    ],
                ),

                "finance_category": get_value(
                    row,
                    [
                        "finance_category",
                        "Finance Category",
                        "Category",
                    ],
                ),

                "order_qty": num(
                    get_value(
                        row,
                        [
                            "Order Qty",
                            "Qty",
                        ],
                    )
                ),

                "picked_qty": num(
                    get_value(
                        row,
                        ["Picked Qty"],
                    )
                ),

                "shipped_qty": num(
                    get_value(
                        row,
                        ["Shipped Qty"],
                    )
                ),

                "open_qty": num(
                    get_value(
                        row,
                        ["Open Qty"],
                    )
                ),

                "actual_ship_date": get_value(
                    row,
                    [
                        "Actual_ShipDate",
                        "Actual_ShipDate ",
                        "Ship_Date",
                        "Shipped_Date",
                    ],
                ),

                "manifest_create_date": get_value(
                    row,
                    [
                        "Manifest Create Date",
                        "Manifest_Create_Date",
                        "ManifestCreateDate",
                    ],
                ),

                "wh_remarks": get_value(
                    row,
                    [
                        "WH remarks",
                        "WH Remarks",
                        "WH remarks ",
                        "Remarks",
                    ],
                ),
            }

        else:

            current = result[oid]

            current["order_qty"] += num(
                get_value(
                    row,
                    [
                        "Order Qty",
                        "Qty",
                    ],
                )
            )

            current["picked_qty"] += num(
                get_value(
                    row,
                    ["Picked Qty"],
                )
            )

            current["shipped_qty"] += num(
                get_value(
                    row,
                    ["Shipped Qty"],
                )
            )

            current["open_qty"] += num(
                get_value(
                    row,
                    ["Open Qty"],
                )
            )

    return list(result.values())


# ============================================================
# DASHBOARD-STYLE SLA / OPEN LOGIC
# ============================================================

def classify_orders(order_rows):
    """
    Same business logic as dashboard's buildSlaRows():

    Cancelled / Closed
        -> not open

    Actual_ShipDate available
        -> Within SLA / Breached

    Shipped-like status but no Actual_ShipDate
        -> use Manifest Create Date

    Shipped-like status and no usable ship date
        -> Shipped complete, not open

    Otherwise:
        <42h       -> Open <48H
        42-48h     -> At Risk
        >48h       -> Open >48H
    """

    now = datetime.now()

    open_rows = []

    for row in order_rows:

        status = norm(
            row["status"]
        ).strip().lower()

        order_date = parse_date(
            row["order_date"]
        )

        # Cancelled / Closed
        if status in {
            "cancelled",
            "closed",
        }:
            continue

        # No order date
        if order_date is None:
            continue

        ship_date = parse_date(
            row["actual_ship_date"]
        )

        # Shipped-like status fallback:
        # Actual ship blank -> Manifest Create Date
        if (
            ship_date is None
            and status in SHIPPED_STATUSES
        ):
            ship_date = parse_date(
                row["manifest_create_date"]
            )

        # Actual ship exists -> already shipped, not open
        if ship_date is not None:
            continue

        # Status says shipped but no usable ship date
        if status in SHIPPED_STATUSES:
            continue

        open_hours = (
            now - order_date
        ).total_seconds() / 3600

        if open_hours > 48:
            sla_status = "Open >48H"

        elif open_hours >= 42:
            sla_status = "At Risk"

        else:
            sla_status = "Open <48H"

        row["open_hours"] = open_hours
        row["sla_status"] = sla_status

        open_rows.append(row)

    return open_rows


# ============================================================
# DATA META / LAST SYNC
# ============================================================

def get_data_last_sync():
    """
    Try reading current generatedAt from data-meta.json.
    This is the timestamp written by Vinculum extraction.
    """

    url = GITHUB_RAW_BASE + "data-meta.json"

    try:

        request = Request(
            url,
            headers={
                "User-Agent": "PW-B2B-Open-Orders-Report"
            },
        )

        with urlopen(request, timeout=30) as response:
            raw = response.read().decode(
                "utf-8",
                errors="replace"
            )

        import json

        meta = json.loads(raw)

        generated_at = meta.get(
            "generatedAt"
        )

        if not generated_at:
            return ""

        parsed = pd.to_datetime(
            generated_at,
            errors="coerce"
        )

        if pd.isna(parsed):
            return ""

        return parsed.to_pydatetime().strftime(
            "%d-%b-%Y %I:%M %p"
        )

    except Exception:
        return ""


# ============================================================
# LOAD CURRENT + HISTORY
# ============================================================

if not INPUT.exists():
    raise FileNotFoundError(
        f"Current data file not found: {INPUT}"
    )

print("Loading current data.xlsx...")

current_df = read_excel_best_sheet(
    INPUT
)

current_rows = current_df.to_dict(
    orient="records"
)

print(
    f"Current rows loaded: {len(current_rows):,}"
)


history_rows = []

history_files = download_history_files()

for filename in history_files:

    path = BASE_DIR / filename

    try:

        temp_df = read_excel_best_sheet(
            path
        )

        temp_rows = temp_df.to_dict(
            orient="records"
        )

        history_rows.extend(
            temp_rows
        )

        print(
            f"{filename}: "
            f"{len(temp_rows):,} rows"
        )

    except Exception as exc:

        print(
            f"Could not read {filename}: {exc}"
        )


print(
    f"History rows loaded: "
    f"{len(history_rows):,}"
)


# ============================================================
# MERGE
# ============================================================

merged_rows = merge_history_and_current(
    history_rows,
    current_rows
)

print(
    f"Merged line rows: "
    f"{len(merged_rows):,}"
)


# ============================================================
# ORDER LEVEL
# ============================================================

order_rows = build_order_level(
    merged_rows
)

print(
    f"Unique order numbers: "
    f"{len(order_rows):,}"
)


# ============================================================
# OPEN ORDERS
# ============================================================

open_orders = classify_orders(
    order_rows
)

open_orders.sort(
    key=lambda x: (
        -(x.get("open_hours") or 0),
        str(x.get("order_no", "")),
    )
)


total_open_orders = len(
    open_orders
)

total_open_qty = sum(
    num(row.get("open_qty", 0))
    for row in open_orders
)


open_lt48 = sum(
    1
    for row in open_orders
    if row["sla_status"] == "Open <48H"
)

at_risk = sum(
    1
    for row in open_orders
    if row["sla_status"] == "At Risk"
)

open_gt48 = sum(
    1
    for row in open_orders
    if row["sla_status"] == "Open >48H"
)


report_generated = datetime.now().strftime(
    "%d-%b-%Y %I:%M %p"
)

data_last_sync = get_data_last_sync()


# ============================================================
# HTML TABLE
# ============================================================

table_rows = []

for row in open_orders:

    table_rows.append(
        "<tr>"
        f"<td>{escape(str(row['order_no']))}</td>"
        f"<td>{escape(fmt_date(row['order_date']))}</td>"
        f"<td>{escape(str(row['warehouse'] or ''))}</td>"
        f"<td>{escape(str(row['finance_category'] or ''))}</td>"
        f"<td>{fmt_num(row['order_qty'])}</td>"
        f"<td>{fmt_num(row['picked_qty'])}</td>"
        f"<td>{fmt_num(row['shipped_qty'])}</td>"
        f"<td><b>{fmt_num(row['open_qty'])}</b></td>"
        f"<td>{escape(str(row['status'] or ''))}</td>"
        f"<td>{fmt_num(row['open_hours'])}</td>"
        f"<td><b>{escape(row['sla_status'])}</b></td>"
        "</tr>"
    )


# ============================================================
# HTML REPORT
# ============================================================

data_sync_html = ""

if data_last_sync:
    data_sync_html = (
        f"<b>Data Last Synced:</b> "
        f"{escape(data_last_sync)}<br>"
    )


OUTPUT_HTML.write_text(
    f"""<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>PW B2B - Open Orders Report</title>
</head>

<body
style="
font-family:Arial,sans-serif;
color:#1f2937;
margin:20px;
">

<h2>PW B2B - Open Orders Report</h2>

<p>
<b>Total Open Orders:</b>
{total_open_orders:,}<br>

<b>Total Open Qty:</b>
{fmt_num(total_open_qty)}<br>

<b>Open &lt;48H:</b>
{open_lt48:,}<br>

<b>At Risk (42-48H):</b>
{at_risk:,}<br>

<b>Open &gt;48H:</b>
{open_gt48:,}<br>

{data_sync_html}

<b>Report Generated:</b>
{escape(report_generated)}
</p>


<table
style="
border-collapse:collapse;
width:100%;
font-size:12px;
">

<thead>

<tr>

<th style="border:1px solid #ddd;padding:7px;text-align:left">
Order No
</th>

<th style="border:1px solid #ddd;padding:7px;text-align:left">
Order Date
</th>

<th style="border:1px solid #ddd;padding:7px;text-align:left">
Warehouse
</th>

<th style="border:1px solid #ddd;padding:7px;text-align:left">
Finance Category
</th>

<th style="border:1px solid #ddd;padding:7px">
Order Qty
</th>

<th style="border:1px solid #ddd;padding:7px">
Picked Qty
</th>

<th style="border:1px solid #ddd;padding:7px">
Shipped Qty
</th>

<th style="border:1px solid #ddd;padding:7px">
Open Qty
</th>

<th style="border:1px solid #ddd;padding:7px;text-align:left">
Status
</th>

<th style="border:1px solid #ddd;padding:7px">
Open Hours
</th>

<th style="border:1px solid #ddd;padding:7px;text-align:left">
SLA Status
</th>

</tr>

</thead>

<tbody>

{''.join(table_rows)}

</tbody>

</table>


<p
style="
font-size:11px;
color:#6b7280;
margin-top:12px;
">

Source:
data.xlsx + history.xlsx/history-N.xlsx

<br>

Open Orders logic:
same dashboard 48H SLA logic —
Open &lt;48H + At Risk + Open &gt;48H.

</p>

</body>
</html>
""",
    encoding="utf-8"
)


# ============================================================
# CONSOLE OUTPUT
# ============================================================

print()
print(
    f"Open Orders: {total_open_orders:,}"
)

print(
    f"Open Qty: {fmt_num(total_open_qty)}"
)

print(
    f"Open <48H: {open_lt48:,}"
)

print(
    f"At Risk: {at_risk:,}"
)

print(
    f"Open >48H: {open_gt48:,}"
)

if data_last_sync:
    print(
        f"Data Last Synced: {data_last_sync}"
    )

print(
    f"Report Generated: {report_generated}"
)

print(
    f"Report written to {OUTPUT_HTML}"
)

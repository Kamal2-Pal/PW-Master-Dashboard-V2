from pathlib import Path
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from urllib.request import urlopen, Request
from html import escape
import json

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
    if row is None:
        return ""

    keys = list(row.keys())

    for name in names:

        if name in row:
            value = row[name]

            if value not in ("", None):
                return value

        target = norm_col(name)

        for key in keys:

            if norm_col(key) == target:
                value = row[key]

                if value not in ("", None):
                    return value

    return ""


def num(value):

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

        parsed = pd.to_datetime(
            value,
            errors="coerce"
        )

        if pd.isna(parsed):
            return str(value)

        return parsed.strftime(
            "%d-%b-%Y"
        )

    except Exception:

        return str(value)


def parse_date(value):

    if value is None or value == "":
        return None

    if isinstance(value, pd.Timestamp):

        return (
            None
            if pd.isna(value)
            else value.to_pydatetime()
        )

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

            return datetime.strptime(
                text,
                fmt
            )

        except ValueError:

            pass

    parsed = pd.to_datetime(
        text,
        errors="coerce"
    )

    if pd.isna(parsed):
        return None

    return parsed.to_pydatetime()


# ============================================================
# EXCEL LOADING
# ============================================================

def read_excel_best_sheet(path):

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

        raise RuntimeError(
            f"No usable sheet found in {path.name}"
        )

    return best_df


# ============================================================
# HISTORY DOWNLOAD
# ============================================================

def download_history_files():

    downloaded = []

    for index in range(
        0,
        HISTORY_FILE_SCAN_LIMIT + 1
    ):

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
                    "User-Agent":
                    "PW-B2B-Open-Orders-Report"
                },
            )

            with urlopen(
                request,
                timeout=60
            ) as response:

                content = response.read()

            if not content:

                raise RuntimeError(
                    "Empty file received"
                )

            local_path.write_bytes(
                content
            )

            downloaded.append(
                filename
            )

            print(
                f"History downloaded: {filename}"
            )

        except Exception as exc:

            if filename == "history.xlsx":

                print(
                    f"history.xlsx not available: {exc}"
                )

            else:

                break

    return downloaded


# ============================================================
# ORDER ID
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


# ============================================================
# MERGE CURRENT + HISTORY
# ============================================================

def merge_history_and_current(
    history_rows,
    current_rows
):

    current_order_numbers = {
        order_id(row)
        for row in current_rows
        if order_id(row)
    }

    history_only = []

    for row in history_rows:

        oid = order_id(row)

        if (
            oid
            and oid not in current_order_numbers
        ):

            history_only.append(
                row
            )

    return history_only + current_rows


# ============================================================
# ORDER LEVEL AGGREGATION
# ============================================================

def build_order_level(rows):

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

    return list(
        result.values()
    )


# ============================================================
# SLA / OPEN ORDER CLASSIFICATION
# ============================================================

def classify_orders(order_rows):

    now = datetime.now()

    open_rows = []

    for row in order_rows:

        status = norm(
            row["status"]
        ).strip().lower()

        order_date = parse_date(
            row["order_date"]
        )

        # Cancelled / Closed are not open
        if status in {
            "cancelled",
            "closed",
        }:

            continue

        # Cannot calculate SLA without order date
        if order_date is None:

            continue

        ship_date = parse_date(
            row["actual_ship_date"]
        )

        # For shipped-like statuses,
        # use Manifest Create Date if actual ship date
        # is not available.
        if (
            ship_date is None
            and status in SHIPPED_STATUSES
        ):

            ship_date = parse_date(
                row["manifest_create_date"]
            )

        # Already shipped
        if ship_date is not None:

            continue

        # Shipped status but no usable ship date
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

        open_rows.append(
            row
        )

    return open_rows


# ============================================================
# DATA LAST SYNC
# ============================================================

def get_data_last_sync():

    """
    Read generatedAt from data-meta.json.

    Vinculum/GitHub timestamp is treated as UTC
    and converted to India Standard Time (IST),
    matching the dashboard's displayed time.
    """

    url = (
        GITHUB_RAW_BASE
        + "data-meta.json"
    )

    try:

        request = Request(
            url,
            headers={
                "User-Agent":
                "PW-B2B-Open-Orders-Report"
            },
        )

        with urlopen(
            request,
            timeout=30
        ) as response:

            raw = response.read().decode(
                "utf-8",
                errors="replace"
            )

        meta = json.loads(
            raw
        )

        generated_at = meta.get(
            "generatedAt"
        )

        if not generated_at:
            return ""

        # Parse timestamp
        parsed = datetime.fromisoformat(
            str(generated_at)
            .replace("Z", "+00:00")
        )

        # If source timestamp has no timezone,
        # treat it as UTC.
        if parsed.tzinfo is None:

            parsed = parsed.replace(
                tzinfo=timezone.utc
            )

        # Convert UTC -> IST
        india_time = parsed.astimezone(
            ZoneInfo("Asia/Kolkata")
        )

        return india_time.strftime(
            "%d-%b-%Y %I:%M %p"
        )

    except Exception as exc:

        print(
            f"Could not read data-meta.json: {exc}"
        )

        return ""


# ============================================================
# LOAD CURRENT DATA
# ============================================================

if not INPUT.exists():

    raise FileNotFoundError(
        f"Current data file not found: {INPUT}"
    )

print(
    "Loading current data.xlsx..."
)

current_df = read_excel_best_sheet(
    INPUT
)

current_rows = current_df.to_dict(
    orient="records"
)

print(
    f"Current rows loaded: "
    f"{len(current_rows):,}"
)


# ============================================================
# LOAD HISTORY
# ============================================================

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
            f"Could not read "
            f"{filename}: {exc}"
        )


print(
    f"History rows loaded: "
    f"{len(history_rows):,}"
)


# ============================================================
# MERGE DATA
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
# BUILD ORDER LEVEL
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
        str(
            x.get(
                "order_no",
                ""
            )
        ),
    )
)


total_open_orders = len(
    open_orders
)

total_open_qty = sum(
    num(
        row.get(
            "open_qty",
            0
        )
    )
    for row in open_orders
)


open_lt48 = sum(
    1
    for row in open_orders
    if row["sla_status"]
    == "Open <48H"
)

at_risk = sum(
    1
    for row in open_orders
    if row["sla_status"]
    == "At Risk"
)

open_gt48 = sum(
    1
    for row in open_orders
    if row["sla_status"]
    == "Open >48H"
)


# ============================================================
# TIMESTAMPS
# ============================================================

data_last_sync = get_data_last_sync()

report_generated = datetime.now().strftime(
    "%d-%b-%Y %I:%M %p"
)


# ============================================================
# BUILD HTML TABLE
# ============================================================

table_rows = []

for row in open_orders:

    table_rows.append(

        "<tr>"

        f"<td>"
        f"{escape(str(row['order_no']))}"
        f"</td>"

        f"<td>"
        f"{escape(fmt_date(row['order_date']))}"
        f"</td>"

        f"<td>"
        f"{escape(str(row['warehouse'] or ''))}"
        f"</td>"

        f"<td>"
        f"{escape(str(row['finance_category'] or ''))}"
        f"</td>"

        f"<td>"
        f"{fmt_num(row['order_qty'])}"
        f"</td>"

        f"<td>"
        f"{fmt_num(row['picked_qty'])}"
        f"</td>"

        f"<td>"
        f"{fmt_num(row['shipped_qty'])}"
        f"</td>"

        f"<td>"
        f"<b>{fmt_num(row['open_qty'])}</b>"
        f"</td>"

        f"<td>"
        f"{escape(str(row['status'] or ''))}"
        f"</td>"

        f"<td>"
        f"{fmt_num(row['open_hours'])}"
        f"</td>"

        f"<td>"
        f"<b>{escape(row['sla_status'])}</b>"
        f"</td>"

        "</tr>"
    )


# ============================================================
# DATA LAST SYNC HTML
# ============================================================

data_sync_html = ""

if data_last_sync:

    data_sync_html = (
        "<b>Data Last Synced:</b> "
        f"{escape(data_last_sync)}"
        "<br>"
    )


# ============================================================
# WRITE HTML REPORT
# ============================================================

OUTPUT_HTML.write_text(
    f"""<!doctype html>
<html>

<head>

<meta charset="utf-8">

<title>
PW B2B - Open Orders Report
</title>

</head>

<body
style="
font-family:Arial,sans-serif;
color:#1f2937;
margin:20px;
">

<h2>
PW B2B - Open Orders Report
</h2>

<p>

<b>Total Open Orders:</b>
{total_open_orders:,}
<br>

<b>Total Open Qty:</b>
{fmt_num(total_open_qty)}
<br>

<b>Open &lt;48H:</b>
{open_lt48:,}
<br>

<b>At Risk (42-48H):</b>
{at_risk:,}
<br>

<b>Open &gt;48H:</b>
{open_gt48:,}
<br>

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

<th
style="
border:1px solid #ddd;
padding:7px;
text-align:left;
">
Order No
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
text-align:left;
">
Order Date
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
text-align:left;
">
Warehouse
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
text-align:left;
">
Finance Category
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
">
Order Qty
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
">
Picked Qty
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
">
Shipped Qty
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
">
Open Qty
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
text-align:left;
">
Status
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
">
Open Hours
</th>

<th
style="
border:1px solid #ddd;
padding:7px;
text-align:left;
">
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
Open &lt;48H + At Risk + Open &gt;48H

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
    f"Open Orders: "
    f"{total_open_orders:,}"
)

print(
    f"Open Qty: "
    f"{fmt_num(total_open_qty)}"
)

print(
    f"Open <48H: "
    f"{open_lt48:,}"
)

print(
    f"At Risk: "
    f"{at_risk:,}"
)

print(
    f"Open >48H: "
    f"{open_gt48:,}"
)

if data_last_sync:

    print(
        f"Data Last Synced: "
        f"{data_last_sync}"
    )

print(
    f"Report Generated: "
    f"{report_generated}"
)

print(
    f"Report written to "
    f"{OUTPUT_HTML}"
)

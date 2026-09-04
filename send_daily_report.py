from pathlib import Path
import pandas as pd
from html import escape
from datetime import datetime

INPUT = Path("data.xlsx")
OUTPUT_HTML = Path("open_orders_email.html")

def norm(s):
    return str(s).strip().lower().replace("_","").replace(" ","")

df = pd.read_excel(INPUT, dtype=object)
cols = {norm(c): c for c in df.columns}

def col(*names):
    for name in names:
        if norm(name) in cols:
            return cols[norm(name)]
    return None

ORDER = col("Order No","OrderNo","Order Number")
DATE = col("OrderDate","Order create date","Order_Date")
QTY = col("Order Qty","Qty")
PICKED = col("Picked Qty")
SHIPPED = col("Shipped Qty")
OPEN = col("Open Qty")
STATUS = col("Status")
WAREHOUSE = col("FulfillmentLocationName","Warehouse","Location")

for label, c in [("Order No", ORDER), ("OrderDate", DATE), ("Open Qty", OPEN)]:
    if c is None:
        raise RuntimeError(f"Required column not found: {label}")

open_qty = pd.to_numeric(df[OPEN], errors="coerce").fillna(0)
src = df.loc[open_qty > 0].copy()

orders = {}
for _, r in src.iterrows():
    oid = str(r[ORDER]).strip()
    if not oid or oid.lower() == "nan":
        continue

    if oid not in orders:
        orders[oid] = {
            "date": r[DATE],
            "warehouse": r[WAREHOUSE] if WAREHOUSE else "",
            "order_qty": pd.to_numeric(r[QTY], errors="coerce") if QTY else 0,
            "picked_qty": pd.to_numeric(r[PICKED], errors="coerce") if PICKED else 0,
            "shipped_qty": pd.to_numeric(r[SHIPPED], errors="coerce") if SHIPPED else 0,
            "open_qty": pd.to_numeric(r[OPEN], errors="coerce") if OPEN else 0,
            "status": r[STATUS] if STATUS else "",
        }
    else:
        for key, c in [
            ("order_qty", QTY),
            ("picked_qty", PICKED),
            ("shipped_qty", SHIPPED),
            ("open_qty", OPEN)
        ]:
            if c:
                v = pd.to_numeric(r[c], errors="coerce")
                if pd.notna(v):
                    orders[oid][key] += float(v)

def fmt_date(v):
    try:
        return pd.to_datetime(v).strftime("%d-%b-%Y")
    except Exception:
        return str(v)

def fmt_num(v):
    try:
        f = float(v)
        return f"{f:,.0f}" if f.is_integer() else f"{f:,.2f}"
    except Exception:
        return str(v)

rows = sorted(orders.items(), key=lambda x: str(x[1]["date"]))
total_orders = len(rows)
total_open_qty = sum(float(v["open_qty"] or 0) for _, v in rows)

# Report generation date and time
report_generated = datetime.now().strftime("%d-%b-%Y %I:%M %p")

tr = []
for oid, d in rows:
    tr.append(
        "<tr>"
        f"<td>{escape(oid)}</td>"
        f"<td>{escape(fmt_date(d['date']))}</td>"
        f"<td>{escape(str(d['warehouse'] if pd.notna(d['warehouse']) else ''))}</td>"
        f"<td>{fmt_num(d['order_qty'])}</td>"
        f"<td>{fmt_num(d['picked_qty'])}</td>"
        f"<td>{fmt_num(d['shipped_qty'])}</td>"
        f"<td><b>{fmt_num(d['open_qty'])}</b></td>"
        f"<td>{escape(str(d['status'] if pd.notna(d['status']) else ''))}</td>"
        "</tr>"
    )

OUTPUT_HTML.write_text(f"""<!doctype html>
<html>
<body style="font-family:Arial,sans-serif;color:#1f2937">

<h2>PW B2B - Open Orders Report</h2>

<p>
<b>Total Open Orders:</b> {total_orders}<br>
<b>Total Open Qty:</b> {fmt_num(total_open_qty)}<br>
<b>Report Generated:</b> {report_generated}
</p>

<table style="border-collapse:collapse;width:100%;font-size:13px">
<thead>
<tr>
<th style="border:1px solid #ddd;padding:7px;text-align:left">Order No</th>
<th style="border:1px solid #ddd;padding:7px;text-align:left">Order Date</th>
<th style="border:1px solid #ddd;padding:7px;text-align:left">Warehouse</th>
<th style="border:1px solid #ddd;padding:7px">Order Qty</th>
<th style="border:1px solid #ddd;padding:7px">Picked Qty</th>
<th style="border:1px solid #ddd;padding:7px">Shipped Qty</th>
<th style="border:1px solid #ddd;padding:7px">Open Qty</th>
<th style="border:1px solid #ddd;padding:7px;text-align:left">Status</th>
</tr>
</thead>

<tbody>{''.join(tr)}</tbody>

</table>

<p style="font-size:11px;color:#6b7280">
Source: data.xlsx | Open Orders = Open Qty &gt; 0
</p>

</body>
</html>""", encoding="utf-8")

print(f"Open Orders: {total_orders}")
print(f"Open Qty: {fmt_num(total_open_qty)}")
print(f"Report Generated: {report_generated}")
print(f"Report written to {OUTPUT_HTML}")

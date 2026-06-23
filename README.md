# GSK Pending Order Report

A Streamlit app that generates the daily GSK order management report from marketplace and warehouse files.

## Features

- Processes orders from **Shopee**, **TikTok**, and **Lazada** via the TC Order Report
- Matches orders against **WMS** status (`FULFILLED`, `PENDING_FULFILLMENT`, `PARTIAL_ALLOCATED`, etc.)
- **WMS file can be a single .xlsx or a .zip** of multiple `SalesOrder_Item_Report` parts — combined automatically
- Optional **inventory stock check** for Partial & Non Allocated orders
- Auto-excludes **CANCEL/CANCELLED** WMS orders from Filtered Data and the Pivot
- **Per-SKU remarks** on Partial & Non Allocated orders (see Business Rules below)
- Outputs a formatted **Excel workbook** with 7 sheets:
  - `FILTERED DATA` — active orders, colour-coded by marketplace and WMS status
  - `PIVOT` — order counts by status × MP SLA date (past = red, today = yellow, future = blue)
  - `NOT PUSHED` — orders with no WMS record
  - `PARTIAL & NON ALLOCATED` — one row per SKU with allocation, stock, and remarks
  - `CLOSE & NOT FOUND` — closed or missing orders
  - `ORIGINAL (All Data)` — full TC export
  - `RMA` — cancelled TC orders that are FULFILLED in WMS

## Project Structure

```
gsk_report/
├── app.py              # Streamlit UI
├── report_engine.py    # Core report logic (framework-agnostic)
├── requirements.txt
└── README.md
```

## Local Setup

```bash
# 1. Clone the repo
git clone https://github.com/your-org/gsk-pending-order-report.git
cd gsk-pending-order-report

# 2. Create a virtual environment
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub (public or private).
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select your repo, branch `main`, and set **Main file** to `app.py`.
4. Click **Deploy** — no secrets or environment variables needed.

## Input Files

| File | Format | Required |
|---|---|---|
| `GSK_TC_ORDER_REPORT.csv` | CSV | ✅ |
| `GSK_SHOPEE_ORDER.xlsx` | Excel | ✅ |
| `GSK_WMS.xlsx` **or** `GSK_WMS.zip` | Excel / Zip | ✅ |
| `InventoryWarehouse_Report_*.xlsx` | Excel | ❌ Optional |

## Using the Engine Without Streamlit

The `report_engine.py` module is fully standalone — pass file-like objects or paths:

```python
from report_engine import generate_report

with open("GSK_TC_ORDER_REPORT.csv", "rb") as tc, \
     open("GSK_SHOPEE_ORDER.xlsx", "rb") as shopee, \
     open("GSK_WMS.xlsx", "rb") as wms, \
     open("InventoryWarehouse_Report.xlsx", "rb") as inv:

    excel_bytes, summary = generate_report(tc, shopee, wms, inv)

with open("GSK_TC_ORDER_REPORT_PROCESSED.xlsx", "wb") as f:
    f.write(excel_bytes)

print(summary)
# {'filtered': 443, 'not_pushed': 0, 'partial': 159, 'close_nf': 0, 'rma': 8}
```

`wms` can also point to a `.zip` file — it's detected by extension and the
contained `.xlsx` parts are concatenated automatically.

## Business Rules

- **Filtered statuses**: `ACCEPTED/PICKED`, `NEW`, `READY TO SHIP`
- **Marketplaces**: Shopee, TikTok, Lazada (matched via `nickname` column)
- **Payment**: `COMPLETED` or `PENDING` with COD
- **MP SLA**: Shopee → `Estimated Ship Out Date`; TikTok/Lazada → `ordered_date + 1 day`
- **WMS matching**: strips `GSK_` prefix from `client_order_id` for reliable matching
- **Auto-excluded** from Filtered Data & Pivot: `CANCEL`/`CANCELLED` WMS orders,
  `NOT FOUND` orders, zero-stock partial orders

### Partial & Non Allocated — Remarks Logic

The TC report has one row per SKU for multi-item orders, so order-summary
columns are deduplicated to one row per invoice, while WMS line items are
expanded so every SKU still gets its own row with its own remark:

| Scenario | Remark |
|---|---|
| Single SKU in the order, zero quantity allocated | `CANCEL` |
| Multiple SKUs, **this** SKU has zero quantity allocated | `REQUEST REFUND` |
| Multiple SKUs, **this** SKU is fully allocated | `FULFILL` |
| Same SKU split across multiple lines, total allocated < total ordered | `REQUEST PARTIAL REFUND` |

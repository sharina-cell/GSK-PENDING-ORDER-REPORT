# GSK Pending Order Report — Streamlit App

Daily report generator + NOT PUSHED order updater for GSK/Haleon marketplace orders.

---

## Features

- **Generate Report** — upload TC, Shopee, WMS, and Inventory files to produce the full 7-sheet Excel workbook
- **Update NOT PUSHED Orders** — paste invoice numbers confirmed in WMS; the app moves them to FILTERED DATA and rebuilds the PIVOT automatically

---

## Local Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

---

## Deploy to Streamlit Community Cloud

1. Push this folder to a GitHub repository (public or private)
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in
3. Click **New app**
4. Select your repo, branch (`main`), and set **Main file path** to `app.py`
5. Click **Deploy**

---

## File Structure

```
gsk-report-updater/
├── app.py              # Main Streamlit app
├── requirements.txt    # Python dependencies
└── README.md
```

---

## How to Use

### Tab 1 — Generate Report
| File | Description |
|---|---|
| TC Order Report (.csv) | Main TC export (`GSK_TC_ORDER_REPORT.csv`) |
| Shopee Order File (.xlsx) | Shopee estimated ship-out dates |
| WMS File(s) (.xlsx) | One or multiple WMS files (auto-combined) |
| Inventory Report (.xlsx) | InventoryWarehouse or InventoryZone format |

Click **Generate Report** → download the Excel file.

### Tab 2 — Update NOT PUSHED Orders
1. The report from Tab 1 is automatically loaded, **or** upload an existing report
2. Paste the invoice numbers (one per line or comma-separated) confirmed as pushed in WMS
3. Select the WMS status (`PENDING_FULFILLMENT`, `FULFILLED`, etc.)
4. Click **Update Report** → download the updated file

The app will:
- Move matching rows from **NOT PUSHED** → **FILTERED DATA**
- Remove them from **CLOSE & NOT FOUND**
- Rebuild the **PIVOT** table

---

## Report Sheet Reference

| Sheet | Contents |
|---|---|
| FILTERED DATA | Active pending orders (excl. NOT FOUND, CANCEL, 0-stock partial) |
| PIVOT | Order count by MP SLA date × Order Item Status |
| NOT PUSHED | Orders not yet found in WMS |
| PARTIAL & NON ALLOCATED | Partial/none allocated + stock levels + remarks |
| CLOSE & NOT FOUND | CLOSE or NOT FOUND WMS status |
| ORIGINAL (All Data) | Full TC report, unfiltered |
| RMA | Cancelled orders already fulfilled in WMS |

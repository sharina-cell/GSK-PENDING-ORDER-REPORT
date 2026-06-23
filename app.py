"""
GSK Pending Order Report — Streamlit App
Upload your daily files and download the processed Excel report.
"""

from datetime import datetime

import streamlit as st

from report_engine import generate_report

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GSK Pending Order Report",
    page_icon="📦",
    layout="centered",
)

st.title("📦 GSK Pending Order Report")
st.caption("Upload your daily files below to generate the processed Excel report.")

# ── File uploaders ────────────────────────────────────────────────────────────
st.subheader("Upload Files")

col1, col2 = st.columns(2)
with col1:
    tc_file     = st.file_uploader("TC Order Report (.csv)", type=["csv"], key="tc")
    shopee_file = st.file_uploader("Shopee Orders (.xlsx)", type=["xlsx"], key="shopee")
with col2:
    wms_file = st.file_uploader(
        "WMS File (.xlsx or .zip)",
        type=["xlsx", "zip"],
        key="wms",
        help="If WMS arrives as a .zip with multiple SalesOrder_Item_Report parts, "
             "they'll be combined automatically.",
    )
    inv_file = st.file_uploader("Inventory Warehouse Report (.xlsx) — optional", type=["xlsx"], key="inv")

if inv_file is None:
    st.info("ℹ️ No inventory file uploaded — WH Stock will be skipped for Partial & Non Allocated orders.")

st.divider()

# ── Generate button ───────────────────────────────────────────────────────────
ready = all([tc_file, shopee_file, wms_file])

if not ready:
    st.warning("Please upload the TC Order Report, Shopee Orders, and WMS file to continue.")

if st.button("🚀 Generate Report", disabled=not ready, use_container_width=True, type="primary"):
    progress_bar = st.progress(0, text="Starting…")
    steps = [
        "Loading TC Order Report…",
        "Loading Shopee orders…",
        "Loading WMS…",
        "Loading Inventory…",
        "Filtering orders…",
        "Building pivot table…",
        "Writing Excel workbook…",
        "Done!",
    ]
    step_pct = {s: int((i + 1) / len(steps) * 100) for i, s in enumerate(steps)}

    def update_progress(msg: str):
        progress_bar.progress(step_pct.get(msg, 100), text=msg)

    try:
        excel_bytes, summary = generate_report(
            tc_file=tc_file,
            shopee_file=shopee_file,
            wms_file=wms_file,
            inv_file=inv_file,
            progress_callback=update_progress,
        )

        progress_bar.progress(100, text="✅ Report generated!")

        # ── Summary ───────────────────────────────────────────────────────────
        st.subheader("📊 Summary")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("✅ Filtered Data",       summary["filtered"])
        c2.metric("🔴 Not Pushed",          summary["not_pushed"])
        c3.metric("🟠 Partial / Non Alloc", summary["partial"])
        c4.metric("⚫ Close & Not Found",    summary["close_nf"])
        c5.metric("🔴 RMA",                 summary["rma"])

        # ── Download ──────────────────────────────────────────────────────────
        filename = f"GSK_TC_ORDER_REPORT_PROCESSED_{datetime.today().strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇️ Download Excel Report",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        progress_bar.empty()
        st.error(f"❌ Error generating report: {e}")
        st.exception(e)

# ── Business rules reference ──────────────────────────────────────────────────
with st.expander("ℹ️ Business rules used in this report"):
    st.markdown(
        """
- **Filtered statuses**: `ACCEPTED/PICKED`, `NEW`, `READY TO SHIP`
- **Marketplaces**: Shopee, TikTok, Lazada
- **Payment**: `COMPLETED`, or `PENDING` with COD
- **MP SLA**: Shopee → Estimated Ship Out Date; TikTok/Lazada → order date + 1 day
- **Auto-excluded** from Filtered Data & Pivot: `CANCEL`/`CANCELLED` WMS orders,
  `NOT FOUND` orders, zero-stock partial orders
- **Partial & Non Allocated remarks**:
    - Single SKU, zero allocated → **CANCEL**
    - Multiple SKUs, this SKU zero allocated → **REQUEST REFUND**
    - Multiple SKUs, this SKU fully allocated → **FULFILL**
    - Same SKU under-allocated across lines → **REQUEST PARTIAL REFUND**
        """
    )

st.divider()
st.caption("GSK Order Management Report • " + str(datetime.today().year))

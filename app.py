"""
SellUp Stock Sync - Streamlit front-end
=======================================
A thin UI over sellup_core.py. Separate codebase from the Shopee automation.

Workflow each run:
  1. Upload the Masterlist (stock_report*.xlsx) and the SellUp bulk export
     (INVENTORIES_*.xlsx). Optionally upload a previously-reviewed registry to
     carry forward your locked/reviewed decisions.
  2. The app matches SellUp rows to the Masterlist, routes quantities into the
     right condition column (New NA / New A / Used->Excellent), merges region
     variants, and respects the price-gate (only writes Qty where Price > 0).
  3. Review the summary and any "Match Review" / "Not on SellUp" rows.
  4. Tick the review gate, then download:
        - the updated Match Review registry
        - the ready-to-upload SellUp file (only Seller Qty rewritten)
"""

import io
import os
import base64
import pandas as pd
import streamlit as st

import sellup_core as sc

_APP_DIR = os.path.dirname(os.path.abspath(__file__))


def _logo_data_uri():
    """Return a data URI for the MM logo if mm_logo.png is present, else None."""
    for name in ("mm_logo.png", "mm-logo.png", "mister_mobile.png"):
        p = os.path.join(_APP_DIR, name)
        if os.path.exists(p):
            with open(p, "rb") as f:
                return "data:image/png;base64," + base64.b64encode(f.read()).decode()
    return None

st.set_page_config(page_title="Mister Mobile · SellUp Stock Sync",
                   page_icon="📦", layout="wide")

# ---- Mister Mobile brand styling ----
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@600;700;800&family=Open+Sans:wght@400;600&display=swap');
    html, body, [class*="css"], .stMarkdown, p, li, label, div { font-family: 'Open Sans', sans-serif; }
    h1, h2, h3, h4 { font-family: 'Montserrat', sans-serif !important; font-weight: 800 !important; color: #000; }
    /* MM header banner */
    .mm-banner { background: #FFEB00; border-radius: 16px; padding: 18px 24px; display: flex;
                 align-items: center; gap: 20px; margin-bottom: 8px; }
    .mm-logo-chip { background: #fff; border-radius: 12px; padding: 8px 12px; display: flex; align-items: center; }
    .mm-logo-chip img { height: 56px; display: block; }
    .mm-mark { width: 54px; height: 54px; background: #000; border-radius: 50%; display: flex;
               align-items: center; justify-content: center; color: #FFEB00; font-family: 'Montserrat';
               font-weight: 800; font-size: 22px; flex: none; }
    .mm-title { font-family: 'Montserrat'; font-weight: 800; font-size: 26px; color: #000; line-height: 1.1; }
    .mm-tag { font-family: 'Open Sans'; font-weight: 600; font-size: 14px; color: #000; opacity: 0.75; }
    .mm-strip { height: 6px; background: #000; border-radius: 3px; margin: 0 0 22px; }
    /* buttons yellow/black */
    .stButton > button, .stDownloadButton > button {
        background: #FFEB00 !important; color: #000 !important; border: 2px solid #000 !important;
        border-radius: 10px !important; font-family: 'Montserrat' !important; font-weight: 700 !important; }
    .stButton > button:hover, .stDownloadButton > button:hover { background: #000 !important; color: #FFEB00 !important; }
    .stButton > button:disabled, .stDownloadButton > button:disabled {
        background: #EDEDED !important; color: #6D6962 !important; border-color: #D6D6D6 !important; }
    [data-testid="stMetricValue"] { font-family: 'Montserrat'; color: #000; }
    </style>
    """,
    unsafe_allow_html=True,
)
_logo = _logo_data_uri()
_logo_html = (f'<div class="mm-logo-chip"><img src="{_logo}" alt="Mister Mobile"/></div>'
              if _logo else '<div class="mm-mark">MM</div>')
st.markdown(
    f"""
    <div class="mm-banner">
      {_logo_html}
      <div>
        <div class="mm-title">SellUp Stock Sync</div>
        <div class="mm-tag">Mister Mobile · Dealer Inventory Bulk Update</div>
      </div>
    </div>
    <div class="mm-strip"></div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    "#### How to use\n"
    "1. **Upload** the Masterlist and the SellUp export below.\n"
    "2. **Download the Match Review registry**, open the **New Masterlist SKUs** tab, "
    "and set a **Reviewer Decision** for *every* row (Linked + SellUp SKU ID, or "
    "Not on SellUp Yet / Not Selling / Skipped).\n"
    "3. **Re-upload** the completed registry (box 3). The ready-to-upload SellUp file "
    "stays locked until **all New Masterlist SKUs are reviewed & matched**.\n"
    "4. **Download** the SellUp file and upload it in SellUp → Dealer Inventory → Bulk Update."
)

with st.expander("How matching works (SellUp rules)", expanded=False):
    st.markdown(
        "- **Key:** `Brand | Model | Storage | Colour`, normalised on both sides.\n"
        "- **Condition routing:** Masterlist `New`+`NA` → **New (Not Activated)**, "
        "`New`+`A` → **New (Activated)**, `Used` → **Excellent**.\n"
        "- **Region variants** (HK/US/AUS/…) are merged into the base model.\n"
        "- **Price-gate:** a condition's Qty is written only if that condition's "
        "Price cell is populated and > 0 (SellUp skips empty-price grades).\n"
        "- Only **Exact** and **High**-confidence matches are written; **Review** "
        "rows wait for your confirmation."
    )

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("1️⃣ Masterlist (stock_report*.xlsx)", type=["xlsx"])
with col2:
    sellup_file = st.file_uploader("2️⃣ SellUp export (INVENTORIES_*.xlsx)", type=["xlsx"])

registry_file = st.file_uploader(
    "3️⃣ Reviewed registry (re-upload your completed SellUp_Match_Review.xlsx to "
    "carry your confirmed links & decisions forward)",
    type=["xlsx"],
)


def _apply_prior_decisions(results, registry_bytes):
    """Promote/demote matches based on a reviewed registry's Decision column."""
    try:
        rev = pd.read_excel(io.BytesIO(registry_bytes), sheet_name="Match Review")
    except Exception:
        return results, 0
    locked = {}
    for _, row in rev.iterrows():
        dec = str(row.get("Reviewer Decision", row.get("Decision", ""))).strip().upper()
        key = (str(row.get("SellUp SKU ID", "")), str(row.get("Condition Column", "")))
        if "LOCKED" in dec or dec in ("LINK", "LINKED", "YES"):
            locked[key] = "lock"
        elif "SKIP" in dec or dec in ("NO", "IGNORE"):
            locked[key] = "skip"
    promoted = 0
    for r in results:
        k = (r.sellup.sku_id, r.condition)
        if k in locked:
            if locked[k] == "lock" and r.confidence == "Review":
                r.confidence = "High"
                promoted += 1
            elif locked[k] == "skip":
                r.confidence = "None"
    return results, promoted


if master_file and sellup_file:
    with st.spinner("Matching…"):
        master = sc.load_masterlist(master_file)
        rows_by_sheet, wb_sell = sc.load_sellup(sellup_file)
        # Confirmed links & decisions are carried forward from the reviewed
        # registry (its Locked Matches + New Masterlist SKU links).
        crosswalk = {}
        prior = {}
        promoted = 0
        if registry_file is not None:
            reg_bytes = registry_file.getvalue()
            crosswalk = sc.read_registry_links(io.BytesIO(reg_bytes))
            prior = sc.read_prior_decisions(io.BytesIO(reg_bytes))
        results, unmatched = sc.match(master, rows_by_sheet, crosswalk=crosswalk)
        if registry_file is not None:
            results, promoted = _apply_prior_decisions(results, reg_bytes)
        # freebies + carried-forward New Masterlist SKU decisions (write path)
        results_final, buckets = sc.reconcile_decisions(
            results, unmatched, rows_by_sheet, prior)

    stats = sc.summarise(results_final)
    locked_n = stats.get("Exact", 0) + stats.get("High", 0)

    st.subheader("Summary")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Locked (Exact+High)", locked_n)
    m2.metric("New Masterlist SKUs (to review)", len(buckets["new"]))
    m3.metric("Match Review (SellUp, no MM match)",
              stats.get("None", 0) + stats.get("Review", 0))
    m4.metric("Not Selling (incl. freebies)", len(buckets["not_selling"]))

    # ---- tables ----
    def results_df(conf):
        rows = [{
            "Sheet": r.sellup.sheet, "SKU ID": r.sellup.sku_id,
            "Model": r.sellup.raw_model or r.sellup.model_base,
            "Storage": r.sellup.storage,
            "Colour": r.sellup.colour, "Condition": r.condition,
            "Qty": r.qty, "Confidence": r.confidence, "Score": round(r.score, 1),
            "Masterlist IDs": ", ".join(m.stock_id for m in r.master_rows),
            "Note": r.note,
        } for r in results_final if r.confidence == conf]
        return pd.DataFrame(rows)

    tab1, tab2, tab3 = st.tabs(
        ["✅ Locked", "🆕 New Masterlist SKUs", "🔎 Match Review"])
    with tab1:
        df = pd.concat([results_df("Exact"), results_df("High")], ignore_index=True)
        st.dataframe(df, use_container_width=True, height=350)
    with tab2:
        un = pd.DataFrame([{
            "Masterlist ID": m.stock_id, "Category": m.category, "Brand": m.brand,
            "Model (raw)": m.raw_model, "Colour": m.raw_colour or m.colour,
            "Routed to": m.condition, "Qty": m.qty, "Note": note,
        } for m, note in buckets["new"]])
        st.dataframe(un, use_container_width=True, height=350)
        st.info("Masterlist items (MM stock) not yet matched to a SellUp listing. "
                "Set **Reviewer Decision** in the registry (Linked / Not on SellUp "
                "Yet / Not Selling in SellUp) and re-upload to carry it forward.")
    with tab3:
        mr = pd.concat([results_df("None"), results_df("Review")],
                       ignore_index=True)
        st.dataframe(mr, use_container_width=True, height=350)
        st.info("SellUp SKUs that did not match the Masterlist — no MM stock "
                "record, or only a weak fuzzy suggestion to confirm.")

    st.divider()

    # ---- Step: registry download (always available) ----
    st.markdown("### Step 1 — Review the Match Review registry")
    reg_wb = sc.build_review_workbook(results, unmatched, master, rows_by_sheet,
                                      manual_decisions=prior)
    reg_buf = io.BytesIO(); reg_wb.save(reg_buf)
    st.download_button(
        "⬇️ Download Match Review registry",
        reg_buf.getvalue(), file_name="SellUp_Match_Review.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption("Open the **New Masterlist SKUs** tab and set a Reviewer Decision for "
               "every row, then re-upload it in box 3 above.")

    # ---- Step: verification gate ----
    st.markdown("### Step 2 — Verification")
    pending = len(buckets["new"])
    if pending == 0:
        st.success("✅ All New Masterlist SKUs are reviewed & matched — you can "
                   "download the SellUp file below.")
    else:
        st.warning(
            f"⚠ {pending} New Masterlist SKU(s) still need a Reviewer Decision. "
            "Complete them in the registry's **New Masterlist SKUs** tab "
            "(Linked / Not on SellUp Yet / Not Selling / Skipped) and re-upload it, "
            "then this unlocks."
        )

    # ---- Step: SellUp file download (gated on verification) ----
    st.markdown("### Step 3 — Download the SellUp bulk update")
    if pending == 0:
        written = sc.apply_quantities(wb_sell, results_final,
                                      only_confidences=("Exact", "High"))
        out_buf = io.BytesIO(); wb_sell.save(out_buf)
        st.success(f"Wrote {written} Seller-Qty cells "
                   f"({sum(r.qty for r in results_final if r.confidence in ('Exact','High'))} units).")
        st.download_button(
            "⬇️ Download ready-to-upload SellUp file",
            out_buf.getvalue(), file_name="INVENTORIES_updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption("Upload this in SellUp → Dealer Inventory → Bulk Update. "
                   "Only Seller Qty is written; prices are never changed.")
    else:
        st.button("⬇️ Download unlocks after all New Masterlist SKUs are reviewed",
                  disabled=True)
else:
    st.info("Upload the Masterlist and the SellUp export to begin.")

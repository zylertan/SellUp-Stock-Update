# SellUp Stock Sync — Streamlit app (deploys like your Shopee app)

This is the SellUp equivalent of your Shopee stock-update app. Same setup:
a Streamlit app hosted **free on Streamlit Community Cloud**.

Styled in the **Mister Mobile** brand (yellow/black, Montserrat + Open Sans).

## Files (put these in a GitHub repo)
- `app.py` — the Streamlit interface  **← set this as the Main file path when deploying**
- `sellup_core.py` — all the matching logic
- `requirements.txt` — dependencies
- `.streamlit/config.toml` — Mister Mobile theme colours

> ⚠️ When deploying, the **Main file path must be `app.py`** (not `sellup_core.py`,
> which has no interface and shows a blank page).

## How it runs (same flow as Shopee)
1. **Upload** the Masterlist (`stock_report*.xlsx`) and the SellUp export
   (`INVENTORIES_*.xlsx`). Re-upload your reviewed registry to carry confirmed
   links & decisions forward (no separate crosswalk file needed).
2. **Download the Match Review registry**, open the **New Masterlist SKUs** tab and
   set a **Reviewer Decision** for every row (Linked + SellUp SKU ID, or Not on
   SellUp Yet / Not Selling / Skipped).
3. **Re-upload** the completed registry. A **verification gate** keeps the SellUp
   bulk file locked until *all* New Masterlist SKUs are reviewed & matched.
4. **Download** `INVENTORIES_updated.xlsx` (only Seller Qty written, prices never
   touched) and upload it in SellUp → Dealer Inventory → Bulk Update.

## Deploy to Streamlit Community Cloud
1. Create a new GitHub repo (e.g. `sellup-stock-update`) and add the three files
   (`app.py`, `sellup_core.py`, `requirements.txt`).
2. Go to https://share.streamlit.io → **Create app** → **Deploy a public app from GitHub**.
3. Pick your repo, branch `main`, and set **Main file path** = `app.py`.
4. Click **Deploy**. You'll get a public URL just like your Shopee app.

Any time you push changes to the repo, Streamlit redeploys automatically.

## Run locally (optional)
```bash
pip install -r requirements.txt
streamlit run app.py
```

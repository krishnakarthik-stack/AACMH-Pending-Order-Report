# AACMH Kose Pending Order Report

A Streamlit app that takes the daily orders CSV and the sale_order Excel
export, and produces the filtered, WMS-matched "AACMH Kose Pending Order
Report" — including the color-coded pivot Summary sheet — ready to download.

## What it does

1. Converts the orders CSV to Excel.
2. Filters by:
   - **Order Item Status (J):** ACCEPTED/PICKED, NEW, READY TO SHIP
   - **Payment Status (G):** COMPLETED, PENDING
   - **Channel (BC):** lazada, shopee
   - **Nickname (BD):** the 4 Kose storefronts (Lazada/Shopee x Kose/Decorté)
3. Extracts columns B, D, G, H, I, J, L, AS, BC, BD.
4. Adds **WMS Status** and **WMS Payment**, matched from the sale_order file
   by Channel Order ID → Order Number.
5. Strips the time off Ordered Date, keeping only the date.
6. Builds a pivot-style **Summary** sheet (Nickname → Status → WMS, dates as
   columns, color-coded counts, Grand Total row).

## Files

- `app.py` — the Streamlit app
- `requirements.txt` — Python dependencies

## Deploy to GitHub + Streamlit Cloud

### 1. Push to GitHub
```bash
git init
git add app.py requirements.txt README.md
git commit -m "Initial commit: Kose pending order report app"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo>.git
git push -u origin main
```

### 2. Deploy on Streamlit Community Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **New app**.
3. Pick your repo, branch (`main`), and set **Main file path** to `app.py`.
4. Click **Deploy**.

Streamlit Cloud installs everything from `requirements.txt` automatically.
Once deployed you'll get a shareable URL — anyone with the link can upload
their two files and download the finished report, no local setup needed.

## Run locally instead

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL it prints (usually `http://localhost:8501`).

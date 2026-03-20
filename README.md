# GST Audit Reconciliation Tool

**Created by: Purva Doshi**

Free web app for GST audit reconciliation. No installation needed — runs in any browser.

## Features
- **ITC Reco** — GSTR-2B vs Tally Purchase Register
- **Sales Reco** — GSTR-1 vs Books (Sales Register)
- Auto-detects Tally, SAP, Excel formats
- Colour-coded Excel output with audit remarks

## Deploy on Streamlit Cloud (Free)
1. Fork this repo on GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect GitHub → select repo → `app.py` → Deploy

## File Formats
**ITC Reco:** GSTR-2B Excel (portal) + Detailed Purchase Register (Tally)
**Sales Reco:** E-Invoice Excel (portal, sheet: `b2b, sez, de`) + Sales Register (Tally)

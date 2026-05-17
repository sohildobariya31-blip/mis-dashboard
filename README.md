# Purchase MIS Dashboard

**Live URL:** https://sohildobariya31-blip.github.io/mis-dashboard/

A self-contained HTML dashboard for purchase records (BOM + Other material sheets) with FX conversion, multi-currency support, smart filters, and Google Sheets live integration.

## Quick share

Just send this link to anyone — they can use it directly in any modern browser:

> **https://sohildobariya31-blip.github.io/mis-dashboard/**

No install, no login. Anyone with the link can:
- Upload Excel files (`.xlsx`) and see the dashboard locally in their browser
- Or click "Load from Google Sheets" to fetch the live published sheets
- Filter by sheet (BOM only, Other material, 25-26, 26-27, etc.)
- Switch themes, export Excel/CSV/PDF/PPT, drag rows and columns to reorder

## What it does

- **5 animated KPI cards** (PO total, POs, vendors, outstanding, items × qty)
- **Multi-currency**: INR, USD, RMB, CNY, EUR, GBP — all converted to ₹ INR using monthly exchange rates (pre-filled from web averages)
- **Sheet picker** — when fetching from Google Sheets, choose which sheets to load
- **Smart filters** — categories, suppliers, items, variants, status, COO, currency, project — all dynamic
- **Drag-and-drop** rows and columns to reorder
- **Editable table** — Excel-style formulas (Amount = Qty × Rate, Total = (Amount + Freight) × (1 + GST), Payable = Total − Debit, Outstanding = Payable − Paid)
- **Special calculations** for Solar Cell (per Wp), Al Frame (per set), Glass (sqm ↔ pieces)
- **Partial deliveries** — split a PO row into received + remaining with status (Received / In Transit / On Hold / Tentative)
- **Export** — Excel with formulas, CSV, PDF (full table + charts), PPT (each chart on a slide), or any chart as JPG
- **7 themes** — Midnight (default), Ocean, Sunset, Forest, Cyber, Mono Dark, Light

## Files in this repo

| File | What it is |
|---|---|
| `index.html` / `dashboard.html` | The dashboard itself (single file, all-in-one) |
| `Purchase-MIS-Dashboard.html` | Older snapshot of the dashboard |
| `upload.html`, `test_upload.html` | Standalone upload helpers |
| `AppsScript-ReadyToPaste.js` | Google Apps Script for serving Sheet data via URL |
| `DailyBackup-GoogleSheets.js` | Daily 1 AM IST backup of the Google Sheet (keeps last 30 days) |
| `GoogleSheet-Setup-BOM.js` | Apps Script to set up the BOM sheet structure |
| `Google-Apps-Script-Setup.md` | Step-by-step Apps Script setup guide |
| `HOSTING-GUIDE.md` | How to host the dashboard yourself |
| `Instructions-Gujarati.md` | User instructions in Gujarati |
| `build_dashboard.py`, `generate_template.py` | Optional Python helpers for building / generating templates |
| `data.js` | Sample baked-in data |

## Local use

1. Download or clone the repo
2. Open `dashboard.html` in any modern browser
3. Upload your Excel file or click "Load from Google Sheets"

That's it — everything runs locally in your browser, no server needed.

## Privacy

Excel files are **not** committed to this repo. Your data stays in your browser unless you choose to publish your Google Sheets.

## License

Private use. Built for internal solar manufacturing purchase tracking.

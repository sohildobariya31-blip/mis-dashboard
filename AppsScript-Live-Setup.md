# Live Apps Script Setup (instant data refresh)

This makes the dashboard pull from your Google Sheets in **~3-8 seconds** instead of waiting **5-15 minutes** for the published-CSV cache.

## One-time setup (~5 min)

### 1. Find your spreadsheet IDs

Open each Google Sheet in your browser. The URL looks like:

```
https://docs.google.com/spreadsheets/d/1aBcDeFgHiJ_xxxxxxxxxxxxxxxxxxxxxx/edit#gid=...
                                       └────────── this is the ID ──────────┘
```

Copy the ID for the **26-27** file and the **25-26** file.

### 2. Paste the script

1. Open EITHER source spreadsheet (e.g. the 26-27 one).
2. Top menu → **Extensions → Apps Script**.
3. Delete whatever's in `Code.gs`.
4. Paste the entire contents of `AppsScript-LiveJSON.js` from this repo.
5. Edit the `SOURCES` config at the top — replace the two `PASTE_…` placeholders with your spreadsheet IDs.
6. Click the **disk icon** to save. Give the project a name like "Purchase MIS Live".

### 3. Test it (optional but recommended)

In the Apps Script editor:
1. Pick `testFetch` from the function dropdown.
2. Click ▶️ Run.
3. The first run will ask for permission — click "Review permissions" → choose your Google account → "Advanced" → "Go to (project name) (unsafe)" → "Allow".
   - "Unsafe" just means it's a personal script that hasn't been verified by Google. It's reading your own sheet — that's fine.
4. View → Logs (or Ctrl+Enter). You should see something like:
   ```
   OK: true
   Elapsed: 1240 ms
     BOM 26-27 → 1531 rows × 44 cols
     Other material 26-27 → 1962 rows × 29 cols
     BOM 25-26 → 2950 rows × 42 cols
     Other material 25-26 → 7066 rows × 37 cols
   ```

If you get an "openById" error, the spreadsheet ID is wrong — double-check step 1.

### 4. Deploy as a Web App

1. Top-right → **Deploy** → **New deployment**.
2. Click the gear ⚙️ next to "Select type" → choose **Web app**.
3. Fill in:
   - **Description**: `Purchase MIS Live` (or anything)
   - **Execute as**: **Me** (your account — this lets the script read the sheets you own)
   - **Who has access**: **Anyone** (read-only — no edits possible from outside)
4. Click **Deploy**.
5. **Copy the Web app URL**. It looks like:
   ```
   https://script.google.com/macros/s/AKfycbzXXXXXXXXXXXXXXXXXXXXXX/exec
   ```

### 5. Connect the dashboard

1. Open the dashboard.
2. Click **Data ▾** (top-right) → **🔗 Apps Script live endpoint**.
3. Paste the Web app URL.
4. Click **Save & Fetch**.

You'll see new data load in seconds. After that, just hit the **🔄 Refresh** button (or re-open the menu) any time to pull fresh data.

## Updating the script later

If you tweak `AppsScript-LiveJSON.js`:
- Editor → **Deploy → Manage deployments → ✏️ pencil → Version: New version → Deploy**.
- The Web app URL stays the same — no need to update the dashboard.

## Security

- The Web app URL is **read-only**. Nobody can edit your sheet through it.
- Anyone with the URL **can read** your sheet data. Don't post it publicly.
- To revoke access: Apps Script → Deploy → Manage deployments → Archive.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `cannot open: …` in `testFetch` log | Wrong spreadsheet ID, or the script's Google account doesn't own/share that sheet. Open the sheet, click Share, add the script's owner email as Editor. |
| `sheet not found: BOM` | Sheet tab name doesn't match exactly. Edit `SOURCES` in the script and put the actual tab name. |
| Web app URL returns "Authorization required" page | "Who has access" wasn't set to "Anyone". Re-deploy with the right setting. |
| Dashboard says "fetch failed" | Open the URL in a browser tab. If you see HTML, deployment is wrong. Should return JSON. |
| Slow (>30s) | The sheets are huge. Apps Script is paginating internally — it'll still work, just slower than CSV. |

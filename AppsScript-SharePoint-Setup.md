# SharePoint via Apps Script — Setup

Same idea as the Google Sheets Apps Script, but for files living on SharePoint / OneDrive. Browsers can't fetch SharePoint directly (CORS), so the Apps Script runs server-side, downloads each Excel file, and returns JSON to the dashboard.

## One-time setup (~10 min)

### 1. Make each SharePoint file shareable

For each file in OneDrive/SharePoint:
1. Right-click the file → **Share**
2. Set access to **"Anyone with the link can view"** (NOT "People in your organization")
3. Copy the share link
4. Append `&download=1` (if there's already a `?` in the URL) or replace `?web=1` with `?download=1`

Example:
- Share link: `https://company-my.sharepoint.com/:x:/g/personal/.../FILE?e=abc`
- Direct download: `https://company-my.sharepoint.com/:x:/g/personal/.../FILE?e=abc&download=1`

### 2. Create a new Apps Script project

1. Open https://script.google.com → **New project**
2. Delete the default `function myFunction(){}` from `Code.gs`
3. Paste the entire contents of `AppsScript-SharePointJSON.js` from this repo
4. Edit the `SOURCES` config at the top — replace the two URLs with your direct download links from step 1
5. Click the **disk icon** to save. Name the project something like "Purchase MIS SharePoint".

### 3. Add the Drive API service

The script uses the Drive API to convert downloaded Excel files. You have to enable it explicitly:

1. In the Apps Script editor, look at the left sidebar
2. Click the **+** next to **Services**
3. Find **Drive API** in the list → click → leave version as v2 → **Add**

### 4. Test it

1. Pick `testFetch` from the function dropdown at the top
2. Click ▶️ **Run**
3. First run will ask for permission — click **Review permissions** → choose your account → **Advanced** → "Go to (project name) (unsafe)" → **Allow**.
   - "Unsafe" just means it's a personal script Google hasn't formally verified. It's reading files you've shared with yourself — fine.
4. View → **Logs**. You should see something like:
   ```
   OK: true
   Elapsed: 18450 ms
     BOM 26-27 → 1531 rows × 40 cols
     Other material 26-27 → 1962 rows × 29 cols
     BOM 25-26 → 2950 rows × 36 cols
     Other material 25-26 → 7066 rows × 33 cols
   ```

If you get an HTML response error: the file isn't truly public. Re-share with **"Anyone with the link"**, not "People in your organization".

### 5. Deploy as Web App

1. Top-right → **Deploy** → **New deployment**
2. Click the gear ⚙️ next to "Select type" → choose **Web app**
3. Fill in:
   - **Description**: `Purchase MIS SharePoint`
   - **Execute as**: **Me** (so it can read files shared with you)
   - **Who has access**: **Anyone** (read-only — no edits possible)
4. Click **Deploy**.
5. **Copy the Web app URL**. Looks like:
   ```
   https://script.google.com/macros/s/AKfycbz.../exec
   ```

### 6. Connect the dashboard

1. Open the dashboard
2. **Data ▾** → **☁️ SharePoint (live, instant)**
3. Paste the Web app URL → **Save & Fetch**

Subsequent visits: just click **🔄 Refresh** in the header.

## Speeding it up with caching (recommended)

A fresh fetch from SharePoint takes 15-25 seconds. To make the dashboard load in 2-3 seconds instead, enable the cached endpoint:

1. In the Apps Script editor, run `setupCacheRefresh()` once.
   - This sets up a time trigger that refreshes the cache every 15 minutes in the background.
2. **Deploy → Manage deployments → ✏️ pencil → Version: New version**.
3. Under "Web app", change the function name from `doGet` to **`doGetCached`**.
4. **Deploy**. URL stays the same.

Now the dashboard hits a pre-warmed cache. Edits in the SharePoint file show up within ~15 minutes (next cache refresh) instead of waiting 25s on every load.

To force an immediate refresh: in the Apps Script editor, run `refreshCache()` manually.

## Updating later

If you change the script:
- **Deploy → Manage deployments → ✏️ pencil → Version: New version → Deploy**
- Web app URL stays the same.

## Security

- The Web app URL is **read-only**. Nobody can edit your files through it.
- Anyone with the URL **can read** the data. Don't post it publicly.
- To revoke: Apps Script → Deploy → Manage deployments → Archive.

## Troubleshooting

| Symptom | Fix |
|---|---|
| `returned HTML (login wall)` | File isn't truly public. Re-share with "Anyone with link" (NOT "People in org"). |
| `HTTP 403` | The share link expired or was revoked. Generate a fresh one. |
| `sheet not found: BOM` | Sheet tab name doesn't match. Check the tab name in Excel and update `SOURCES` in the script. |
| Times out (>60s) | The Excel files are huge. Enable the cached endpoint (see above) — first cache fill happens off the user's request. |
| `Drive API not enabled` | Step 3 above — add Drive API as a service. |

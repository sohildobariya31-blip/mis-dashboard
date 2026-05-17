# Google Apps Script — Auto-sync SharePoint Excel to Google Sheets

## Overview
This script automatically fetches your Excel files from SharePoint every 15 minutes and imports the data into Google Sheets. The dashboard then reads from the published Google Sheet — always up to date.

---

## Step 1: Get SharePoint Direct Download URLs

For each Excel file you want to sync:

1. Open SharePoint → navigate to the file
2. Click the **⋯** (three dots) next to the file → **Copy link**
3. Set sharing to **"Anyone with the link can view"** → Copy
4. The link looks like: `https://yourcompany.sharepoint.com/:x:/s/sitename/XXXXXXXXX?e=YYYY`
5. Modify it to a direct download URL:
   - Change `/:x:/` to `/personal/` or just add `&download=1` at the end
   - Or use this format: `https://yourcompany.sharepoint.com/sites/sitename/_layouts/15/download.aspx?UniqueId=XXXXX`

**Alternative (easier):** 
- Open the file in SharePoint → click **⋯** → **Download** 
- While it downloads, copy the download URL from browser's download manager
- That's your direct URL

---

## Step 2: Create Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com) → create a **new blank spreadsheet**
2. Name it: `Purchase Record - Live Sync`
3. Create sheets (tabs) matching your Excel:
   - Rename "Sheet1" to `BOM`
   - Add another sheet tab, name it `Other material`

---

## Step 3: Add the Apps Script

1. In your Google Sheet, go to **Extensions → Apps Script**
2. Delete any existing code in the editor
3. Paste the following script:

```javascript
/**
 * Purchase Record Auto-Sync from SharePoint
 * 
 * This script fetches Excel files from SharePoint and imports them into this Google Sheet.
 * Set up a time-based trigger to run every 5-15 minutes.
 */

// ============================================================
// CONFIGURATION — Edit these values
// ============================================================

const CONFIG = {
  files: [
    {
      // Current year file (26-27)
      url: 'https://solexsurat-my.sharepoint.com/:x:/g/personal/nikita_s_solex_in/IQBdsVh-DR2_QYhX1Cp4HMooASsR2kSY8uxpl80Lr0fMYZA?e=U274fA&download=1',
      sheets: ['BOM', 'Other material'],
    },
    {
      // Previous year file (25-26)
      url: 'https://solexsurat-my.sharepoint.com/:x:/g/personal/svikruti_p_solex_in/IQARFOs9snpGS6ASbA6Euu-xAd-r7MWWL-aHCg2TXXBhm3k?e=xJOOHI&download=1',
      sheets: ['BOM', 'Other material'],
    },
  ],
};

// ============================================================
// MAIN SYNC FUNCTION
// ============================================================

function syncFromSharePoint() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  for (const file of CONFIG.files) {
    if (!file.url || file.url.includes('PASTE_YOUR')) {
      Logger.log('Skipping unconfigured file URL');
      continue;
    }
    
    try {
      Logger.log('Fetching: ' + file.url);
      const response = UrlFetchApp.fetch(file.url, {
        followRedirects: true,
        muteHttpExceptions: true,
      });
      
      if (response.getResponseCode() !== 200) {
        Logger.log('Error fetching file: HTTP ' + response.getResponseCode());
        continue;
      }
      
      const blob = response.getBlob();
      const tempFile = DriveApp.createFile(blob.setName('temp_import.xlsx'));
      
      try {
        // Convert Excel to Google Sheets format to read it
        const tempSS = SpreadsheetApp.open(
          Drive.Files.copy(
            { title: 'temp_convert', mimeType: MimeType.GOOGLE_SHEETS },
            tempFile.getId()
          )
        );
        
        for (const sheetName of file.sheets) {
          const sourceSheet = tempSS.getSheetByName(sheetName);
          if (!sourceSheet) {
            Logger.log('Sheet "' + sheetName + '" not found in file, skipping');
            continue;
          }
          
          // Get or create destination sheet
          // Use a combined name if multiple files have same sheet name
          const destName = CONFIG.files.length > 1 && CONFIG.files.indexOf(file) > 0
            ? sheetName + ' (prev)'
            : sheetName;
          
          let destSheet = ss.getSheetByName(destName);
          if (!destSheet) {
            destSheet = ss.insertSheet(destName);
          }
          
          // Clear existing data
          destSheet.clear();
          
          // Copy data
          const data = sourceSheet.getDataRange().getValues();
          if (data.length > 0) {
            destSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
            Logger.log('Imported ' + data.length + ' rows into "' + destName + '"');
          }
        }
        
        // Clean up temp converted file
        DriveApp.getFileById(tempSS.getId()).setTrashed(true);
        
      } finally {
        // Clean up temp xlsx file
        tempFile.setTrashed(true);
      }
      
    } catch (error) {
      Logger.log('Error processing file: ' + error.message);
    }
  }
  
  // Update last sync timestamp
  const ss2 = SpreadsheetApp.getActiveSpreadsheet();
  let metaSheet = ss2.getSheetByName('_sync_log');
  if (!metaSheet) metaSheet = ss2.insertSheet('_sync_log');
  metaSheet.getRange('A1').setValue('Last sync');
  metaSheet.getRange('B1').setValue(new Date().toLocaleString());
  
  Logger.log('Sync complete at ' + new Date().toLocaleString());
}

// ============================================================
// MANUAL TEST — Run this first to verify it works
// ============================================================

function testSync() {
  syncFromSharePoint();
  SpreadsheetApp.getUi().alert('Sync complete! Check the sheets.');
}

// ============================================================
// SETUP TRIGGER — Run this once to set up auto-sync
// ============================================================

function setupAutoSync() {
  // Remove existing triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  
  // Create new trigger: run every 15 minutes
  ScriptApp.newTrigger('syncFromSharePoint')
    .timeBased()
    .everyMinutes(15)
    .create();
  
  SpreadsheetApp.getUi().alert(
    'Auto-sync set up! The script will run every 15 minutes.\n\n' +
    'You can change the interval by editing the setupAutoSync function.'
  );
}
```

4. Click **Save** (💾 icon) — name the project "SharePoint Sync"

---

## Step 4: Enable Drive API

1. In the Apps Script editor, click **Services** (+ icon on the left panel)
2. Find **Drive API** → click **Add**
3. This is needed for the Excel-to-Sheets conversion

---

## Step 5: Configure URLs

1. In the script, replace `PASTE_YOUR_SHAREPOINT_DOWNLOAD_URL_FOR_26-27_FILE_HERE` with your actual SharePoint download URL
2. Same for the 25-26 file (or remove that block if not needed)
3. Save

---

## Step 6: Test

1. Click the function dropdown (top bar) → select **testSync**
2. Click **Run** (▶️)
3. First time: it will ask for permissions — click "Review Permissions" → choose your Google account → "Allow"
4. Wait for it to complete (may take 30-60 seconds for large files)
5. Check your Google Sheet — the BOM and Other material tabs should now have data

---

## Step 7: Set Up Auto-Sync

1. Select **setupAutoSync** from the function dropdown
2. Click **Run**
3. It will confirm: "Auto-sync set up! Runs every 15 minutes."

---

## Step 8: Publish to Web

1. In your Google Sheet, go to **File → Share → Publish to web**
2. Select **BOM** tab → format: **CSV** → click **Publish** → copy the URL
3. Repeat for **Other material** tab
4. Paste these URLs into the dashboard's Google Sheets connection

---

## Done!

Now the flow is:
```
SharePoint Excel → (every 15 min) → Google Sheets → (published CSV) → Dashboard (live)
```

- Team edits Excel in SharePoint as usual
- Script auto-copies data to Google Sheets every 15 minutes
- Dashboard reads from Google Sheets published CSV
- Everyone sees updated data within 15 minutes of any change

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Access denied" when fetching | Make sure SharePoint file is shared with "Anyone with link can view" |
| "Sheet not found" | Check that sheet names in CONFIG match exactly (case-sensitive) |
| Script times out | Large files (>5000 rows) may need the script split into smaller chunks |
| Data not updating | Check Extensions → Apps Script → Executions to see if triggers are running |
| Want faster sync | Change `everyMinutes(15)` to `everyMinutes(5)` in setupAutoSync |

---

## Notes

- The `_sync_log` sheet shows when the last sync happened
- Previous year data goes into sheets named "BOM (prev)" and "Other material (prev)"
- The script runs under YOUR Google account — only you need permissions
- Free Google accounts have a limit of ~90 minutes/day of script runtime (plenty for this)

/**
 * Purchase MIS — Live JSON endpoint
 * ==================================================
 * Exposes the linked Google Sheets as a JSON Web App so the
 * dashboard can read fresh data in seconds (bypassing the
 * 5-15 minute Publish-to-web CSV cache).
 *
 * SETUP (one-time, takes ~5 minutes)
 * ------------------------------------------------
 * 1. Open EITHER of the two source Google Sheets in your browser.
 * 2. Extensions → Apps Script. Delete any existing code in Code.gs.
 * 3. Paste THIS file in.
 * 4. Save (disk icon).
 * 5. Click "Deploy" (top-right) → "New deployment"
 *      • Type      → "Web app"
 *      • Description → "Purchase MIS Live"
 *      • Execute as → "Me (your@email.com)"
 *      • Who has access → "Anyone"   (read-only, no edits)
 *    Click "Deploy".
 * 6. Authorize (it will pop a Google login → "Advanced" → "Go to ... unsafe").
 *    This is normal for personal Apps Script.
 * 7. Copy the "Web app URL" — looks like:
 *    https://script.google.com/macros/s/AKfycbz.../exec
 * 8. Paste that URL into the dashboard:
 *    Data ▾  →  "🔗 Apps Script live endpoint"  →  paste URL  →  Save
 *    The dashboard will now fetch live data in ~3-8 seconds.
 *
 * To redeploy after editing:
 *    Deploy → Manage deployments → ✏️ pencil → New version → Deploy
 *
 * SECURITY NOTE
 * ------------------------------------------------
 * The Web App URL is read-only (no writes from the dashboard).
 * Anyone with the URL can read your published sheet data, so keep
 * it private — share only with people you trust.
 */

// ==============================================================
// CONFIG — fill in the spreadsheet IDs of your two purchase files
// ==============================================================
//
// Open the spreadsheet, look at the URL:
//   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
// Copy the SPREADSHEET_ID and paste below.

var SOURCES = [
  {
    id: 'PASTE_26-27_SPREADSHEET_ID_HERE',
    label: '26-27',
    sheets: [
      { name: 'BOM',            tag: 'BOM 26-27' },
      { name: 'Other material', tag: 'Other material 26-27' },
    ],
  },
  {
    id: 'PASTE_25-26_SPREADSHEET_ID_HERE',
    label: '25-26',
    sheets: [
      { name: 'BOM',            tag: 'BOM 25-26' },
      { name: 'Other material', tag: 'Other material 25-26' },
    ],
  },
];

// ==============================================================
// MAIN entry point — invoked when dashboard hits the Web App URL
// ==============================================================
function doGet(e) {
  var t0 = Date.now();
  var out = { ok: true, fetchedAt: new Date().toISOString(), sheets: [] };

  try {
    for (var i = 0; i < SOURCES.length; i++) {
      var src = SOURCES[i];
      if (!src.id || src.id.indexOf('PASTE_') === 0) continue; // skip placeholders

      var ss;
      try {
        ss = SpreadsheetApp.openById(src.id);
      } catch (err) {
        out.sheets.push({ tag: src.label, error: 'cannot open: ' + err.message });
        continue;
      }

      for (var j = 0; j < src.sheets.length; j++) {
        var spec = src.sheets[j];
        var sheet = ss.getSheetByName(spec.name);

        // Case-insensitive fallback
        if (!sheet) {
          var all = ss.getSheets();
          for (var k = 0; k < all.length; k++) {
            if (all[k].getName().toLowerCase().indexOf(spec.name.toLowerCase()) !== -1) {
              sheet = all[k];
              break;
            }
          }
        }
        if (!sheet) {
          out.sheets.push({ tag: spec.tag, error: 'sheet not found: ' + spec.name });
          continue;
        }

        var rows = sheet.getLastRow();
        var cols = sheet.getLastColumn();
        if (rows === 0 || cols === 0) {
          out.sheets.push({ tag: spec.tag, rows: 0, data: [] });
          continue;
        }

        // Use display values so dates appear formatted, not as serials
        var values = sheet.getRange(1, 1, rows, cols).getDisplayValues();
        out.sheets.push({
          tag: spec.tag,
          rows: rows,
          cols: cols,
          data: values,
        });
      }
    }

    out.elapsedMs = Date.now() - t0;
  } catch (e) {
    out.ok = false;
    out.error = e.message;
    out.stack = e.stack;
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================
// Quick test from the editor — run this to confirm the script
// can read both sheets before deploying
// ==============================================================
function testFetch() {
  var fake = doGet({});
  var out = JSON.parse(fake.getContent());
  Logger.log('OK: ' + out.ok);
  Logger.log('Elapsed: ' + out.elapsedMs + ' ms');
  for (var i = 0; i < out.sheets.length; i++) {
    var s = out.sheets[i];
    Logger.log('  ' + s.tag + ' → ' + (s.error || (s.rows + ' rows × ' + s.cols + ' cols')));
  }
}

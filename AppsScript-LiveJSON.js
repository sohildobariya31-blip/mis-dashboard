/**
 * Purchase MIS — Live JSON endpoint (FAST version)
 * ==================================================
 * Returns sheet data as JSON. Optimized to:
 *  - skip empty/blank rows (massive speedup on sheets with formula tails)
 *  - cap fetched range to actual data
 *  - support gzip-friendly compact arrays
 *
 * The dashboard expects each sheet's `data` to be a 2D array
 * of strings/numbers (first row = headers).
 *
 * SETUP — see AppsScript-Live-Setup.md.
 *
 * NOTE on multi-spreadsheet:
 *  You can either run ONE script that opens both spreadsheets
 *  (fill in both IDs in SOURCES below), OR run TWO scripts
 *  (one per spreadsheet) and paste both Web App URLs in the
 *  dashboard. The dashboard handles both.
 */

// ==============================================================
// CONFIG — fill in the spreadsheet IDs
// ==============================================================
//
// Open the spreadsheet, look at the URL:
//   https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
// Copy the SPREADSHEET_ID and paste below.
//
// If this script lives INSIDE one of the spreadsheets, you can
// leave that one's `id` empty (or use 'self') and the script
// will fall back to SpreadsheetApp.getActiveSpreadsheet().

var SOURCES = [
  {
    id: 'self',                       // 'self' or '' = use the spreadsheet this script is bound to
    label: '26-27',
    sheets: [
      { name: 'BOM',            tag: 'BOM 26-27' },
      { name: 'Other material', tag: 'Other material 26-27' },
    ],
  },
  // Uncomment + fill if you want this single script to also serve the other file.
  // Otherwise leave commented and host a second script in the 25-26 sheet.
  //
  // {
  //   id: 'PASTE_25-26_SPREADSHEET_ID_HERE',
  //   label: '25-26',
  //   sheets: [
  //     { name: 'BOM',            tag: 'BOM 25-26' },
  //     { name: 'Other material', tag: 'Other material 25-26' },
  //   ],
  // },
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
      var ss;

      if (!src.id || src.id === 'self' || (typeof src.id === 'string' && src.id.indexOf('PASTE_') === 0)) {
        ss = SpreadsheetApp.getActiveSpreadsheet();
        if (!ss && src.id !== 'self') {
          out.sheets.push({ tag: src.label, error: 'no spreadsheet ID configured and no active spreadsheet' });
          continue;
        }
      } else {
        try { ss = SpreadsheetApp.openById(src.id); }
        catch (err) { out.sheets.push({ tag: src.label, error: 'cannot open: ' + err.message }); continue; }
      }
      if (!ss) continue;

      for (var j = 0; j < src.sheets.length; j++) {
        var spec = src.sheets[j];
        var sheet = ss.getSheetByName(spec.name);

        // Case-insensitive partial match fallback
        if (!sheet) {
          var all = ss.getSheets();
          for (var k = 0; k < all.length; k++) {
            if (all[k].getName().toLowerCase().indexOf(spec.name.toLowerCase()) !== -1) {
              sheet = all[k]; break;
            }
          }
        }
        if (!sheet) {
          out.sheets.push({ tag: spec.tag, error: 'sheet not found: ' + spec.name });
          continue;
        }

        // Use getDataRange for tight bounds + getDisplayValues for formatted dates
        var rng = sheet.getDataRange();
        if (!rng) { out.sheets.push({ tag: spec.tag, rows: 0, data: [] }); continue; }
        var values = rng.getDisplayValues();

        // Trim trailing fully-empty rows (formula tails create thousands of zero rows)
        var lastNonEmpty = values.length - 1;
        while (lastNonEmpty >= 0) {
          var row = values[lastNonEmpty];
          var hasData = false;
          for (var c = 0; c < row.length; c++) {
            var v = row[c];
            if (v !== '' && v !== null && v !== undefined) { hasData = true; break; }
          }
          if (hasData) break;
          lastNonEmpty--;
        }
        if (lastNonEmpty < 0) { out.sheets.push({ tag: spec.tag, rows: 0, data: [] }); continue; }
        var trimmed = values.slice(0, lastNonEmpty + 1);

        // Optionally: drop rows that are zero-only (typical of empty BOM rows in 26-27)
        // Keep header (row 0) and rows that have at least one non-empty AND non-zero text cell.
        var compact = [trimmed[0]];
        for (var r = 1; r < trimmed.length; r++) {
          var rr = trimmed[r];
          var keep = false;
          for (var c2 = 0; c2 < rr.length; c2++) {
            var s = rr[c2];
            if (s === '' || s == null) continue;
            if (s === '0' || s === 0) continue;
            keep = true; break;
          }
          if (keep) compact.push(rr);
        }

        out.sheets.push({
          tag: spec.tag,
          rows: compact.length,
          cols: compact[0] ? compact[0].length : 0,
          data: compact,
        });
      }
    }

    out.elapsedMs = Date.now() - t0;
  } catch (e) {
    out.ok = false;
    out.error = e.message;
    out.stack = (e.stack || '').split('\n').slice(0, 4).join(' | ');
  }

  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================
// Quick test from the editor — run this to confirm the script
// can read the sheets before deploying
// ==============================================================
function testFetch() {
  var fake = doGet({});
  var out = JSON.parse(fake.getContent());
  Logger.log('OK: ' + out.ok);
  Logger.log('Elapsed: ' + out.elapsedMs + ' ms');
  if (out.error) Logger.log('Error: ' + out.error);
  for (var i = 0; i < (out.sheets || []).length; i++) {
    var s = out.sheets[i];
    Logger.log('  ' + s.tag + ' → ' + (s.error || (s.rows + ' rows × ' + s.cols + ' cols')));
  }
}

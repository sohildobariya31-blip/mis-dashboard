/**
 * Purchase MIS — SharePoint live JSON endpoint (FAST version)
 * ==================================================
 * Fetches Excel files from SharePoint / OneDrive and returns
 * sheet data as JSON. The dashboard reads this exactly like
 * the Google Sheets Apps Script — same UX, same picker, etc.
 *
 * Why a script: browsers block direct SharePoint fetches via
 * CORS. Apps Script runs server-side and isn't subject to CORS.
 *
 * SETUP — see AppsScript-SharePoint-Setup.md.
 */

// ==============================================================
// CONFIG — fill in your SharePoint/OneDrive direct download URLs
// ==============================================================
//
// Each entry is one Excel file. Get the URL by:
//  1. Open file in OneDrive/SharePoint
//  2. Share → "Anyone with the link can view"
//  3. Copy the share link
//  4. Append &download=1 (or replace ?web=1 with ?download=1)
//
// `tags`: which sheets you want exposed and what the dashboard
// should call them. The dashboard's filter expects names like
// "BOM 26-27" / "Other material 26-27".

var SOURCES = [
  {
    label: '26-27',
    url: 'https://solexsurat-my.sharepoint.com/:x:/g/personal/nikita_s_solex_in/IQBdsVh-DR2_QYhX1Cp4HMooASsR2kSY8uxpl80Lr0fMYZA?download=1',
    sheets: [
      { name: 'BOM',            tag: 'BOM 26-27' },
      { name: 'Other material', tag: 'Other material 26-27' },
    ],
  },
  {
    label: '25-26',
    url: 'https://solexsurat-my.sharepoint.com/:x:/g/personal/svikruti_p_solex_in/IQARFOs9snpGS6ASbA6Euu-xAd-r7MWWL-aHCg2TXXBhm3k?download=1',
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
      if (!src.url) {
        out.sheets.push({ tag: src.label, error: 'no URL configured' });
        continue;
      }

      // Step 1: download the xlsx blob from SharePoint
      var blob;
      try {
        var resp = UrlFetchApp.fetch(src.url, {
          followRedirects: true,
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0' },
        });
        var code = resp.getResponseCode();
        var ct = resp.getHeaders()['Content-Type'] || '';
        if (code !== 200) {
          out.sheets.push({ tag: src.label, error: 'HTTP ' + code });
          continue;
        }
        if (ct.indexOf('html') !== -1) {
          out.sheets.push({ tag: src.label, error: 'returned HTML (login wall) — re-share with "Anyone with link"' });
          continue;
        }
        blob = resp.getBlob().setName(src.label + '.xlsx');
      } catch (err) {
        out.sheets.push({ tag: src.label, error: 'fetch error: ' + err.message });
        continue;
      }

      // Step 2: convert to Google Spreadsheet (Drive) so we can read it
      var convertedFileId = null;
      var tempFileId = null;
      try {
        var tempFile = DriveApp.createFile(blob);
        tempFileId = tempFile.getId();
        var converted = Drive.Files.copy(
          { title: '_temp_sp_' + src.label, mimeType: 'application/vnd.google-apps.spreadsheet' },
          tempFileId
        );
        convertedFileId = converted.id;
        var ss = SpreadsheetApp.openById(convertedFileId);

        // Step 3: extract each requested sheet
        for (var j = 0; j < src.sheets.length; j++) {
          var spec = src.sheets[j];
          var sheet = ss.getSheetByName(spec.name);
          if (!sheet) {
            // Case-insensitive partial match fallback
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

          var rng = sheet.getDataRange();
          if (!rng) { out.sheets.push({ tag: spec.tag, rows: 0, data: [] }); continue; }
          var values = rng.getDisplayValues();

          // Trim trailing empty rows
          var lastNonEmpty = values.length - 1;
          while (lastNonEmpty >= 0) {
            var row = values[lastNonEmpty];
            var hasData = false;
            for (var c = 0; c < row.length; c++) {
              var v = row[c];
              if (v !== '' && v != null) { hasData = true; break; }
            }
            if (hasData) break;
            lastNonEmpty--;
          }
          if (lastNonEmpty < 0) { out.sheets.push({ tag: spec.tag, rows: 0, data: [] }); continue; }
          var trimmed = values.slice(0, lastNonEmpty + 1);

          // Drop zero-only / blank rows beyond the header
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
      } catch (err) {
        out.sheets.push({ tag: src.label, error: 'conversion error: ' + err.message });
      } finally {
        // Always clean up temp files
        try { if (convertedFileId) DriveApp.getFileById(convertedFileId).setTrashed(true); } catch (e) {}
        try { if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true); } catch (e) {}
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
// can read the SharePoint files before deploying as Web App
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

// ==============================================================
// Caching version (recommended for production):
// SharePoint fetches are slow (~10-20s). To make the dashboard
// snappier, the doGetCached function reads from a cache that's
// refreshed every 15 minutes by a time-based trigger.
//
// To enable:
//   1. Run setupCacheRefresh() once
//   2. In your Web App deployment, change the entry point from
//      doGet to doGetCached (re-deploy as new version)
// ==============================================================
function setupCacheRefresh() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'refreshCache') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('refreshCache').timeBased().everyMinutes(15).create();
  refreshCache();
  SpreadsheetApp.getUi && SpreadsheetApp.getUi().alert('✓ Cache will refresh every 15 min.');
}

function refreshCache() {
  var fake = doGet({});
  var json = fake.getContent();
  var cache = CacheService.getScriptCache();
  // Cache up to ~100KB chunks (Apps Script cache limit is 100KB per key)
  var chunkSize = 90 * 1024;
  var chunks = Math.ceil(json.length / chunkSize);
  cache.put('sp_chunks', String(chunks), 21600);
  for (var i = 0; i < chunks; i++) {
    cache.put('sp_chunk_' + i, json.slice(i * chunkSize, (i + 1) * chunkSize), 21600);
  }
  Logger.log('Cached ' + json.length + ' bytes in ' + chunks + ' chunks');
}

function doGetCached(e) {
  var cache = CacheService.getScriptCache();
  var chunks = parseInt(cache.get('sp_chunks') || '0', 10);
  if (!chunks) {
    // Cache empty — refresh now (slow first call)
    refreshCache();
    chunks = parseInt(cache.get('sp_chunks') || '0', 10);
  }
  var pieces = [];
  for (var i = 0; i < chunks; i++) {
    var p = cache.get('sp_chunk_' + i);
    if (!p) { return doGet(e); /* cache miss — fall back to live */ }
    pieces.push(p);
  }
  return ContentService
    .createTextOutput(pieces.join(''))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Purchase Record Auto-Sync from SharePoint/OneDrive
 * 
 * SETUP:
 * 1. Paste this in Extensions → Apps Script
 * 2. Click Services (+) → Add "Drive API"
 * 3. Run "testUrlAccess" to check if URL works
 * 4. If works, run "testSync" to import data
 * 5. Run "setupAutoSync" for auto-refresh
 * 6. File → Share → Publish to web → each sheet → CSV
 */

// ============================================================
// YOUR SHAREPOINT LINKS (Anyone with link can view)
// ============================================================

var FILE_1_URL = 'https://solexsurat-my.sharepoint.com/:x:/g/personal/nikita_s_solex_in/IQBdsVh-DR2_QYhX1Cp4HMooASsR2kSY8uxpl80Lr0fMYZA?download=1';
var FILE_2_URL = 'https://solexsurat-my.sharepoint.com/:x:/g/personal/svikruti_p_solex_in/IQARFOs9snpGS6ASbA6Euu-xAd-r7MWWL-aHCg2TXXBhm3k?download=1';

var SHEETS_TO_IMPORT = ['BOM', 'Other material'];

// ============================================================
// TEST URL ACCESS — Run this FIRST
// ============================================================

function testUrlAccess() {
  var url = FILE_1_URL;
  Logger.log('Testing: ' + url);
  
  var options = {
    followRedirects: true,
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  };
  
  try {
    var response = UrlFetchApp.fetch(url, options);
    var code = response.getResponseCode();
    var ct = response.getHeaders()['Content-Type'] || 'unknown';
    var size = response.getContent().length;
    
    Logger.log('HTTP Status: ' + code);
    Logger.log('Content-Type: ' + ct);
    Logger.log('Size: ' + size + ' bytes');
    
    if (code === 200 && ct.indexOf('html') === -1 && size > 1000) {
      SpreadsheetApp.getUi().alert(
        '✅ SUCCESS!\n\n' +
        'HTTP: ' + code + '\n' +
        'Type: ' + ct + '\n' +
        'Size: ' + size + ' bytes\n\n' +
        'The URL is accessible. Now run "testSync".'
      );
    } else if (ct.indexOf('html') !== -1) {
      var body = response.getContentText().substring(0, 200);
      Logger.log('HTML body: ' + body);
      SpreadsheetApp.getUi().alert(
        '❌ FAILED - Got HTML page (login required)\n\n' +
        'HTTP: ' + code + '\n' +
        'The file is not publicly accessible.\n\n' +
        'Make sure sharing is set to:\n"Anyone with the link can view"\n\n' +
        'NOT "People in Solex Energy with the link"'
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '⚠️ Unexpected response\n\n' +
        'HTTP: ' + code + '\n' +
        'Type: ' + ct + '\n' +
        'Size: ' + size + '\n\n' +
        'Check the execution log for details.'
      );
    }
  } catch (e) {
    Logger.log('ERROR: ' + e.message);
    SpreadsheetApp.getUi().alert('❌ ERROR:\n' + e.message);
  }
}

// ============================================================
// TEST URL with alternate format
// ============================================================

function testUrlAlternate() {
  // Try different URL formats for OneDrive personal
  var baseId = 'IQBdsVh-DR2_QYhX1Cp4HMooASsR2kSY8uxpl80Lr0fMYZA';
  var urls = [
    FILE_1_URL,
    'https://solexsurat-my.sharepoint.com/personal/nikita_s_solex_in/_layouts/15/download.aspx?share=' + baseId,
    FILE_1_URL.replace('?download=1', '?e=mfnRkL&download=1'),
  ];
  
  for (var i = 0; i < urls.length; i++) {
    Logger.log('\n--- Trying URL format ' + (i+1) + ' ---');
    Logger.log(urls[i]);
    try {
      var response = UrlFetchApp.fetch(urls[i], {
        followRedirects: true,
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      var code = response.getResponseCode();
      var ct = response.getHeaders()['Content-Type'] || '';
      var size = response.getContent().length;
      Logger.log('Status: ' + code + ' | Type: ' + ct + ' | Size: ' + size);
      
      if (code === 200 && ct.indexOf('html') === -1 && size > 1000) {
        Logger.log('✅ THIS FORMAT WORKS!');
        SpreadsheetApp.getUi().alert('✅ Format ' + (i+1) + ' works!\n\nURL: ' + urls[i] + '\n\nSize: ' + size + ' bytes');
        return;
      }
    } catch (e) {
      Logger.log('Error: ' + e.message);
    }
  }
  
  SpreadsheetApp.getUi().alert(
    '❌ None of the URL formats worked.\n\n' +
    'This means the file requires authentication that Google cannot provide.\n\n' +
    'Solution: Use manual upload in the dashboard instead.\n' +
    'Your files are synced via OneDrive to your PC — just upload from there.'
  );
}

// ============================================================
// SYNC — imports data from SharePoint to this Google Sheet
// ============================================================

function syncFromSharePoint() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var totalRows = 0;
  var files = [
    { name: '26-27', url: FILE_1_URL },
    { name: '25-26', url: FILE_2_URL }
  ];
  
  for (var fi = 0; fi < files.length; fi++) {
    var file = files[fi];
    Logger.log('\n=== Processing: ' + file.name + ' ===');
    
    try {
      var response = UrlFetchApp.fetch(file.url, {
        followRedirects: true,
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      
      if (response.getResponseCode() !== 200) {
        Logger.log('HTTP Error: ' + response.getResponseCode());
        continue;
      }
      
      var ct = response.getHeaders()['Content-Type'] || '';
      if (ct.indexOf('html') !== -1) {
        Logger.log('Got HTML - file not accessible');
        continue;
      }
      
      var blob = response.getBlob().setName(file.name + '.xlsx');
      var tempFile = DriveApp.createFile(blob);
      
      try {
        var resource = { title: '_temp_' + file.name, mimeType: 'application/vnd.google-apps.spreadsheet' };
        var converted = Drive.Files.copy(resource, tempFile.getId());
        var tempSS = SpreadsheetApp.openById(converted.id);
        
        for (var si = 0; si < SHEETS_TO_IMPORT.length; si++) {
          var sheetName = SHEETS_TO_IMPORT[si];
          var src = tempSS.getSheetByName(sheetName);
          
          // Case-insensitive fallback
          if (!src) {
            var all = tempSS.getSheets();
            for (var s = 0; s < all.length; s++) {
              if (all[s].getName().toLowerCase().indexOf(sheetName.toLowerCase()) !== -1) {
                src = all[s];
                sheetName = all[s].getName();
                break;
              }
            }
          }
          if (!src) { Logger.log('  Sheet "' + sheetName + '" not found'); continue; }
          
          var destName = fi === 0 ? sheetName : sheetName + ' (25-26)';
          var dest = ss.getSheetByName(destName) || ss.insertSheet(destName);
          dest.clear();
          
          var rows = src.getLastRow();
          var cols = src.getLastColumn();
          if (rows > 0 && cols > 0) {
            dest.getRange(1, 1, rows, cols).setValues(src.getRange(1, 1, rows, cols).getValues());
            totalRows += rows;
            Logger.log('  ✓ ' + rows + ' rows → "' + destName + '"');
          }
        }
        
        DriveApp.getFileById(converted.id).setTrashed(true);
      } finally {
        tempFile.setTrashed(true);
      }
    } catch (e) {
      Logger.log('ERROR: ' + e.message + '\n' + e.stack);
    }
  }
  
  // Log
  var log = ss.getSheetByName('_sync_log') || ss.insertSheet('_sync_log');
  log.getRange('A1').setValue('Last sync');
  log.getRange('B1').setValue(new Date());
  log.getRange('A2').setValue('Rows');
  log.getRange('B2').setValue(totalRows);
  log.getRange('A3').setValue('Status');
  log.getRange('B3').setValue(totalRows > 0 ? 'OK' : 'FAILED');
  
  Logger.log('Done: ' + totalRows + ' rows');
}

function testSync() {
  syncFromSharePoint();
  SpreadsheetApp.getUi().alert('Done! Check sheet tabs and View → Execution log.');
}

function setupAutoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  ScriptApp.newTrigger('syncFromSharePoint').timeBased().everyMinutes(15).create();
  SpreadsheetApp.getUi().alert('✓ Auto-sync every 15 min.\n\nNext: File → Share → Publish to web → CSV');
}

function stopAutoSync() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  SpreadsheetApp.getUi().alert('Stopped.');
}

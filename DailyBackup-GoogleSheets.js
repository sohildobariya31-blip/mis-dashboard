/**
 * DAILY BACKUP — Simple version
 * Paste in Extensions → Apps Script → Run "setupDailyBackup"
 */

function backupNow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  var backupName = 'Backup_' + today + '_' + ss.getName();
  
  // Create backup folder if needed
  var folders = DriveApp.getFoldersByName('MIS_Backups');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('MIS_Backups');
  
  // Just copy the entire spreadsheet
  var copy = ss.copy(backupName);
  var file = DriveApp.getFileById(copy.getId());
  file.moveTo(folder);
  
  Logger.log('Backup done: ' + backupName);
}

function setupDailyBackup() {
  // Remove old triggers
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  // Daily at 1 AM
  ScriptApp.newTrigger('backupNow').timeBased().atHour(1).everyDays(1).inTimezone('Asia/Kolkata').create();
  SpreadsheetApp.getUi().alert('Done! Daily backup at 1 AM to "MIS_Backups" folder.');
}

function stopBackup() {
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert('Stopped.');
}

/**
 * SETUP SCRIPT FOR BOM SHEET (26-27 Purchase Record)
 * 
 * HOW TO USE:
 * 1. Upload "26-27-Purchase Record.xlsx" to Google Drive
 * 2. Open it with Google Sheets
 * 3. Go to Extensions → Apps Script
 * 4. Paste this entire script
 * 5. Run "setupSheet" function
 * 6. Then: File → Share → Publish to web → BOM tab → CSV → copy URL for dashboard
 */

function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('BOM');
  if (!ws) { SpreadsheetApp.getUi().alert('BOM sheet not found!'); return; }
  
  var lastRow = ws.getLastRow();
  Logger.log('BOM sheet has ' + lastRow + ' rows');
  
  // Add formulas to calculated columns (only for rows without values)
  // Col U (21) = Amount = Qty(I) * Rate(T)
  // Col X (24) = Total Amount = (U+V)+(U+V)*W
  // Col Y (25) = Payable = X - Z
  // Col AB (28) = Due date = L + AA
  // Col AE (31) = Outstanding = Y - AD (if exists)
  
  for (var row = 2; row <= lastRow; row++) {
    var qty = ws.getRange(row, 9).getValue();
    var rate = ws.getRange(row, 20).getValue();
    var amount = ws.getRange(row, 21).getValue();
    
    // Only add formula if Amount is empty but Qty and Rate exist
    if (!amount && qty && rate) {
      ws.getRange(row, 21).setFormula('=I' + row + '*T' + row);
    }
    
    var totalAmt = ws.getRange(row, 24).getValue();
    if (!totalAmt && amount) {
      ws.getRange(row, 24).setFormula('=(U' + row + '+V' + row + ')+((U' + row + '+V' + row + ')*W' + row + ')');
    }
    
    var payable = ws.getRange(row, 25).getValue();
    if (!payable && totalAmt) {
      ws.getRange(row, 25).setFormula('=X' + row + '-Z' + row);
    }
  }
  
  // Add data validation dropdowns
  var listSheet = ss.getSheetByName('List');
  if (!listSheet) {
    listSheet = ss.insertSheet('List');
    // Populate lists
    var lists = {
      1: ['Category', 'RM', 'PM'],
      2: ['Module type', 'Perc+M10R+G12R', 'Perc+M10R', 'M10R+G12R', 'M10R', 'G12R', 'Perc+Topcon'],
      3: ['UOM', 'Nos', 'Kgs', 'Mtrs', 'Sqm', 'Rolls', 'Sets', 'Ltrs', 'Pcs', 'NOS'],
      4: ['Status', 'Received', 'In Transit', 'On Hold', 'Tentative', 'Cancelled', 'Short Close'],
      5: ['COO', 'India', 'China', 'Malaysia', 'Indonesia', 'Thailand', 'Vietnam', 'Taiwan'],
      6: ['Currency', 'INR', 'USD', 'RMB', 'EUR'],
      7: ['Payment Status', 'Approved', 'Paid', 'Pending', 'Hold'],
    };
    for (var col in lists) {
      var vals = lists[col];
      for (var i = 0; i < vals.length; i++) {
        listSheet.getRange(i + 1, parseInt(col)).setValue(vals[i]);
      }
    }
  }
  
  // Apply data validations
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Received', 'In Transit', 'On Hold', 'Tentative', 'Cancelled', 'Short Close'], true)
    .setAllowInvalid(true).build();
  ws.getRange('N2:N' + lastRow).setDataValidation(statusRule);
  
  var currencyRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['INR', 'USD', 'RMB', 'EUR'], true)
    .setAllowInvalid(true).build();
  ws.getRange('S2:S' + lastRow).setDataValidation(currencyRule);
  
  var categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['RM', 'PM'], true)
    .setAllowInvalid(true).build();
  ws.getRange('A2:A' + lastRow).setDataValidation(categoryRule);
  
  // Conditional formatting
  var rules = ws.getConditionalFormatRules();
  
  // Status = Received → green
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Received')
    .setBackground('#C6EFCE')
    .setRanges([ws.getRange('N2:N' + lastRow)])
    .build());
  
  // Status = In Transit → yellow
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('In Transit')
    .setBackground('#FFEB9C')
    .setRanges([ws.getRange('N2:N' + lastRow)])
    .build());
  
  // Status = Cancelled → red
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cancelled')
    .setBackground('#FFC7CE')
    .setRanges([ws.getRange('N2:N' + lastRow)])
    .build());
  
  ws.setConditionalFormatRules(rules);
  
  // Freeze header row
  ws.setFrozenRows(1);
  
  // Format header
  var headerRange = ws.getRange(1, 1, 1, ws.getLastColumn());
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1F4E79');
  headerRange.setFontColor('#FFFFFF');
  
  SpreadsheetApp.getUi().alert(
    '✓ BOM sheet setup complete!\n\n' +
    '• Formulas added for Amount, Total, Payable\n' +
    '• Dropdown validations for Status, Currency, Category\n' +
    '• Conditional formatting for status colors\n' +
    '• Header formatted and frozen\n\n' +
    'Rows: ' + lastRow + '\n\n' +
    'Next: File → Share → Publish to web → BOM → CSV'
  );
}

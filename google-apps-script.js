/**
 * ===================================================================
 * Bruker World Cup 2026 — Google Apps Script
 * ===================================================================
 *
 * HOW TO INSTALL:
 * 1. Open the Google Sheet:
 *    https://docs.google.com/spreadsheets/d/1zVoXUcy-17ctpic7e75rXhJ23XhLG2YKi1PbDtGeJME/edit
 * 2. Go to Extensions → Apps Script
 * 3. Delete any existing code and paste this entire file
 * 4. Click Save (Ctrl+S)
 * 5. Run the function "setupSheetTabs" from the dropdown & click ▶ Run
 *    — It will ask for permissions (allow it)
 *    — It creates the "GroupMatches" and "KnockoutResults" tabs with headers
 *
 * OPTIONAL — Auto-update scores:
 * - You can manually fill in scores in the GroupMatches / KnockoutResults tabs
 * - OR set up a timed trigger (see bottom) to auto-fetch from a football API
 */

/**
 * Creates the "GroupMatches" and "KnockoutResults" tabs if they don't already exist,
 * with the correct headers and formatting.
 */
function setupSheetTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ---- GroupMatches tab ----
  let gm = ss.getSheetByName('GroupMatches');
  if (!gm) {
    gm = ss.insertSheet('GroupMatches');
    Logger.log('Created "GroupMatches" tab');
  }
  // Set headers
  const gmHeaders = ['Group', 'Team1', 'Score1', 'Team2', 'Score2', 'Status'];
  gm.getRange(1, 1, 1, gmHeaders.length).setValues([gmHeaders]);
  gm.getRange(1, 1, 1, gmHeaders.length)
    .setFontWeight('bold')
    .setBackground('#005a9c')
    .setFontColor('#ffffff');
  // Freeze header row
  gm.setFrozenRows(1);
  // Set column widths
  gm.setColumnWidth(1, 70);   // Group
  gm.setColumnWidth(2, 200);  // Team1
  gm.setColumnWidth(3, 70);   // Score1
  gm.setColumnWidth(4, 200);  // Team2
  gm.setColumnWidth(5, 70);   // Score2
  gm.setColumnWidth(6, 100);  // Status
  // Add data validation for Status column
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Scheduled', 'Live', 'Finished'], true)
    .setAllowInvalid(false)
    .build();
  gm.getRange('F2:F200').setDataValidation(statusRule);

  // ---- KnockoutResults tab ----
  let kr = ss.getSheetByName('KnockoutResults');
  if (!kr) {
    kr = ss.insertSheet('KnockoutResults');
    Logger.log('Created "KnockoutResults" tab');
  }
  const krHeaders = ['StageKey', 'TeamName'];
  kr.getRange(1, 1, 1, krHeaders.length).setValues([krHeaders]);
  kr.getRange(1, 1, 1, krHeaders.length)
    .setFontWeight('bold')
    .setBackground('#005a9c')
    .setFontColor('#ffffff');
  kr.setFrozenRows(1);
  kr.setColumnWidth(1, 120);
  kr.setColumnWidth(2, 200);

  // Pre-fill stage keys so user just needs to type team names
  const stageKeys = [
    ['m1', ''], ['m2', ''], ['m3', ''], ['m4', ''],
    ['m5', ''], ['m6', ''], ['m7', ''], ['m8', ''],
    ['q1', ''], ['q2', ''], ['q3', ''], ['q4', ''],
    ['sf1', ''], ['sf2', ''],
    ['winner', '']
  ];
  kr.getRange(2, 1, stageKeys.length, 2).setValues(stageKeys);
  // Add notes explaining each stage key
  kr.getRange('A2').setNote('Round of 16 — Match 49 winner');
  kr.getRange('A3').setNote('Round of 16 — Match 50 winner');
  kr.getRange('A4').setNote('Round of 16 — Match 51 winner');
  kr.getRange('A5').setNote('Round of 16 — Match 52 winner');
  kr.getRange('A6').setNote('Round of 16 — Match 53 winner');
  kr.getRange('A7').setNote('Round of 16 — Match 54 winner');
  kr.getRange('A8').setNote('Round of 16 — Match 55 winner');
  kr.getRange('A9').setNote('Round of 16 — Match 56 winner');
  kr.getRange('A10').setNote('Quarterfinal 1 winner (m1 vs m2)');
  kr.getRange('A11').setNote('Quarterfinal 2 winner (m3 vs m4)');
  kr.getRange('A12').setNote('Quarterfinal 3 winner (m5 vs m6)');
  kr.getRange('A13').setNote('Quarterfinal 4 winner (m7 vs m8)');
  kr.getRange('A14').setNote('Semifinal 1 winner (q1 vs q2)');
  kr.getRange('A15').setNote('Semifinal 2 winner (q3 vs q4)');
  kr.getRange('A16').setNote('FIFA World Cup 2026 Champion');

  SpreadsheetApp.getUi().alert(
    '✅ Setup Complete!\n\n' +
    '• "GroupMatches" tab — fill in group stage scores\n' +
    '• "KnockoutResults" tab — fill in knockout round winners\n\n' +
    'The web page auto-refreshes from this sheet every 3 minutes.'
  );
}

/**
 * Adds a custom menu to the spreadsheet for easy access
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 World Cup')
    .addItem('Setup Tabs', 'setupSheetTabs')
    .addToUi();
}

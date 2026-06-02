/**
 * ===================================================================
 * Bruker World Cup 2026 — Google Apps Script (Fully Automated)
 * ===================================================================
 *
 * HOW TO INSTALL (one-time, ~2 minutes):
 *
 * 1. Get a FREE API key:
 *    → Go to https://www.football-data.org/client/register
 *    → Sign up with your email — you'll get an API key instantly
 *
 * 2. Open the Google Sheet:
 *    → https://docs.google.com/spreadsheets/d/1zVoXUcy-17ctpic7e75rXhJ23XhLG2YKi1PbDtGeJME/edit
 *
 * 3. Go to Extensions → Apps Script
 *
 * 4. Delete any existing code and paste this entire file
 *
 * 5. Click Save (Ctrl+S)
 *
 * 6. From the function dropdown at the top, select "setup" → click ▶ Run
 *    → Allow permissions when prompted
 *    → Paste your API key when asked
 *    → Done! Scores auto-update every 10 minutes.
 *
 * That's it — fully automated, no manual entry needed.
 * ===================================================================
 */

// ===== CONFIGURATION =====
const COMPETITION = 'WC';   // football-data.org code for FIFA World Cup
const UPDATE_INTERVAL = 10; // minutes between auto-updates

// API team names → our team names (only where they differ)
const TEAM_NAME_MAP = {
  'Korea Republic':          'South Korea',
  "Côte d'Ivoire":           'Ivory Coast',
  'Cote d\'Ivoire':          'Ivory Coast',
  'Türkiye':                 'Turkey',
  'Czechia':                 'Czech Republic',
  'Congo DR':                'DR Congo',
  'Bosnia-Herzegovina':      'Bosnia and Herzegovina',
  'Cape Verde Islands':      'Cape Verde',
  'Republic of Ireland':     'Ireland',
  'Curacao':                 'Curaçao',
};

function normalizeName(apiName) {
  return TEAM_NAME_MAP[apiName] || apiName;
}

// ===== ONE-TIME SETUP =====
/**
 * Run this once. It:
 *  1. Asks for your football-data.org API key
 *  2. Creates the GroupMatches & KnockoutResults tabs
 *  3. Sets a 10-minute auto-update trigger
 *  4. Does an immediate first fetch
 */
function setup() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    '🏆 World Cup Auto-Updater Setup',
    'Paste your football-data.org API key\n(free at football-data.org/client/register):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const key = resp.getResponseText().trim();
  if (!key) { ui.alert('❌ No API key provided. Setup cancelled.'); return; }

  // Store API key securely in Script Properties
  PropertiesService.getScriptProperties().setProperty('API_KEY', key);

  // Create sheet tabs
  createSheetTabs_();

  // Install timed trigger
  installTrigger_();

  // Do immediate first fetch
  updateScores();

  ui.alert(
    '✅ Setup Complete!\n\n' +
    '• GroupMatches tab — auto-populated with group stage fixtures\n' +
    '• KnockoutResults tab — auto-populated as knockout rounds are played\n' +
    '• Auto-updates every ' + UPDATE_INTERVAL + ' minutes\n\n' +
    'No manual work needed — everything is automatic!'
  );
}

// ===== MENU =====
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏆 World Cup')
    .addItem('▶ Initial Setup', 'setup')
    .addItem('🔄 Update Now', 'updateScores')
    .addItem('⏱ Check Trigger Status', 'showTriggerStatus')
    .addToUi();
}

// ===== SHEET TAB CREATION =====
function createSheetTabs_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── GroupMatches ──
  let gm = ss.getSheetByName('GroupMatches');
  if (!gm) gm = ss.insertSheet('GroupMatches');
  const gmH = ['Group', 'Team1', 'Score1', 'Team2', 'Score2', 'Status'];
  gm.getRange(1, 1, 1, gmH.length).setValues([gmH])
    .setFontWeight('bold').setBackground('#005a9c').setFontColor('#fff');
  gm.setFrozenRows(1);
  [70, 200, 70, 200, 70, 100].forEach((w, i) => gm.setColumnWidth(i + 1, w));

  // ── KnockoutResults ──
  let kr = ss.getSheetByName('KnockoutResults');
  if (!kr) kr = ss.insertSheet('KnockoutResults');
  const krH = ['StageKey', 'TeamName'];
  kr.getRange(1, 1, 1, krH.length).setValues([krH])
    .setFontWeight('bold').setBackground('#005a9c').setFontColor('#fff');
  kr.setFrozenRows(1);
  kr.setColumnWidth(1, 120);
  kr.setColumnWidth(2, 200);

  // Pre-fill stage keys
  const keys = [
    'm1','m2','m3','m4','m5','m6','m7','m8',
    'q1','q2','q3','q4',
    'sf1','sf2',
    'winner'
  ];
  kr.getRange(2, 1, keys.length, 2).setValues(keys.map(k => [k, '']));
}

// ===== TRIGGER MANAGEMENT =====
function installTrigger_() {
  // Remove old triggers for this function
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateScores') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateScores')
    .timeBased()
    .everyMinutes(UPDATE_INTERVAL)
    .create();
}

function showTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers().filter(
    t => t.getHandlerFunction() === 'updateScores'
  );
  const ui = SpreadsheetApp.getUi();
  if (triggers.length === 0) {
    ui.alert('⚠️ No auto-update trigger found.\nRun "Initial Setup" to create one.');
  } else {
    ui.alert('✅ Auto-update is active.\nRunning every ' + UPDATE_INTERVAL + ' minutes.');
  }
}

// ===== MAIN UPDATE FUNCTION =====
/**
 * Fetches all World Cup matches from football-data.org and
 * writes group scores + knockout results to the sheet.
 * Called automatically every 10 minutes by the trigger.
 */
function updateScores() {
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key) { Logger.log('No API key — run setup() first'); return; }

  const url = 'https://api.football-data.org/v4/competitions/' + COMPETITION + '/matches';
  const response = UrlFetchApp.fetch(url, {
    headers: { 'X-Auth-Token': key },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  if (code === 429) { Logger.log('Rate limited — will retry next cycle'); return; }
  if (code !== 200) { Logger.log('API error ' + code + ': ' + response.getContentText().substring(0, 300)); return; }

  const data = JSON.parse(response.getContentText());
  const matches = data.matches || [];
  if (matches.length === 0) { Logger.log('No matches returned'); return; }

  writeGroupMatches_(matches);
  writeKnockoutResults_(matches);

  Logger.log('Updated ' + matches.length + ' matches at ' + new Date().toISOString());
}

// ===== GROUP MATCHES =====
function writeGroupMatches_(matches) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('GroupMatches');
  if (!sheet) { sheet = ss.insertSheet('GroupMatches'); }

  const groupMatches = matches.filter(m => m.stage === 'GROUP_STAGE');
  if (groupMatches.length === 0) return;

  // Sort by group then by match date
  groupMatches.sort((a, b) => {
    const ga = (a.group || '').replace('GROUP_', '');
    const gb = (b.group || '').replace('GROUP_', '');
    if (ga !== gb) return ga.localeCompare(gb);
    return new Date(a.utcDate) - new Date(b.utcDate);
  });

  const rows = groupMatches.map(m => {
    const group = (m.group || '').replace('GROUP_', '');
    const team1 = normalizeName(m.homeTeam.name || m.homeTeam.shortName || '');
    const team2 = normalizeName(m.awayTeam.name || m.awayTeam.shortName || '');

    let score1 = '', score2 = '', status = 'Scheduled';
    const ft = m.score && m.score.fullTime;

    if (m.status === 'FINISHED') {
      score1 = ft && ft.home != null ? ft.home : '';
      score2 = ft && ft.away != null ? ft.away : '';
      status = 'Finished';
    } else if (m.status === 'IN_PLAY' || m.status === 'PAUSED') {
      score1 = ft && ft.home != null ? ft.home : 0;
      score2 = ft && ft.away != null ? ft.away : 0;
      status = 'Live';
    }
    // SCHEDULED, TIMED, POSTPONED → leave as Scheduled with no scores

    return [group, team1, score1, team2, score2, status];
  });

  // Write everything (overwrite below header)
  if (rows.length > 0) {
    // Clear old data
    const lastRow = Math.max(sheet.getLastRow(), 1);
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 6).clearContent();
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
}

// ===== KNOCKOUT RESULTS =====
function writeKnockoutResults_(matches) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('KnockoutResults');
  if (!sheet) return;

  // Separate knockout matches by stage
  const r16 = matches.filter(m => m.stage === 'LAST_16').sort((a, b) => new Date(a.utcDate) - new Date(b.utcDate));
  const qf  = matches.filter(m => m.stage === 'QUARTER_FINALS').sort((a, b) => new Date(a.utcDate) - new Date(b.utcDate));
  const sf  = matches.filter(m => m.stage === 'SEMI_FINALS').sort((a, b) => new Date(a.utcDate) - new Date(b.utcDate));
  const fin = matches.filter(m => m.stage === 'FINAL');

  function getWinner(m) {
    if (!m || m.status !== 'FINISHED') return '';
    const w = m.score && m.score.winner;
    if (w === 'HOME_TEAM') return normalizeName(m.homeTeam.name || m.homeTeam.shortName || '');
    if (w === 'AWAY_TEAM') return normalizeName(m.awayTeam.name || m.awayTeam.shortName || '');
    // Penalty shootout — check penalties score
    const pen = m.score && m.score.penalties;
    if (pen && pen.home != null && pen.away != null) {
      return pen.home > pen.away
        ? normalizeName(m.homeTeam.name || '')
        : normalizeName(m.awayTeam.name || '');
    }
    return '';
  }

  // Map to our stage keys (m1-m8, q1-q4, sf1-sf2, winner)
  const results = {
    m1: getWinner(r16[0]), m2: getWinner(r16[1]),
    m3: getWinner(r16[2]), m4: getWinner(r16[3]),
    m5: getWinner(r16[4]), m6: getWinner(r16[5]),
    m7: getWinner(r16[6]), m8: getWinner(r16[7]),
    q1: getWinner(qf[0]),  q2: getWinner(qf[1]),
    q3: getWinner(qf[2]),  q4: getWinner(qf[3]),
    sf1: getWinner(sf[0]), sf2: getWinner(sf[1]),
    winner: getWinner(fin[0])
  };

  // Write to sheet (stage keys are already in column A from setup)
  const keys = ['m1','m2','m3','m4','m5','m6','m7','m8','q1','q2','q3','q4','sf1','sf2','winner'];
  const rows = keys.map(k => [k, results[k] || '']);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

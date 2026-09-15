/***** Daily Internship Quota Tracker – Ultra-Fast Version 
 * 
 * This Google Apps Script tracks daily internship application quotas for a team.
 * Each person gets their own sheet to log applications, with automatic quota tracking,
 * penalties for missed quotas, and consolidated reporting.
 * 
 * Features:
 * - Per-person logging sheets with auto-date stamping
 * - Daily quota tracking (default: 10 applications/day)
 * - Master log aggregating all applications
 * - Daily summary dashboard showing quota completion
 * - Penalties ledger for missed quotas
 * - Conditional formatting for visual feedback
 ****************************************************************/

// Column headers for each sheet type
const HEADERS = {
  TEMPLATE: ['Date', 'Company', 'Position', 'Status', 'CountToday', 'Completed', 'Notes', 'Time', 'DayKey'], // Per-person sheet structure
  MASTER: ['Date', 'Applicant', 'Company', 'Position', 'Status', 'CountToday', 'Completed', 'Notes', 'Time'], // Consolidated log
  DAILY: ['Date', 'Applicant', 'AppliedCount', 'QuotaMet', 'Missed', 'PenaltyEligible'], // Daily summary dashboard
  PENS: ['Date', 'Applicant', 'Missed', 'Owed', 'User', 'MissedDays', 'TotalOwed'] // Penalties tracking
};

/**
 * Creates the custom "Tracker" menu when the spreadsheet opens
 * Provides easy access to main functions without needing to run scripts manually
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tracker')
    .addItem('Setup', 'setup')                    // Initial setup - creates all sheets and formulas
    .addItem('Sync Sheets', 'syncPersonSheets')   // Create/update individual person sheets
    .addItem('Rebuild Master', 'rebuildMaster')   // Refresh the master log from all person sheets
    .addItem('Install Nightly', 'installNightly') // Set up automatic nightly master rebuild
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();

  // Batch create all sheets first
  const sheets = ['Settings', '_TEMPLATE', 'Master', 'Daily Summary', 'Penalties'];
  const sheetRefs = {};
  sheets.forEach(name => {
    sheetRefs[name] = ss.getSheetByName(name) || ss.insertSheet(name);
  });

  // Settings - column-oriented table with per-user overrides and a 'default' row
  sheetRefs.Settings.clear().getRange('A1:D10').setValues([
    ['Users', 'QuotaPerDay', 'PenaltyPerMiss', 'PenaltyStartDate'],
    ['default', 10, 0.5, ''],
    ['Porter', '', '', ''],
    ['Maia', '', '', ''],
    ['Kainoa', '', '', ''],
    ['Caren', '', '', ''],
    ['Armaan', '', '', ''],
    ['Steph', '', '', ''],
    ['', '', '', ''],
    ['', '', '', '']
  ]);
  sheetRefs.Settings.getRange('D2:D').setNumberFormat('yyyy-mm-dd');
  // Auto-set PenaltyStartDate to today for any users without a start date
  try {
    const usersCol = sheetRefs.Settings.getRange('A2:A').getValues();
    const startDates = sheetRefs.Settings.getRange('D2:D').getValues();
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let changed = false;
    for (let i = 0; i < usersCol.length; i++) {
      const hasUser = usersCol[i][0] && String(usersCol[i][0]).trim() !== '';
      const hasDate = !!startDates[i][0];
      if (hasUser && !hasDate) {
        startDates[i][0] = today;
        changed = true;
      }
    }
    if (changed) {
      sheetRefs.Settings.getRange('D2:D').setValues(startDates);
    }
  } catch (err) { }
  // Style Settings header like others and increase height
  sheetRefs.Settings.getRange(1, 1, 1, 4)
    .setBackground('#b4cdcd')
    .setFontColor(null)
    .setFontWeight('bold')
    .setHorizontalAlignment('left');
  sheetRefs.Settings.setRowHeight(1, 108);

  // Template - minimal setup (hidden sheet)
  const tmpl = sheetRefs['_TEMPLATE'];
  tmpl.clear().getRange(1, 1, 1, 9).setValues([HEADERS.TEMPLATE]);
  tmpl.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  // Set formulas across E..I: [CountToday, Completed, Notes(blank), Time(blank), DayKey]
  tmpl.getRange('E2:I2').setFormulas([[
    `=ARRAYFORMULA(IF(A2:A="",,IF(D2:D="","",COUNTIFS(A$2:A,A2:A,ROW(A$2:A),"<="&ROW(A2:A),D$2:D,"<>"))))`,
    `=ARRAYFORMULA(IF(A2:A="",,IF(E2:E=IFERROR(VLOOKUP("default",Settings!A2:D,2,false),10),"met "&TEXT(A2:A,"mmm d")&" quota",IF(E2:E>IFERROR(VLOOKUP("default",Settings!A2:D,2,false),10),"Extra",""))))`,
    "",
    "",
    `=ARRAYFORMULA(IF(A2:A="","",A2:A))`
  ]]);
  tmpl.hideSheet();

  // Master
  sheetRefs.Master.clear().getRange(1, 1, 1, 9).setValues([HEADERS.MASTER]);
  sheetRefs.Master.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sheetRefs.Master.getRange(1, 1, 1, 9).setBackground('#b4cdcd').setFontColor(null).setFontWeight('bold').setHorizontalAlignment('left');
  sheetRefs.Master.setRowHeight(1, 108);

  // Daily Summary - batch formulas (includes PenaltyEligible)
  const daily = sheetRefs['Daily Summary'];
  daily.clear().getRange(1, 1, 1, 6).setValues([HEADERS.DAILY]);
  daily.getRange(1, 1, 1, 6).setBackground('#b4cdcd').setFontColor(null).setFontWeight('bold').setHorizontalAlignment('left');
  daily.setRowHeight(1, 108);
  daily.getRange('A2:F2').setFormulas([[
    `=ARRAYFORMULA(IF(B2:B="","",DATE(1899,12,30)+VALUE(INDEX(SPLIT(B2:B,"|"),,2))))`,
    `=ARRAYFORMULA(FLATTEN(IF(UNIQUE(FILTER(Master!A2:A,Master!A2:A<>""))="",,TRANSPOSE(FILTER(Settings!A2:A,(Settings!A2:A<>"")*(Settings!A2:A<>"default")))&"|"&VALUE(UNIQUE(FILTER(Master!A2:A,Master!A2:A<>""))))))`,
    `=ARRAYFORMULA(IF(B2:B="","",BYROW(B2:B,LAMBDA(k,IF(k="","",LET(u,INDEX(SPLIT(k,"|"),1,1),d,INDEX(A2:A,ROW(k)-ROW($B$2)+1),COUNTIFS(Master!B:B,u,Master!A:A,d)))))))`,
    `=ARRAYFORMULA(IF(B2:B="","",BYROW(B2:B,LAMBDA(k,IF(k="","",LET(u,INDEX(SPLIT(k,"|"),1,1),q,IFERROR(VLOOKUP(u,Settings!A2:D,2,false),VLOOKUP("default",Settings!A2:D,2,false)),cnt,INDEX(C2:C,ROW(k)-ROW($B$2)+1),cnt>=q))))))`,
    `=ARRAYFORMULA(IF(B2:B="","",IF(D2:D,"Met","Missed")))`,
    `=ARRAYFORMULA(IF(B2:B="","",BYROW(B2:B,LAMBDA(k,IF(k="","",LET(u,INDEX(SPLIT(k,"|"),1,1),d,INDEX(A2:A,ROW(k)-ROW($B$2)+1),start,IFERROR(VLOOKUP(u,Settings!A2:D,4,false),TODAY()),d>=start))))))`
  ]]);

  // Penalties - batch formulas (filters to penalty-eligible missed days)
  const pens = sheetRefs.Penalties;
  pens.clear().getRange(1, 1, 1, 7).setValues([HEADERS.PENS]);
  pens.getRange(1, 1, 1, 7).setBackground('#b4cdcd').setFontColor(null).setFontWeight('bold').setHorizontalAlignment('left');
  pens.setRowHeight(1, 108);
  pens.getRange('A2:G2').setFormulas([[
    // A2: Date - Get dates where quotas were missed and penalty-eligible
    `=FILTER('Daily Summary'!A2:A,('Daily Summary'!E2:E="Missed")*('Daily Summary'!F2:F=TRUE))`,

    // B2: Applicant - Get users who missed penalty-eligible quotas
    `=FILTER(INDEX(SPLIT('Daily Summary'!B2:B,"|"),,1),('Daily Summary'!E2:E="Missed")*('Daily Summary'!F2:F=TRUE))`,

    // C2: Missed - Mark all as "Yes" (they all missed)
    `=ARRAYFORMULA(IF(A2:A="","", "Yes"))`,

    // D2: Owed - Get penalty amount per miss for each user (or default)
    `=ARRAYFORMULA(IF(B2:B="","", BYROW(B2:B, LAMBDA(u, IF(u="","", IFERROR(VLOOKUP(u,Settings!A2:D,3,false),VLOOKUP("default",Settings!A2:D,3,false)))))))`,

    // E2: User - Unique list of users who have penalties
    `=UNIQUE(FILTER(B2:B,B2:B<>""))`,

    // F2: MissedDays - Count total missed days per user
    `=ARRAYFORMULA(IF(E2:E="",,COUNTIF(B2:B,E2:E)))`,

    // G2: TotalOwed - Calculate total amount owed per user
    `=ARRAYFORMULA(IF(E2:E="",,F2:F*IFERROR(VLOOKUP(E2:E,Settings!A2:D,3,false),VLOOKUP("default",Settings!A2:D,3,false))))`
  ]]);

  syncPersonSheets();
  rebuildMaster();
}

function syncPersonSheets() {
  const ss = SpreadsheetApp.getActive();
  // Users from Settings column A, excluding header and 'default'
  const users = ss.getSheetByName('Settings').getRange('A2:A').getValues().flat().filter(u => u && u !== 'default');
  const existing = new Set(ss.getSheets().map(s => s.getName()));
  const newSheets = users.filter(user => !existing.has(user)).map(user => ss.insertSheet(user));

  // Create shared formatting objects once (kept for fallback)
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['applied', 'interviewing', 'Rejected :(', 'OFFERRRR!!!'], true)
    .setAllowInvalid(false)
    .setHelpText('Select application status')
    .build();
  const timeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Sp 26', 'Sum 26'], true)
    .setAllowInvalid(false)
    .setHelpText('Select time')
    .build();

  // Note: Per-sheet formulas depend on user name; will be constructed inside the loop

  // Batch setup all new sheets
  newSheets.forEach(sh => {
    const quotaExpr = `VLOOKUP("default",Settings!A2:D,2,false)`;
    // Single batch operations
    sh.getRange(1, 1, 1, 9).setValues([HEADERS.TEMPLATE]);
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
    sh.getRange('E:E').setNumberFormat('0');
    // Apply formulas and blanks across E..I to ensure DayKey in I
    sh.getRange('E2:I2').setFormulas([[
      `=ARRAYFORMULA(IF(A2:A="",,IF(D2:D="","",COUNTIFS(A$2:A,A2:A,ROW(A$2:A),"<="&ROW(A2:A),D$2:D,"<>"))))`,
      `=ARRAYFORMULA(IF(A2:A="",,IF(E2:E=${quotaExpr},"met "&TEXT(A2:A,"mmm d")&" quota",IF(E2:E>${quotaExpr},"Extra",""))))`,
      "",
      "",
      `=ARRAYFORMULA(IF(A2:A="","",A2:A))`
    ]]);
    // Apply basic validation for safety (fallback if Advanced service unavailable)
    sh.getRange('D2:D').setDataValidation(validation);
    sh.getRange('H2:H').setDataValidation(timeValidation);

    // Upgrade to chip-style dropdowns using Advanced Sheets API
    try {
      const sheetId = sh.getSheetId();
      const spreadsheetId = ss.getId();
      Sheets.Spreadsheets.batchUpdate({
        requests: [
          {
            setDataValidation: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1, // row 2
                startColumnIndex: 3, // column D (0-based)
                endColumnIndex: 4
              },
              rule: {
                condition: {
                  type: 'ONE_OF_LIST',
                  values: [
                    { userEnteredValue: 'applied' },
                    { userEnteredValue: 'interviewing' },
                    { userEnteredValue: 'Rejected :(' },
                    { userEnteredValue: 'OFFERRRR!!!' }
                  ]
                },
                strict: true,
                showCustomUi: true
              }
            }
          }
          ,
          {
            setDataValidation: {
              range: {
                sheetId: sheetId,
                startRowIndex: 1, // row 2
                startColumnIndex: 7, // column H (Time)
                endColumnIndex: 8
              },
              rule: {
                condition: {
                  type: 'ONE_OF_LIST',
                  values: [
                    { userEnteredValue: 'Sp 26' },
                    { userEnteredValue: 'Sum 26' }
                  ]
                },
                strict: true,
                showCustomUi: true
              }
            }
          }
        ]
      }, spreadsheetId);
    } catch (err) {
      // If Advanced service not enabled, the basic validation above remains in effect
    }

    // Batch column widths
    [100, 120, 200, 100, 90, 150, 150, 120, 100].forEach((width, i) => sh.setColumnWidth(i + 1, width));

    // Header formatting (bold text with background)
    sh.getRange(1, 1, 1, 9).setBackground('#b4cdcd').setFontColor(null).setFontWeight('bold').setHorizontalAlignment('left');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 36);
    sh.getRange('A2:I').setVerticalAlignment('middle');
    sh.getRange('B2:C').setWrap(true);

    // Conditional formatting - keep only quota-based rules (remove Status coloring)
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange('F2:F')]).whenFormulaSatisfied(`=$F2="Extra"`).setBackground('#D8BFD8').build(),
      SpreadsheetApp.newConditionalFormatRule().setRanges([sh.getRange('A2:I')]).whenFormulaSatisfied(`=$F2="met "&TEXT($A2,"mmm d")&" quota"`).setBackground('#E2F0D9').build()
    ]);

    // Remove heavy outer borders around the sheet data
    // (No explicit borders applied)
  });
}



function onEdit(e) {
  const sheet = e.range.getSheet();
  const name = sheet.getName();

  // Fast exit for system sheets
  if (name === 'Settings' || name === '_TEMPLATE' || name === 'Master' || name === 'Daily Summary' || name === 'Penalties') return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Only trigger on Status column (D=4) edits in data rows (2+)
  if (row < 2 || col !== 4) return;

  // Only stamp date if Status has a value and date is empty
  const statusCell = sheet.getRange(row, 4);
  const dateCell = sheet.getRange(row, 1);

  if (statusCell.getValue() && !dateCell.getValue()) {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    dateCell.setValue(today);
  }

  // Insert a red "Missed [Date] quota" row at date boundaries if the previous date did not meet quota
  try {
    if (row >= 3) {
      const currentDate = sheet.getRange(row, 1).getValue();
      const prevDate = row > 2 ? sheet.getRange(row - 1, 1).getValue() : '';
      if (currentDate && prevDate && currentDate.getTime && prevDate.getTime && currentDate.getTime() !== prevDate.getTime()) {
        // We crossed into a new date; check if previous date met quota
        const ssActive = SpreadsheetApp.getActive();
        const quota = ssActive.getSheetByName('Settings').getRange('B1').getValue();
        // Only scan up to the previous row for performance
        const prevDateCount = sheet.getRange(2, 1, row - 2, 5).getValues()
          .filter(r => r[0] && r[0].getTime && r[0].getTime() === prevDate.getTime() && r[3]) // D not empty
          .length;
        if (prevDateCount < quota) {
          // Insert a row at the boundary (before the current row), write message in Completed col (F), and color it red
          sheet.insertRowsBefore(row, 1);
          const missedRowRange = sheet.getRange(row, 1, 1, 9);
          missedRowRange.clearContent();
          const tz = ssActive.getSpreadsheetTimeZone ? ssActive.getSpreadsheetTimeZone() : 'America/Los_Angeles';
          const prevDateStr = Utilities.formatDate(prevDate, tz, 'MMM d');
          sheet.getRange(row, 6).setValue(`Missed ${prevDateStr} quota`);
          missedRowRange.setBackground('#F8D7DA').setFontColor('#721c24');
        }
      }
    }
  } catch (err) {
    // Non-fatal; skip if anything goes wrong
  }
}

function rebuildMaster() {
  const ss = SpreadsheetApp.getActive();
  const master = ss.getSheetByName('Master');
  const exclude = new Set(['Settings', '_TEMPLATE', 'Master', 'Daily Summary', 'Penalties']);

  // Get person sheets with data in one pass
  const personSheets = ss.getSheets().filter(s => !exclude.has(s.getName()) && s.getLastRow() >= 2);

  // Pre-allocate array with estimated size
  const out = [HEADERS.MASTER];

  // Batch read all data at once
  const allData = personSheets.map(s => ({
    name: s.getName(),
    data: s.getRange(2, 1, s.getLastRow() - 1, 9).getValues()
  }));

  // Process in memory (fastest)
  allData.forEach(({ name, data }) => {
    data.forEach(row => {
      if (row[0] || row[1] || row[2] || row[3] || row[4]) {
        out.push([row[0], name, row[1], row[2], row[3], row[5], row[6], row[7], row[8]]);
      }
    });
  });

  // Single write operation
  master.clear();
  if (out.length > 1) master.getRange(1, 1, out.length, 9).setValues(out);
  else master.getRange(1, 1, 1, 9).setValues([HEADERS.MASTER]);
}

function installNightly() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'rebuildMaster') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rebuildMaster').timeBased().atHour(23).nearMinute(59).everyDays(1).create();
}

/**
 * Debug function to check penalty calculations
 * Run this to verify penalties are working correctly
 */
function debugPenalties() {
  const ss = SpreadsheetApp.getActive();
  const daily = ss.getSheetByName('Daily Summary');
  const penalties = ss.getSheetByName('Penalties');
  const settings = ss.getSheetByName('Settings');

  console.log('=== PENALTY DEBUG ===');

  // Check Settings structure
  const settingsData = settings.getRange('A1:D10').getValues();
  console.log('Settings data:', settingsData);

  // Check Daily Summary data
  const dailyData = daily.getRange('A1:F10').getValues();
  console.log('Daily Summary data:', dailyData);

  // Check Penalties data
  const penaltiesData = penalties.getRange('A1:G10').getValues();
  console.log('Penalties data:', penaltiesData);

  // Check for missed days that should have penalties
  const missedDays = dailyData.filter((row, i) => i > 0 && row[4] === 'Missed' && row[5] === true);
  console.log('Missed penalty-eligible days:', missedDays);

  // Calculate expected penalties
  const defaultPenalty = settingsData.find(row => row[0] === 'default')?.[2] || 0.5;
  console.log('Default penalty per miss:', defaultPenalty);

  SpreadsheetApp.getActiveSpreadsheet().toast('Check console for penalty debug info', 'Debug Complete');
}
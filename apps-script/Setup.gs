/**
 * Run once after uploading the .xlsx to Google Drive and opening it as a Google Sheet:
 * Louise menu > Set up this Sheet (or run setupForSheets from the script editor).
 * Safe to run again; it rewrites the same things each time.
 */
function setupForSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  const N = CONFIG.LOG_ROWS + 1;

  // 1. Food Log: Excel's spill formulas become ARRAYFORMULA in Sheets
  const fl = ss.getSheetByName(CONFIG.FOOD);
  [FL.AGE, FL.ALLERGEN, FL.IRON, FL.GROUP].forEach(c => fl.getRange(3, c, N - 2, 1).clearContent());
  fl.getRange(2, FL.AGE).setFormula(
    `=ARRAYFORMULA(IF((A2:A${N}="")+(B2:B${N}="")>0,"",DATEDIF(DOB,A2:A${N},"m")))`);
  fl.getRange(2, FL.ALLERGEN).setFormula(
    `=ARRAYFORMULA(IF(B2:B${N}="","",IFERROR(VLOOKUP(B2:B${N},'Food Guide'!A$1:I$1000,9,FALSE)&"","")))`);
  fl.getRange(2, FL.IRON).setFormula(
    `=ARRAYFORMULA(IF(B2:B${N}="","",IF((IFERROR(VLOOKUP(B2:B${N},'Food Guide'!A$1:B$1000,2,FALSE),"")="Iron protein")` +
    `+(IFERROR(VLOOKUP(B2:B${N},'Food Guide'!A$1:B$1000,2,FALSE),"")="Fish")+(B2:B${N}="Oatmeal")+(B2:B${N}="Barley")` +
    `+(B2:B${N}="Quinoa")>0,"Iron","")))`);
  fl.getRange(2, FL.GROUP).setFormula(
    `=ARRAYFORMULA(IF(B2:B${N}="","",IFERROR(VLOOKUP(B2:B${N},'Food Guide'!A$1:B$1000,2,FALSE)&"","")))`);

  // 2. Allergens: "First tried" without Excel's LET
  const al = ss.getSheetByName('Allergens');
  for (let r = 2; r <= 10; r++) {
    al.getRange(r, 4).setFormula(
      `=IF(COUNTIF('Food Log'!E:E,B${r})=0,"",IF(MINIFS('Food Log'!A:A,'Food Log'!E:E,B${r})=0,"Before log",` +
      `MINIFS('Food Log'!A:A,'Food Log'!E:E,B${r})))`).setNumberFormat('mmm d, yyyy');
  }

  // 3. Monthly Progress: LOOKUP(2,1/(...)) needs array context in Sheets
  const mp = ss.getSheetByName('Monthly Progress');
  const growth = mp.getRange(5, 29, 55, 2);
  growth.setFormulas(growth.getFormulas().map(row => row.map(f =>
    f && f.indexOf('ARRAYFORMULA') < 0 ? '=ARRAYFORMULA(' + f.replace(/^=/, '') + ')' : f)));

  // 4. Monthly Progress colours: Sheets rules cannot point at another tab without INDIRECT
  rebuildProgressColours_(mp);

  // 5. Change Log tab and IDs for any row typed before setup
  changeLog_('setup', 'all', '', 'setupForSheets ran');
  stampMissingIds_(ss.getSheetByName(CONFIG.FOOD), FL.FOOD, FL.ID, 'F');
  stampMissingIds_(ss.getSheetByName(CONFIG.DRINKS), DL.DRINK, DL.ID, 'D');

  // 6. Drive folders and the nightly job
  folder_(CONFIG.BACKUP_FOLDER);
  folder_(CONFIG.ARCHIVE_FOLDER);
  installTriggers();

  SpreadsheetApp.getUi().alert('Louise is set up',
    'Formulas converted, colours rebuilt, IDs stamped, and a nightly backup scheduled for about 2am.\n\n' +
    'Next: Deploy > New deployment > Web app (see README), then open the link on both phones and add it to the home screen.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function stampMissingIds_(sh, keyCol, idCol, prefix) {
  const last = lastRow_(sh, keyCol);
  if (last < 2) return;
  const keys = sh.getRange(2, keyCol, last - 1, 1).getValues();
  const ids = sh.getRange(2, idCol, last - 1, 1).getValues();
  sh.getRange(2, idCol, last - 1, 1).setValues(ids.map((r, i) => [r[0] || (keys[i][0] ? newId_(prefix) : '')]));
}

function rebuildProgressColours_(mp) {
  const first = 5, last = 59;
  const T = col => `INDEX(INDIRECT("'Age Targets'!$${col}$5:$${col}$9"),MATCH($A${first},INDIRECT("'Age Targets'!$B$5:$B$9"),1))`;
  const rule = (a1, formula, colour) => SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(formula).setBackground(colour).setRanges([mp.getRange(a1)]).build();
  const AMBER = '#F9DCA4', RED = '#F4C7C3', GREEN = '#CAEAC7';
  const rules = [];
  [['Q', 'F'], ['R', 'G'], ['S', 'H'], ['T', 'I'], ['U', 'J']].forEach(([c, t]) =>
    rules.push(rule(`${c}${first}:${c}${last}`, `=AND(ISNUMBER(${c}${first}),$A${first}>=12,${c}${first}<${T(t)})`, AMBER)));
  rules.push(rule(`W${first}:W${last}`, `=AND(ISNUMBER(W${first}),W${first}<${T('P')})`, AMBER));
  rules.push(rule(`Z${first}:Z${last}`, `=AND(ISNUMBER(Z${first}),Z${first}>${T('U')})`, RED));
  rules.push(rule(`AA${first}:AA${last}`, `=AND(ISNUMBER(AA${first}),AA${first}>${T('V')})`, RED));
  rules.push(rule(`AB${first}:AB${last}`, `=AND(AB${first}<>"",AB${first}<>"On track")`, AMBER));
  rules.push(rule(`A${first}:AG${last}`, `=AND(TODAY()>=$C${first},TODAY()<=$D${first})`, GREEN));
  mp.setConditionalFormatRules(rules);
}

function folder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

/** Replaces any earlier Louise triggers with one nightly job. */
function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'nightlyJobs') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('nightlyJobs').timeBased().everyDays(1).atHour(2).create();
}

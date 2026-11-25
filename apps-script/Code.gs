/**
 * Louise food tracker: web app and logging API.
 *
 * Lives inside the Google Sheet converted from Louises_First_Foods_Log.xlsx
 * (Extensions > Apps Script). Everything runs on a free Google account.
 * See README.md in this folder for setup.
 */

const CONFIG = {
  BACKUP_FOLDER: 'Louise backups',
  ARCHIVE_FOLDER: 'Louise monthly archive',
  KEEP_NIGHTLY_DAYS: 60,          // nightly copies older than this go to the Drive trash
  REPORT_EMAILS: [],              // empty = the Sheet owner's address
  DAILY_AS_VALUES: false,         // true = Daily tab is filled by script (faster Sheet, see README)
  LOG_ROWS: 20000,                // matches the workbook's log capacity
  FOOD: 'Food Log',
  DRINKS: 'Drinks Log',
  CHANGES: 'Change Log',
};

// Food Log columns (1-based), as in the workbook
const FL = { DATE: 1, FOOD: 2, AGE: 3, TRY: 4, ALLERGEN: 5, IRON: 6, FORM: 7, AMOUNT: 8, LIKED: 9,
  REACTION: 10, TIME: 11, BRAND: 12, NOTES: 13, SERVINGS: 14, GROUP: 15, ID: 16 };
// Drinks Log columns
const DL = { DATE: 1, TIME: 2, DRINK: 3, OZ: 4, NOTES: 5, ID: 6 };

const FORMS = ['Purée', 'Mash', 'Gnaw / strip', 'Mixed into another food'];
const LIKES = ['Loved', 'Okay', 'Not yet'];
const REACTIONS = ['None', 'Mouth redness (acid)', 'Hives or rash', 'Vomiting', 'Diarrhea', 'Constipation',
  'Swelling or breathing (911)', 'Other'];
const DRINKS = ['Breast milk', 'Formula', 'Water', 'Whole milk', 'Low-fat or skim milk', 'Juice', 'Other'];
const IRON_GRAINS = ['Oatmeal', 'Barley', 'Quinoa'];

/* ---------------------------------------------------------------- web app */

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Louise')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Louise')
    .addItem('Set up this Sheet (run once)', 'setupForSheets')
    .addSeparator()
    .addItem('Back up now', 'nightlyBackup')
    .addItem("Build this month's archive now", 'monthlyArchive')
    .addItem('Email me the doctor report', 'emailReportNow')
    .addItem('Refresh Daily values', 'refreshDailyValues')
    .addToUi();
}

/* ---------------------------------------------------------------- helpers */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}
function tz_() { return ss_().getSpreadsheetTimeZone(); }
function key_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }
function dob_() { return new Date(ss_().getRangeByName('DOB').getValue()); }

/** Whole months between two dates, like Excel's DATEDIF(a, b, "m"). */
function monthsBetween_(a, b) {
  let m = (b.getFullYear() - a.getFullYear()) * 12 + (b.getMonth() - a.getMonth());
  if (b.getDate() < a.getDate()) m -= 1;
  return m;
}
/** Same date n months later, clamped to month end, like Excel's EDATE. */
function edate_(d, n) {
  const y = d.getFullYear(), m = d.getMonth() + n;
  const last = new Date(y, m + 1, 0).getDate();
  return new Date(y, m, Math.min(d.getDate(), last));
}
function newId_(prefix) { return prefix + '-' + Utilities.getUuid().slice(0, 8); }

/** Last row with something in the given column (array formulas make getLastRow unreliable). */
function lastRow_(sheet, col) {
  const n = Math.min(sheet.getMaxRows(), CONFIG.LOG_ROWS + 1) - 1;
  const vals = sheet.getRange(2, col, n, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) if (vals[i][0] !== '' && vals[i][0] !== null) return i + 2;
  return 1;
}

function foodGuide_() {
  const sh = ss_().getSheetByName('Food Guide');
  const last = lastRow_(sh, 1);
  if (last < 2) return {};
  const rows = sh.getRange(2, 1, last - 1, 10).getValues();
  const out = {};
  rows.forEach(r => {
    if (!r[0]) return;
    const group = String(r[1]);
    out[r[0]] = {
      name: String(r[0]), group, allergen: String(r[8] || ''), serve: String(r[6] || ''), insight: String(r[9] || ''),
      iron: group === 'Iron protein' || group === 'Fish' || IRON_GRAINS.indexOf(String(r[0])) >= 0,
    };
  });
  return out;
}

function changeLog_(action, sheetName, id, details, book) {
  const ss = book || ss_();
  let sh = ss.getSheetByName(CONFIG.CHANGES);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.CHANGES);
    sh.appendRow(['When', 'Who', 'Action', 'Sheet', 'ID', 'Details']);
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), Session.getActiveUser().getEmail() || 'unknown', action, sheetName, id || '', details || '']);
}

function withLock_(fn) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------------------------------------------------------- read API */

/** Everything the web app shows, in one call. */
function getDashboard() {
  const ss = ss_();
  const now = new Date();
  const today = key_(now);
  const dob = dob_();
  const age = monthsBetween_(dob, now);

  // Age band and targets
  const at = ss.getSheetByName('Age Targets').getRange(5, 1, 5, 23).getValues();
  let band = null;
  at.forEach(r => { if (r[1] !== '' && age >= Number(r[1])) band = r; });
  const targets = band ? {
    band: band[0], meals: band[3], snacks: band[4],
    waterMin: band[15], waterMax: band[16], formulaMin: band[17], formulaMax: band[18],
    cowMax: band[20], juiceMax: band[21], focus: band[22], cueLed: age < 12,
  } : { band: 'Under 6 months', cueLed: true };

  // Food Guide and tries so far
  const guide = foodGuide_();
  const food = ss.getSheetByName(CONFIG.FOOD);
  const fLast = lastRow_(food, FL.FOOD);
  const tries = {};
  if (fLast >= 2) food.getRange(2, FL.FOOD, fLast - 1, 1).getValues().forEach(r => { if (r[0]) tries[r[0]] = (tries[r[0]] || 0) + 1; });

  // Today's foods (scan the most recent 400 rows)
  const fFrom = Math.max(2, fLast - 400);
  const fRows = fLast >= 2 ? food.getRange(fFrom, 1, fLast - fFrom + 1, FL.ID).getValues() : [];
  const fDisp = fLast >= 2 ? food.getRange(fFrom, 1, fLast - fFrom + 1, FL.ID).getDisplayValues() : [];
  const todayFoods = [];
  fRows.forEach((r, i) => {
    if (!(r[0] instanceof Date) || key_(r[0]) !== today || !r[1]) return;
    todayFoods.push({
      id: r[FL.ID - 1], time: fDisp[i][FL.TIME - 1], food: r[1], tryNo: r[FL.TRY - 1],
      allergen: r[FL.ALLERGEN - 1], iron: r[FL.IRON - 1] === 'Iron', form: r[FL.FORM - 1],
      servings: r[FL.SERVINGS - 1] || 1, liked: r[FL.LIKED - 1], reaction: r[FL.REACTION - 1] || 'None',
    });
  });

  // Today's drinks
  const drinks = ss.getSheetByName(CONFIG.DRINKS);
  const dLast = lastRow_(drinks, DL.DRINK);
  const dFrom = Math.max(2, dLast - 400);
  const dRows = dLast >= 2 ? drinks.getRange(dFrom, 1, dLast - dFrom + 1, DL.ID).getValues() : [];
  const drinkToday = { water: 0, feeds: 0, formula: 0, cow: 0, juice: 0 };
  dRows.forEach(r => {
    if (!(r[0] instanceof Date) || key_(r[0]) !== today) return;
    const oz = Number(r[DL.OZ - 1]) || 0;
    if (r[2] === 'Water') drinkToday.water += oz;
    else if (r[2] === 'Breast milk') drinkToday.feeds += 1;
    else if (r[2] === 'Formula') drinkToday.formula += oz;
    else if (r[2] === 'Whole milk' || r[2] === 'Low-fat or skim milk') drinkToday.cow += oz;
    else if (r[2] === 'Juice') drinkToday.juice += oz;
  });

  // Allergen plan, straight from the Allergens tab
  const al = ss.getSheetByName('Allergens').getRange(2, 1, 9, 9).getDisplayValues().map(r => ({
    week: r[0], name: r[1], start: r[2], first: r[3], last7: Number(r[4]) || 0, target: Number(r[5]) || 0,
    status: r[6], keep: r[7], ifReacts: r[8],
  }));

  // This month and the months around it, from Monthly Progress
  const mp = ss.getSheetByName('Monthly Progress').getRange(4, 1, 56, 33).getDisplayValues();
  const head = mp[0];
  const rowObj = r => { const o = {}; head.forEach((h, i) => { o[h] = r[i]; }); return o; };
  const monthRows = mp.slice(1).map(rowObj);
  const month = monthRows.find(r => parseInt(r['Age'], 10) === age) || null;
  const months = monthRows.filter(r => { const a = parseInt(r['Age'], 10); return a >= 6 && a <= Math.max(12, age + 3); });

  // Day by day for this month, from the Daily tab
  const daily = ss.getSheetByName('Daily').getRange(2, 1, 1700, 16).getValues();
  const days = daily.filter(r => r[1] === age).map(r => ({
    date: r[0] instanceof Date ? key_(r[0]) : '', iron: r[3] === 1, newFoods: r[4], water: r[11], future: r[0] > now,
  }));
  const algDays = {};
  fRows.forEach(r => { if (r[0] instanceof Date && r[FL.ALLERGEN - 1]) algDays[key_(r[0])] = true; });
  days.forEach(d => { d.allergen = !!algDays[d.date]; });

  // Foods for the logging screen, most useful first
  const due = al.filter(a => a.status === 'Due').map(a => a.name);
  const foods = Object.keys(guide).map(n => Object.assign({}, guide[n], { tries: tries[n] || 0 }));
  foods.forEach(f => { f.rank = (due.indexOf(f.allergen) >= 0 ? 0 : 10) + (f.iron ? 0 : 3) + (f.tries ? 0 : 1); });
  foods.sort((a, b) => a.rank - b.rank || a.name.localeCompare(b.name));

  return {
    today, age, targets, todayFoods, drinkToday, allergens: al, month, months, days, foods,
    triedCount: Object.keys(tries).length, guideCount: Object.keys(guide).length,
    user: Session.getActiveUser().getEmail(),
    lists: { forms: FORMS, likes: LIKES, reactions: REACTIONS, drinks: DRINKS },
  };
}

/* ---------------------------------------------------------------- write API */

/**
 * Log one or more foods. payload = { time: 'HH:mm', items: [{ food, form, servings, liked, reaction, notes }] }
 * Age, Allergen, Iron? and Group are left to the Sheet's formulas.
 */
function logFoods(payload) {
  const items = (payload && payload.items) || [];
  if (!items.length) throw new Error('Pick at least one food.');
  const guide = foodGuide_();
  const ids = withLock_(() => {
    const sh = ss_().getSheetByName(CONFIG.FOOD);
    let row = lastRow_(sh, FL.FOOD) + 1;
    const tries = {};
    if (row > 2) sh.getRange(2, FL.FOOD, row - 2, 1).getValues().forEach(r => { if (r[0]) tries[r[0]] = (tries[r[0]] || 0) + 1; });
    const now = new Date();
    const date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const time = payload.time || Utilities.formatDate(now, tz_(), 'H:mm');
    const out = [];
    items.forEach(it => {
      if (!guide[it.food]) throw new Error(it.food + ' is not in the Food Guide. Add it there first.');
      if (row > CONFIG.LOG_ROWS + 1) throw new Error('The Food Log is full. Extend LOG_ROWS and the Sheet formulas.');
      const id = newId_('F');
      const t = (tries[it.food] || 0) + 1; tries[it.food] = t;
      sh.getRange(row, FL.DATE).setValue(date).setNumberFormat('mmm d, yyyy');
      sh.getRange(row, FL.FOOD).setValue(it.food);
      sh.getRange(row, FL.TRY).setValue(t);
      sh.getRange(row, FL.FORM, 1, 1).setValue(FORMS.indexOf(it.form) >= 0 ? it.form : '');
      sh.getRange(row, FL.LIKED, 1, 3).setValues([[LIKES.indexOf(it.liked) >= 0 ? it.liked : '',
        REACTIONS.indexOf(it.reaction) >= 0 ? it.reaction : 'None', time]]);
      sh.getRange(row, FL.NOTES).setValue(it.notes || '');
      sh.getRange(row, FL.SERVINGS).setValue(Number(it.servings) || 1);
      sh.getRange(row, FL.ID).setValue(id);
      changeLog_('added', CONFIG.FOOD, id, it.food + ', try ' + t + (it.reaction && it.reaction !== 'None' ? ', reaction: ' + it.reaction : ''));
      out.push(id); row++;
    });
    return out;
  });
  SpreadsheetApp.flush();
  if (CONFIG.DAILY_AS_VALUES) refreshDailyValues();
  return { ids, dashboard: getDashboard() };
}

/** Log one drink. payload = { drink, oz, time } ; a breastfeed needs no oz. */
function logDrink(payload) {
  if (DRINKS.indexOf(payload.drink) < 0) throw new Error('Unknown drink: ' + payload.drink);
  const oz = payload.drink === 'Breast milk' ? '' : Number(payload.oz);
  if (payload.drink !== 'Breast milk' && !(oz > 0 && oz <= 40)) throw new Error('Amount must be 1 to 40 oz.');
  const id = withLock_(() => {
    const sh = ss_().getSheetByName(CONFIG.DRINKS);
    const row = lastRow_(sh, DL.DRINK) + 1;
    const now = new Date();
    const newId = newId_('D');
    sh.getRange(row, 1, 1, 6).setValues([[new Date(now.getFullYear(), now.getMonth(), now.getDate()),
      payload.time || Utilities.formatDate(now, tz_(), 'H:mm'), payload.drink, oz, payload.notes || '', newId]]);
    sh.getRange(row, 1).setNumberFormat('mmm d, yyyy');
    changeLog_('added', CONFIG.DRINKS, newId, payload.drink + (oz ? ', ' + oz + ' oz' : ''));
    return newId;
  });
  SpreadsheetApp.flush();
  if (CONFIG.DAILY_AS_VALUES) refreshDailyValues();
  return { ids: [id], dashboard: getDashboard() };
}

/** Undo: copies the row into the Change Log, then clears it. Works for F- and D- IDs. */
function removeEntries(ids) {
  withLock_(() => {
    (ids || []).forEach(id => {
      const isFood = String(id).indexOf('F-') === 0;
      const sh = ss_().getSheetByName(isFood ? CONFIG.FOOD : CONFIG.DRINKS);
      const idCol = isFood ? FL.ID : DL.ID;
      const last = lastRow_(sh, idCol);
      if (last < 2) return;
      const vals = sh.getRange(2, idCol, last - 1, 1).getValues();
      const i = vals.findIndex(r => r[0] === id);
      if (i < 0) return;
      const row = i + 2;
      const snapshot = sh.getRange(row, 1, 1, idCol).getDisplayValues()[0].join(' | ');
      changeLog_('removed', sh.getName(), id, snapshot);
      if (isFood) {
        // keep the formula columns (Age, Allergen, Iron?, Group) intact
        [[FL.DATE, 2], [FL.TRY, 1], [FL.FORM, 7], [FL.SERVINGS, 1], [FL.ID, 1]].forEach(([c, n]) => sh.getRange(row, c, 1, n).clearContent());
      } else {
        sh.getRange(row, 1, 1, DL.ID).clearContent();
      }
    });
  });
  SpreadsheetApp.flush();
  if (CONFIG.DAILY_AS_VALUES) refreshDailyValues();
  return getDashboard();
}

/* ---------------------------------------------------------------- edits made in the Sheet */

/** Simple trigger: stamps IDs on rows typed by hand and records edits in the Change Log. */
function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (name !== CONFIG.FOOD && name !== CONFIG.DRINKS) return;
  const isFood = name === CONFIG.FOOD;
  const keyCol = isFood ? FL.FOOD : DL.DRINK;
  const idCol = isFood ? FL.ID : DL.ID;
  const r0 = Math.max(2, e.range.getRow());
  const r1 = e.range.getLastRow();
  if (r1 < 2) return;
  const keys = sh.getRange(r0, keyCol, r1 - r0 + 1, 1).getValues();
  const ids = sh.getRange(r0, idCol, r1 - r0 + 1, 1).getValues();
  const stamped = ids.map((r, i) => [r[0] || (keys[i][0] ? newId_(isFood ? 'F' : 'D') : '')]);
  sh.getRange(r0, idCol, stamped.length, 1).setValues(stamped);
  if (e.range.getColumn() === idCol && e.oldValue) {
    sh.getRange(e.range.getRow(), idCol).setValue(e.oldValue);   // IDs are permanent
  }
  const detail = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1
    ? e.range.getA1Notation() + ': "' + (e.oldValue || '') + '" to "' + (e.value || '') + '"'
    : e.range.getA1Notation() + ' (' + e.range.getNumRows() + ' rows pasted or cleared)';
  try { changeLog_('edited in Sheet', name, stamped[0][0], detail, e.source); } catch (err) { /* simple triggers may lack access; skip */ }
}

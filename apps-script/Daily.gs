/**
 * Optional speed-up. The Daily tab ships with live formulas (about 27,000 lookups). If the Sheet
 * gets slow as the log grows, set CONFIG.DAILY_AS_VALUES = true: this function then fills Daily
 * with plain numbers every night and after every log from the app. Edits typed straight into the
 * Sheet show up after the next refresh (Louise menu > Refresh Daily values).
 * Running it once replaces Daily's formulas; to go back, copy Daily from a backup.
 */
function refreshDailyValues() {
  const ss = ss_();
  const tz = tz_();
  const k = d => Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  const dob = dob_();
  const per = {};
  const get = key => per[key] || (per[key] = { foods: 0, iron: 0, newf: 0, fruit: 0, veg: 0, grain: 0, protein: 0, dairy: 0,
    drinks: 0, water: 0, formula: 0, feeds: 0, cow: 0, juice: 0 });

  const fl = ss.getSheetByName(CONFIG.FOOD);
  const fLast = lastRow_(fl, FL.FOOD);
  if (fLast >= 2) fl.getRange(2, 1, fLast - 1, FL.ID).getValues().forEach(r => {
    if (!(r[0] instanceof Date) || !r[1]) return;
    const d = get(k(r[0]));
    const s = r[FL.SERVINGS - 1] === '' ? 1 : Number(r[FL.SERVINGS - 1]) || 0;
    d.foods += 1;
    if (r[FL.IRON - 1] === 'Iron') d.iron = 1;
    if (Number(r[FL.TRY - 1]) === 1) d.newf += 1;
    const g = r[FL.GROUP - 1];
    if (g === 'Fruit') d.fruit += s;
    else if (g === 'Vegetable') d.veg += s;
    else if (g === 'Grain') d.grain += s;
    else if (g === 'Dairy') d.dairy += s;
    else if (g === 'Iron protein' || g === 'Fish' || g === 'Nut & seed') d.protein += s;
  });

  const dl = ss.getSheetByName(CONFIG.DRINKS);
  const dLast = lastRow_(dl, DL.DRINK);
  if (dLast >= 2) dl.getRange(2, 1, dLast - 1, DL.ID).getValues().forEach(r => {
    if (!(r[0] instanceof Date) || !r[2]) return;
    const d = get(k(r[0]));
    const oz = Number(r[DL.OZ - 1]) || 0;
    d.drinks += 1;
    if (r[2] === 'Water') d.water += oz;
    else if (r[2] === 'Formula') d.formula += oz;
    else if (r[2] === 'Breast milk') d.feeds += 1;
    else if (r[2] === 'Whole milk' || r[2] === 'Low-fat or skim milk') d.cow += oz;
    else if (r[2] === 'Juice') d.juice += oz;
  });

  const start = edate_(dob, 6);
  const out = [];
  for (let i = 0; i < 1700; i++) {
    const date = new Date(start.getFullYear(), start.getMonth(), start.getDate() + i);
    const d = per[k(date)] || get('__none__');
    out.push([date, monthsBetween_(dob, date), d.foods, d.iron, d.newf, d.fruit, d.veg, d.grain, d.protein, d.dairy,
      d.drinks, d.water, d.formula, d.feeds, d.cow, d.juice]);
  }
  const sh = ss.getSheetByName('Daily');
  sh.getRange(2, 1, out.length, 16).setValues(out);
  sh.getRange(2, 1, out.length, 1).setNumberFormat('mmm d, yyyy');
}

/**
 * The pediatrician report: one month of age, built from Monthly Progress, Allergens,
 * the Food Log and Pediatrician Qs, then turned into a PDF by Apps Script (free).
 */

function reportPdf_(age) {
  const html = buildReportHtml_(age);
  return Utilities.newBlob(html, 'text/html', 'report.html').getAs(MimeType.PDF)
    .setName(`Louise feeding report, ${age} months.pdf`);
}

function esc_(v) {
  return String(v === null || v === undefined ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function buildReportHtml_(age) {
  const ss = ss_();
  const tz = tz_();
  const dob = dob_();
  const fmt = d => Utilities.formatDate(d, tz, 'MMM d, yyyy');
  const from = edate_(dob, age), to = new Date(edate_(dob, age + 1).getTime() - 86400000);

  // Monthly Progress row
  const mp = ss.getSheetByName('Monthly Progress').getRange(4, 1, 56, 33).getDisplayValues();
  const head = mp[0];
  const row = mp.slice(1).find(r => parseInt(r[0], 10) === age) || [];
  const m = {}; head.forEach((h, i) => { m[h] = row[i] || ''; });

  // Targets for her band
  const at = ss.getSheetByName('Age Targets').getRange(5, 1, 5, 23).getDisplayValues();
  let band = at[0]; at.forEach(r => { if (r[1] !== '' && age >= Number(r[1])) band = r; });

  // Foods tried, from the whole Food Log
  const fl = ss.getSheetByName(CONFIG.FOOD);
  const last = lastRow_(fl, FL.FOOD);
  const foods = {};
  if (last >= 2) {
    fl.getRange(2, 1, last - 1, FL.ID).getValues().forEach(r => {
      if (!r[1]) return;
      const f = foods[r[1]] || (foods[r[1]] = { first: null, tries: 0, liked: '', reactions: [] });
      f.tries += 1;
      if (r[0] instanceof Date && r[0] <= to && (!f.first || r[0] < f.first)) f.first = r[0];
      if (r[FL.LIKED - 1]) f.liked = r[FL.LIKED - 1];
      const rx = r[FL.REACTION - 1];
      if (rx && rx !== 'None') f.reactions.push((r[0] instanceof Date ? fmt(r[0]) + ': ' : '') + rx);
    });
  }
  const foodRows = Object.keys(foods).sort((a, b) => {
    const fa = foods[a].first, fb = foods[b].first;
    return (fa ? fa.getTime() : 0) - (fb ? fb.getTime() : 0);
  }).map(n => `<tr><td>${esc_(n)}</td><td>${foods[n].first ? fmt(foods[n].first) : 'Before log'}</td>` +
    `<td class="n">${foods[n].tries}</td><td>${esc_(foods[n].liked)}</td><td>${esc_(foods[n].reactions.join('; ') || 'None')}</td></tr>`).join('');

  // Allergens
  const al = ss.getSheetByName('Allergens').getRange(2, 1, 9, 7).getDisplayValues();
  const alRows = al.map(r => `<tr><td>${esc_(r[1])}</td><td>${esc_(r[0])}</td><td>${esc_(r[3])}</td>` +
    `<td class="n">${esc_(r[4])}</td><td class="n">${esc_(r[5])}</td><td>${esc_(r[6])}</td></tr>`).join('');

  // Questions for the next visit
  const qs = ss.getSheetByName('Pediatrician Qs');
  const qLast = lastRow_(qs, 2);
  const qVals = qLast >= 2 ? qs.getRange(2, 1, qLast - 1, 4).getDisplayValues() : [];
  const visits = qVals.map(r => parseInt(r[0], 10)).filter(n => n >= age).sort((a, b) => a - b);
  const visit = visits.length ? visits[0] : null;
  const qList = qVals.filter(r => parseInt(r[0], 10) === visit)
    .map(r => `<li>${esc_(r[1])}${r[3] ? ' <span class="muted">Answer: ' + esc_(r[3]) + '</span>' : ''}</li>`).join('');

  const measure = (label, val, target) => `<tr><td>${label}</td><td class="n">${esc_(val)}</td><td class="n">${esc_(target)}</td></tr>`;
  const cue = age < 12;

  return `<!doctype html><html><head><meta charset="utf-8"><style>
  body{font-family:Arial,Helvetica,sans-serif;color:#1F2A33;font-size:11pt;margin:36px 44px}
  h1{font-size:20pt;margin:0}h2{font-size:13pt;margin:18px 0 4px}
  table{width:100%;border-collapse:collapse;font-size:10pt}
  th,td{text-align:left;padding:4px 6px;border-bottom:1px solid #C9C2B4;vertical-align:top}
  th{border-bottom:1.5px solid #1F2A33}.n{text-align:right}.muted{color:#4B5763}
  .top{display:flex;justify-content:space-between}.small{font-size:9pt;color:#4B5763}
  </style></head><body>
  <div class="top"><div><h1>Feeding and allergen report</h1>
  <p style="margin:4px 0 0">Louise, born ${fmt(dob)}, age ${age} months</p></div>
  <div class="small" style="text-align:right">${fmt(from)} to ${fmt(to)}<br>Age band: ${esc_(band[0])}<br>From parent logs</div></div>

  <h2>This month against her age band</h2>
  <table><tr><th>Measure</th><th class="n">Louise</th><th class="n">Target</th></tr>
  ${measure('Days logged', m['Days logged'] + (m['Days so far'] ? ' of ' + m['Days so far'] : ''), '')}
  ${measure('Days with an iron food', m['Iron days'], cue ? 'Every day' : '')}
  ${measure('New foods this month', m['New foods'], '')}
  ${measure('Foods tried to date', m['New foods to date'], '')}
  ${measure('Allergens started', m['Allergens started to date'] + ' of 9', '')}
  ${measure('Reactions', m['Reactions'], '')}
  ${cue ? measure('Food groups', 'Follows her cues', 'No amounts before 12 months')
        : measure('Food groups on target', m['Food groups on target'], 'All 5')}
  ${measure('Water, oz a day', m['Water oz a day'], band[15] + ' to ' + band[16])}
  ${measure('Breastfeeds a day', m['Breastfeeds a day'], 'On demand')}
  ${measure('Formula, oz a day', m['Formula oz a day'], band[17] ? band[17] + ' to ' + band[18] + ' if used' : '')}
  ${measure("Cow's milk, oz a day", m["Cow's milk oz a day"], band[20] === '0' ? 'None' : 'Up to ' + band[20])}
  ${measure('Juice, oz a day', m['Juice oz a day'], band[21] === '0' ? 'None' : 'Up to ' + band[21])}
  ${measure('Drinks check', m['Drinks check'], '')}
  ${measure('Weight (lb) / length (in)', (m['Weight (lb)'] || '-') + ' / ' + (m['Length (in)'] || '-'), '')}
  </table>

  <h2>Allergens</h2>
  <table><tr><th>Allergen</th><th>Plan</th><th>First tried</th><th class="n">Last 7 days</th><th class="n">Weekly target</th><th>Status</th></tr>${alRows}</table>

  <h2>Foods tried</h2>
  <table><tr><th>Food</th><th>First</th><th class="n">Tries</th><th>Liked it?</th><th>Reactions</th></tr>${foodRows || '<tr><td colspan="5">None logged yet</td></tr>'}</table>

  ${qList ? `<h2>Our questions for the ${visit}-month visit</h2><ol>${qList}</ol>` : ''}
  <p class="small" style="margin-top:18px">Parent-recorded data, not a clinical assessment. Targets from the Age Targets tab.
  Generated ${fmt(new Date())} from Louise's First Foods Log in Google Sheets.</p>
  </body></html>`;
}

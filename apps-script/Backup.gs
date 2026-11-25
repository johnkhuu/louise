/**
 * Backups, all inside your own Google Drive.
 *  - Nightly: a full copy of the Sheet in "Louise backups"; copies older than KEEP_NIGHTLY_DAYS go to the trash.
 *  - Monthly, on her monthly birthday: .xlsx, CSVs of both logs and the PDF doctor report in
 *    "Louise monthly archive", and the report emailed to you.
 * Download the monthly folder to a computer or USB drive now and then: that copy survives
 * anything that happens to the Google account.
 */

function nightlyJobs() {
  nightlyBackup();
  if (CONFIG.DAILY_AS_VALUES) refreshDailyValues();
  if (isMonthlyBirthday_(new Date())) monthlyArchive();
}

function nightlyBackup() {
  const ss = ss_();
  const folder = folder_(CONFIG.BACKUP_FOLDER);
  const stamp = Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
  DriveApp.getFileById(ss.getId()).makeCopy(`Louise backup ${stamp}`, folder);
  const cutoff = Date.now() - CONFIG.KEEP_NIGHTLY_DAYS * 86400000;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf('Louise backup ') === 0 && f.getDateCreated().getTime() < cutoff) f.setTrashed(true);
  }
}

function isMonthlyBirthday_(d) {
  const day = dob_().getDate();
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate();
  return d.getDate() === Math.min(day, lastDay);
}

/** Builds the archive for the month that just ended (or the current one, when run by hand). */
function monthlyArchive() {
  const ss = ss_();
  const now = new Date();
  const age = monthsBetween_(dob_(), now);
  const reportAge = isMonthlyBirthday_(now) ? age - 1 : age;
  const label = Utilities.formatDate(now, tz_(), 'yyyy-MM') + `, ${reportAge} months`;
  const parent = folder_(CONFIG.ARCHIVE_FOLDER);
  const existing = parent.getFoldersByName(label);
  const folder = existing.hasNext() ? existing.next() : parent.createFolder(label);

  // Whole workbook as .xlsx (Drive export, free)
  const url = `https://docs.google.com/spreadsheets/d/${ss.getId()}/export?format=xlsx`;
  const xlsx = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  folder.createFile(xlsx.setName(`Louises_First_Foods_Log ${label}.xlsx`));

  // Both logs as CSV: plain text that any program will open in 20 years
  [[CONFIG.FOOD, FL.FOOD, FL.ID], [CONFIG.DRINKS, DL.DRINK, DL.ID]].forEach(([name, keyCol, width]) => {
    const sh = ss.getSheetByName(name);
    const last = lastRow_(sh, keyCol);
    const rows = sh.getRange(1, 1, Math.max(1, last), width).getDisplayValues();
    const csv = rows.map(r => r.map(v => /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v).join(',')).join('\n');
    folder.createFile(`${name} ${label}.csv`, csv, MimeType.CSV);
  });

  // Doctor report as PDF, saved and emailed
  const pdf = reportPdf_(reportAge);
  folder.createFile(pdf);
  emailReport_(pdf, reportAge);
  changeLog_('archived', 'all', '', label);
}

function emailReportNow() {
  const age = monthsBetween_(dob_(), new Date());
  emailReport_(reportPdf_(age), age);
  try { SpreadsheetApp.getUi().alert('Report sent', 'Check your inbox for the ' + age + '-month report.', SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* called from the web app */ }
  return true;
}

function emailReport_(pdf, age) {
  const to = CONFIG.REPORT_EMAILS.length ? CONFIG.REPORT_EMAILS.join(',') : Session.getEffectiveUser().getEmail();
  MailApp.sendEmail({
    to,
    subject: `Louise feeding report, ${age} months`,
    htmlBody: `<p>The ${age}-month feeding and allergen report is attached, ready to show her pediatrician.</p>` +
      '<p>A copy is also saved in the "Louise monthly archive" folder in Google Drive.</p>',
    attachments: [pdf],
  });
}

# Louise app on Google Sheets (free)

The logging app, backups and doctor reports, running entirely inside a free Google account.
No server, no database, no subscription.

| What | Where it runs | Cost |
|---|---|---|
| The log and all the dashboards | The Google Sheet (converted from `Louises_First_Foods_Log.xlsx`) | Free |
| The phone app (Today, Food, Drinks, Progress, Allergens) | Apps Script web app, `Index.html` | Free |
| IDs, change history, undo | `Code.gs` | Free |
| Nightly backup, monthly archive, emailed PDF report | `Backup.gs`, `Report.gs` | Free |
| Fixes for the Excel to Sheets conversion | `Setup.gs` | Free |
| Optional speed-up for a big log | `Daily.gs` | Free |

## Set up (about 20 minutes, once)

1. **Make the Google Sheet.** Upload `Louises_First_Foods_Log.xlsx` to Google Drive, open it, then
   File > Save as Google Sheets. Work in the new Google Sheet from now on, not the .xlsx.
2. **Add the code.** In the Sheet: Extensions > Apps Script. Create these files and paste each one in:
   `Code.gs`, `Setup.gs`, `Backup.gs`, `Report.gs`, `Daily.gs`, and an HTML file named `Index`.
   In Project Settings, tick "Show appsscript.json manifest file in editor", then paste `appsscript.json`
   over the one there. Change `timeZone` if you're not in Chicago. Save.
3. **Run setup.** Reload the Sheet. A **Louise** menu appears: Louise > Set up this Sheet.
   Google will ask you to authorize. Because it's your own script and not a published app, you'll see
   "Google hasn't verified this app": click Advanced, then Go to (project name). This is expected.
   Setup then:
   - rewrites the Food Log's Age, Allergen, Iron? and Group columns as `ARRAYFORMULA`
   - rewrites Allergens "First tried" without Excel's `LET`
   - wraps the weight and length lookups on Monthly Progress
   - rebuilds the Monthly Progress colours (Sheets can't point a colour rule at another tab without `INDIRECT`)
   - creates the Change Log tab, stamps IDs on existing rows, creates the two Drive folders and
     schedules the nightly job for about 2am.
4. **Check the conversion.** Make sure these work before relying on the Sheet:
   - Food Log columns C, E, F and O fill down for every row that has a food.
   - Allergens column D shows a date for Peanut once it's logged, and column G shows a status.
   - Monthly Progress shows this month in green.
   - Overview's "Right now" block has values.
5. **Publish the app.** In Apps Script: Deploy > New deployment > type Web app.
   - Execute as: **User accessing the web app**
   - Who has access: **Anyone with Google account**

   Copy the web app URL. Only people the Sheet is shared with can read or write anything; anyone
   else who opens the link gets an error. Never deploy as "Execute as: Me" with "Anyone", which would
   let anyone with the link read Louise's data.
6. **Share and install.** Share the Sheet with your partner as an Editor. Both of you open the
   web app URL on your phones, sign in, allow access once, then Add to Home Screen (iPhone: Share >
   Add to Home Screen; Android: menu > Add to Home screen).

**To update the code later:** paste the change, then Deploy > Manage deployments > Edit (pencil) >
Version: New version > Deploy. The link stays the same.

## Settings (top of `Code.gs`)

- `REPORT_EMAILS`: who gets the monthly report. Empty means the Sheet owner.
- `KEEP_NIGHTLY_DAYS`: how many days of nightly copies to keep (60).
- `DAILY_AS_VALUES`: turn on if the Sheet gets slow (see Limits). The script then fills the Daily tab
  with plain numbers instead of 27,000 live formulas.

## How your data is kept

- **Every row has a permanent ID** (`F-` for food, `D-` for drinks). The app stamps it, and so does
  the Sheet when you type a row by hand. IDs can't be edited.
- **Change Log tab.** Every add, undo and hand edit is recorded with who and when. Undo copies the
  row into the Change Log before clearing it, so nothing is lost.
- **Nightly copy** of the whole Sheet in Drive > Louise backups, kept for 60 days.
- **Monthly archive** on each monthly birthday (the 22nd), in Drive > Louise monthly archive:
  the workbook as .xlsx, both logs as CSV, and the PDF doctor report, which is also emailed to you.
- **Your offline copy.** Every few months, download the Louise monthly archive folder to a computer
  or USB drive. It's the only copy that survives losing the Google account.

## Limits of the free setup

- **Slower than a real app.** Pages take about 1 to 3 seconds, and the app sits inside Google's frame.
  Visitors other than the owner see a small "created by a Google Apps Script user" banner.
- **Needs a signal.** No offline logging; a log made with no connection fails and has to be redone.
- **Script quotas on a free account.** 6 minutes per run, about 90 minutes of scheduled runs a day,
  100 emails a day. This app uses a few seconds a day.
- **Storage.** The free 15 GB is shared with Gmail and Photos. Google Sheets count toward it, so 60
  nightly copies may take a few hundred MB. Don't store photos in the Sheet.
- **Speed as the log grows.** Sheets recalculates more slowly than Excel. If edits start to lag, set
  `DAILY_AS_VALUES = true` and run Louise > Refresh Daily values. Daily then updates after each app
  log and every night, and edits typed into the Sheet show up after the next refresh.
- **One account holds everything.** Turn on 2-step verification for the owner account, and keep the
  offline copy.
- **Version history** in Sheets is a convenience, not a guarantee: Google merges old versions. The
  nightly copies are the real history.

## Files

- `Code.gs`: web app entry point (`doGet`), dashboard data (`getDashboard`), logging (`logFoods`,
  `logDrink`, `removeEntries`), ID stamping and change logging (`onEdit`), Louise menu.
- `Setup.gs`: one-time conversion fixes, colour rules, folders, nightly trigger.
- `Backup.gs`: `nightlyJobs`, `nightlyBackup`, `monthlyArchive`, `emailReportNow`.
- `Report.gs`: the monthly pediatrician report, built as HTML and turned into a PDF.
- `Daily.gs`: optional script-filled Daily tab.
- `Index.html`: the phone app.
- `appsscript.json`: manifest (time zone, web app access).

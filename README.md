# Louise

A free food, drink and allergen tracker for Louise's first years of eating, built for her parents to log
quickly and to show her pediatrician a clear monthly report.

- **Data:** a Google Sheet converted from `Louises_First_Foods_Log.xlsx`. It holds the Food Guide,
  the allergen plan, age-based targets from 6 months to 5 years, the food and drink logs, and
  monthly progress.
- **App and automation:** Google Apps Script in [`apps-script/`](apps-script/). See its
  [README](apps-script/README.md) for setup, backups and limits.
- **Cost:** nothing beyond a free Google account.
- **Vercel:** the repo is connected to Vercel, but nothing here is deployed there. `vercel.json` tells
  Vercel to skip every build; delete it if a Vercel-hosted page is ever added.

Keep the workbook and any exports out of this repository: they contain a child's health data.

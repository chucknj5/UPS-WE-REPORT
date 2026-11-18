# UPS W.E. Report Builder

Daily → weekly W.E. report for UPS Time Clock (uAttend) timecards.

- `index.html` — the tool (drop timecards, or **⟳ SYNC** to pull from uAttend via Google Sheets)
- `netlify/functions/state.mjs` — shared SAVE for everyone on the Netlify link (Netlify Blobs, one backup per day)
- `apps-script/uAttend_Sync.gs` — Google Apps Script for the Sheet "uAttend Timecard Data (W.E. Report)":
  pulls uAttend daily at 8 AM, imports Timecard Report CSVs, and serves the data to the ⟳ SYNC button.
  The uAttend API key lives only in the script's Properties — never in this repo.

Deploy: Netlify → Add new project → Import an existing project → GitHub → this repo. No build settings; every push redeploys.

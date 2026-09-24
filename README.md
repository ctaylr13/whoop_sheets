# WHOOP → Google Sheets

Google Apps Script integration that syncs WHOOP recovery and sleep data into a
personal check-in spreadsheet.

The integration writes WHOOP-owned data to the separate `WHOOP Daily` tab and
upserts a base row in `Raw Entries` for each WHOOP date. The base row contains
wake/bed time, HRV, RHR, recovery, and sleep metrics. The Shortcut finds that
same name/date row and fills the behavioral check-in fields; it remains the
only thing that marks `Morning Complete`.

## What syncs

- Main sleep start/end, bedtime, and wake time
- HRV and resting heart rate
- Recovery score
- Sleep performance, consistency, and efficiency
- Respiratory rate
- Total sleep, time in bed, awake/light/deep/REM durations, and sleep need
- SpO2 and skin temperature when WHOOP returns them
- WHOOP record IDs and sync timestamps for safe upserts

The sync re-reads the last fourteen days so later WHOOP corrections are picked
up.

## One-time setup

1. In the Google Sheet, open **Extensions → Apps Script**.
2. Add the Google-maintained OAuth2 library:
   - Script ID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`
   - Identifier: `OAuth2`
   - Select the latest published version.
3. In **Project Settings → Script Properties**, add:
   - `SPREADSHEET_ID`: the ID from the Google Sheet URL
   - `DEFAULT_NAME`: the name used when a request omits one
   - `DAILY_REQUEST_SECRET`: a long, randomly generated secret for daily
     requests
   - `NFC_REQUEST_SECRET`: a different long, randomly generated secret for NFC
     requests
   - `WHOOP_CLIENT_ID`: the client ID from the WHOOP Developer Dashboard
   - `WHOOP_CLIENT_SECRET`: the client secret from the WHOOP Developer Dashboard
4. Copy `Code.gs`, `Whoop.gs`, and `appsscript.json` into the bound Apps Script
   project.
5. Run `showWhoopRedirectUri()` and copy the returned URI.
6. In the WHOOP Developer Dashboard, create an app with `read:sleep` and
   `read:recovery`, then register the displayed redirect URI.
7. Run `showWhoopAuthorizationUrl()`, open the logged URL, and grant access.
8. Run `syncWhoopDaily()` to test the connection.
9. Run `installWhoopSyncTrigger()` to install the hourly sync.

Do not commit a client secret, access token, or refresh token.

## Optional clasp workflow

Copy `.clasp.json.example` to `.clasp.json`, insert the bound script ID, then
use `clasp push` to publish local changes. `.clasp.json` is ignored so each
developer can point at their own bound script.

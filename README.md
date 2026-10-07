# Genmo GT Monitoring

Schedule Transformer for Gen Mobile GT schedules.

## Current workflow

- Upload a new Excel schedule when you have one.
- The original Excel is stored as a template in Firebase Storage.
- Saved weekly schedules are written to the configured Google Sheet.
- The Google Sheet `Usertimeline` is the source for the GT agent roster.
- Reusing a saved week filters to GT agents valid for the new week and removes termination/leave rows before shifting the schedule to the new week.

## Google Sheets connection

The app uses a Google Apps Script Web App as the bridge to the Google Sheet.

1. Open the target Google Sheet.
2. Go to **Extensions → Apps Script**.
3. Copy `google-apps-script/Code.gs` from this repository into the Apps Script project.
4. Deploy it as a **Web app**.
5. Set the deployment to execute as the spreadsheet owner.
6. Copy the Web App URL.
7. Configure that URL as `VITE_GOOGLE_SHEETS_API_URL` when building the app.

The Apps Script reads the `Usertimeline` tab and can create/update weekly tabs such as `10-12-2026`.

> The Google Apps Script deployment is intentionally not committed with an access token or secret.

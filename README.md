# running-schedule-sync

A Google Apps Script that syncs Strava running activity into a Google Sheet used to track a
training schedule. It compares actual mileage/workouts against a planned schedule, records
workout splits, and maintains a per-body-part injury report.

## What it does

The script is bound to a Google Sheet with (up to) four tabs:

- **Actual Runs** — one row per training week, one column per day (Mon–Sun) plus a weekly total.
- **Planned Schedule** — the same layout, but with the planned mileage/rest days for comparison.
- **Workout Splits** — a log of interval/workout splits parsed out of Strava activity descriptions.
- **`<year> Injury Report`** — a sheet (one per year, auto-created) with a row per day and a column
  per reported body part, color-coded by injury severity.

The entry point is `syncStravaToActualRuns()` (in [DataPlacement.js](DataPlacement.js)), which is
meant to be run on a time-driven trigger (e.g. daily). Each run:

1. Finds the current week's row in **Actual Runs** (and the matching row in **Planned Schedule**).
2. Fetches all Strava activities since that Monday, including each activity's full description.
3. Updates the daily cells in **Actual Runs** with mileage and any parsed workout text, color-coded
   against the plan.
4. Updates the weekly total cell's color based on how close actual miles are to planned miles.
5. Appends any workout splits found in activity descriptions to **Workout Splits**.
6. Updates the **Injury Report** sheet for the week from any injury reports found in activity
   descriptions.

## How it's organized

The code is a plain Apps Script project (no build step) split into small, single-purpose files:
one for talking to the Strava API, one for date/week math, one for computing mileage from
activities, and one each for maintaining the Actual Runs, Workout Splits, and Injury Report
sheets. [DataPlacement.js](DataPlacement.js) is the only file that wires them together into the
`syncStravaToActualRuns()` entry point described above.

Everything is deployed with `clasp`, and there's a small local harness ([run-local.js](run-local.js)
plus the mocks in [local/](local)) that lets any function run and be inspected in Node without
touching a real spreadsheet or Strava account — see "Local development" below.

## Activity description format

Strava activity descriptions are parsed for special sections:

- **`Workout:`** followed by a line of free text — copied into the daily **Actual Runs** cell and
  the **Workout Splits** sheet.
- **`Supplemental Workout:`** followed by one supplemental exercise per line — each one is added as
  its own line in the daily **Actual Runs** cell.
- **`Splits:`** followed by one split per line (e.g. `6:45` or `6:45 6:32`) — each line's first
  duration is the split time; a second duration, if present, is used as the pace, otherwise pace is
  calculated from the workout's parsed distance.
- **`Injury Report:`** followed by `Area: <body part>` and a `Severity: X/10` line, plus free-text
  description — recorded on the **Injury Report** sheet.

## Injury report coloring

Each body-part column on the Injury Report sheet is colored independently for the day, based on
that specific report's severity (0 = green, 10 = red, blended in between). Body parts with no
injury reported that day default to severity 0 (green) rather than being left blank, so the sheet
always shows a color per cell rather than coloring the whole row by the worst injury of the day.

## Local development

```powershell
npm test           # run the unit tests (node --test)
npm run local       # run syncStravaToActualRuns() locally against fixtures/spreadsheet.json
npm run push         # git push, then clasp push to deploy to Apps Script
```

`npm run local` executes the script in Node using `local/GoogleSheetsMock.js` in place of
`SpreadsheetApp`, and `curl` in place of `UrlFetchApp`, writing the resulting sheet state to
`output/spreadsheet.json` (and `output/spreadsheet.html` for a visual preview). Strava credentials
for local runs are read from a `.env` file (`STRAVA_CLIENT_ID`, `STRAVA_CLIENT_SECRET`,
`STRAVA_REFRESH_TOKEN`); when deployed, the same keys are read from Apps Script's
`PropertiesService` script properties instead.

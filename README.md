# running-schedule-sync

A Google Apps Script that syncs Strava running activity into a Google Sheet used to track a
training schedule. It compares actual mileage/workouts against a planned schedule, records
workout splits, and maintains a per-body-part injury report.

## What it does

The script is bound to a Google Sheet with (up to) four tabs:

- **Actual Runs** — one row per training week, one column per day (Mon–Sun) plus a weekly total.
- **Planned Schedule** — the same layout, but with the planned mileage/rest days for comparison.
- **Workout Splits** — a log of interval/workout splits parsed out of Strava activity descriptions.
- **Injury Report** — a sheet with a row per day and a column per reported body part,
  color-coded by injury severity.

The entry point is `syncStravaToActualRuns()` (in [DataPlacement.js](DataPlacement.js)), which is
meant to be run on a time-driven trigger (e.g. daily). Each run:

1. Finds the current week's row in **Actual Runs** (and the matching row in **Planned Schedule**).
2. Fetches all Strava activities since that Monday, including each activity's full description.
3. Updates the daily cells in **Actual Runs** with mileage and any parsed workout text, color-coded
   against the plan (within 1/2 mile green, 1/2–1 mile yellow, more than 1 mile red).
4. Updates the weekly total cell's color based on what percentage actual miles are off the planned
   total (within 5% green, 5–10% yellow, more than 10% red).
5. Appends any workout splits found in activity descriptions to **Workout Splits**.
6. Updates the **Injury Report** sheet for the week from any injury reports found in activity
   descriptions.
7. Re-reads the **Supplemental Workouts** sheet and appends each day's logged `Workout` categories
   to that day's **Actual Runs** cell (see "Supplemental workouts in Actual Runs" below).

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
- **`Supplemental Workouts:`** followed by workout categories and numbered exercise rows — category
  names are added to **Actual Runs**, and each exercise is logged on the **Supplemental Workouts**
  sheet. Exercises in the same category on the same day share a merged Date/Workout block, with a
  thin black divider row between different days. Weights mentioned in an exercise's parentheses
  (e.g. `70 lbs`) go in their own Weight column, one per line when a set uses multiple weights;
  exercises with no weight get `N/A`. **A day's existing rows are only rewritten if that day's
  Actual Runs cell shows a run** — see "Overwrite rules" below. A day's rows can also be entered
  manually (e.g. a supplemental workout on a rest day), and those are detected and preserved
  automatically; putting `Non Running Day` in that day's Notes column (case insensitive,
  dashes/underscores allowed) pins a block explicitly as well.
- **`Splits:`** followed by one split per line (e.g. `6:45` or `6:45 6:32`) — each line's first
  duration is the split time; a second duration, if present, is used as the pace, otherwise pace is
  calculated from the workout's parsed distance.
- **`Injury Report:`** followed by `Area: <body part>` and a `Severity: X/10` line, plus free-text
  description — recorded on the **Injury Report** sheet.

## Overwrite rules

Rest-day entries on **Supplemental Workouts** are usually typed by hand, but nothing in the sheet
marks a row as manual — so it is derived instead. The script only looks at the week since the most
recent Monday, and within that week any date whose **Actual Runs** cell reads `Rest` (or is still
blank) had no Strava run, which means its Supplemental Workouts rows must have been entered by hand.
Those rows are left exactly as typed, even if some activity's description mentions that day. A day
whose Actual Runs cell shows a run is still regenerated from the activity description as before, and
brand-new days are always added. (The `Non Running Day` note still pins a block as well.)

Anything dated before that Monday belongs to the historical log and is never rewritten or moved.

Because a hand-typed entry only exists where it was typed, `npm run local` cannot invent one: it
syncs whatever is already in `fixtures/spreadsheet.json`. If you added a rest-day workout directly
in the Google Sheet, run `npm run update:supworkouts` (or `npm run update:all`) first to pull that
tab down — otherwise the local preview will simply be missing the entry that production still has.

Once every day has been processed, the sheet is put back in chronological order. Only blocks that
sit in the wrong place *and* fall inside the current week are moved, and a moved block's cells are
carried over verbatim, so rows the script did not generate keep their values, merges, and
formatting. Ordering is skipped entirely if any block's date cannot be read.

Daily cells on **Actual Runs** are merged line-by-line rather than rebuilt, so re-running a sync
never duplicates a workout name that is already in the cell.

## Supplemental workouts in Actual Runs

After the planned/actual evaluation finishes, the script makes a second pass over the
**Supplemental Workouts** sheet for the week being synced. For each day with logged exercises, the
distinct values in that day's `Workout` column are appended as extra lines to the matching
**Actual Runs** cell:

```
Rest
Upper Body Lift
Core
```

This is what surfaces rest-day work that Strava never sees — a day whose rows were entered by hand
(e.g. a `Non Running Day` block) still shows up next to its `Rest` line. The pass runs for every day
up to today, not just rest days, so a running day picks up its categories too:

```
5.78 miles
Shoulder Rehab
Accessory Exercises
Lower Body Lift
```

Matching is case-insensitive and only missing lines are added, so repeated syncs leave an
already-complete cell untouched. Categories are appended in the order they appear on the sheet, and
mileage, workout names, and cell coloring are preserved.

## Injury report coloring

Each body-part column on the Injury Report sheet is colored independently for the day, based on
that specific report's severity (0 = green, 10 = red, blended in between). Body parts with no
injury reported that day default to severity 0 (green) rather than being left blank, so the sheet
always shows a color per cell rather than coloring the whole row by the worst injury of the day.

## Local development

```powershell
npm test           # run the unit tests (node --test)
npm run local       # refresh Actual Runs, then sync live Strava since the most recent Monday
npm run local -- 2026-09-01 # live Strava, since the supplied date
npm run local -- 2026-09-01 mocks # local mocks, since the supplied date
npm run local -- pull 2026-09-01 # pull Strava, save mocks, then update spreadsheet.json
npm run auth:google # authorize local read-only Google Sheets access once
npm run update:planned # pull production Planned Schedule into fixtures/spreadsheet.json
npm run update:supworkouts # pull production Supplemental Workouts into fixtures/spreadsheet.json
npm run update:all   # pull both production tabs in a single pass
npm run push         # git push, then clasp push to deploy to Apps Script
```

`clasp` is a local dev dependency, so deploying needs no global install — but it does need a
one-time `npx clasp login`. `npm run push` runs `git push` first and aborts if that fails, so the
working copy needs its `origin` remote configured. Use `npx clasp status` to confirm which files are
tracked; only the nine Apps Script files are pushed, while `*.test.js`, `local/`, `fixtures/`, and
`output/` stay local via `.claspignore`.

Because the deployed script and the local harness differ (`SupplementalWorkouts.js` branches on
whether `PropertiesService` exists), `SupplementalWorkouts.production.test.js` re-runs the manual
rest-day scenarios through the production branch, with a mock that enforces Google Sheets' real
merged-cell rules — writing a range that splits a merged region throws, and a merged region reads
back only through its top-left cell.

`npm run local` first refreshes **Actual Runs** from the production spreadsheet so completed prior
weeks cannot be replaced by stale fixture data. It then executes the script in Node using
`local/GoogleSheetsMock.js` in place of `SpreadsheetApp`, and `curl` in place of `UrlFetchApp`,
writing the resulting sheet state to `output/spreadsheet.json` (and `output/spreadsheet.html` for a
visual preview). Strava credentials for live runs are read from a `.env` file
(`STRAVA_CLIENT_ID`, `STRAVA_CLIENT_SECRET`, `STRAVA_REFRESH_TOKEN`); when deployed, the same keys
are read from Apps Script's `PropertiesService` script properties instead. Add the positional word
`mocks` to use files from `local/mocks/` instead of contacting Strava; mock runs remain offline and
use the existing fixture.

`npm run local -- pull 2026-09-01` runs the pull and local spreadsheet sync sequentially. It saves
each running activity under `local/mocks/` and then writes the complete local workbook to
`output/spreadsheet.json`. This is the only combined pull-and-sync command.

## Local Google Sheets authentication

For local access to the private production spreadsheet:

1. In Google Cloud Console, create an OAuth client for a **Desktop app** and download its JSON.
2. Save that file as `local/google-oauth-client.json`.
3. Run `npm run auth:google` and authorize the requested read-only Sheets scope in the browser.
4. Run `npm run update:planned` to update only `fixtures/spreadsheet.json`'s **Planned Schedule** tab.
5. Run `npm run update:supworkouts` to pull the **Supplemental Workouts** tab, which is how manually
   entered rows (such as `Non Running Day` entries) propagate from production into local runs.
6. Run `npm run update:all` to pull both tabs at once.
7. Run `npm run local:last-month` to sync the previous month using that planned data.

The client JSON and refresh token are ignored by git. Production does not use this OAuth flow; it
continues reading the active spreadsheet through Apps Script's `SpreadsheetApp` authorization.

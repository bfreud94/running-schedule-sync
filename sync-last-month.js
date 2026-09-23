const { execFileSync } = require('node:child_process');
const { readFileSync } = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createGoogleSheetsMock } = require('./local/GoogleSheetsMock');
const { loadActivityMocks } = require('./local/activityMocks');

const fixtureArgumentIndex = process.argv.indexOf('--sheet-fixture');
const fixturePath = fixtureArgumentIndex === -1
  ? path.resolve('fixtures/spreadsheet.json')
  : path.resolve(process.argv[fixtureArgumentIndex + 1]);
const sinceValue = process.argv.slice(2).find(argument => /^\d{4}-\d{2}-\d{2}$/.test(argument));
const outputPath = path.resolve('output/spreadsheet.json');

process.chdir(__dirname);

const sheetsMock = createGoogleSheetsMock(fixturePath, outputPath);
globalThis.SpreadsheetApp = sheetsMock.SpreadsheetApp;
globalThis.Logger = { log: (...args) => console.log(...args) };
globalThis.UrlFetchApp = {
  fetch(url, options = {}) {
    const args = [
      '--silent',
      '--show-error',
      '--location',
      '--request',
      String(options.method || 'get').toUpperCase()
    ];

    Object.entries(options.headers || {}).forEach(([name, value]) => {
      args.push('--header', `${name}: ${value}`);
    });
    Object.entries(options.payload || {}).forEach(([name, value]) => {
      args.push('--data-urlencode', `${name}=${value}`);
    });

    args.push('--write-out', '\n%{http_code}', String(url));
    const output = execFileSync('curl', args, {
      encoding: 'utf8',
      maxBuffer: 10 * 1024 * 1024
    });
    const statusSeparator = output.lastIndexOf('\n');
    const content = output.slice(0, statusSeparator);
    const statusCode = Number(output.slice(statusSeparator + 1).trim());

    return {
      getContentText: () => content,
      getResponseCode: () => statusCode
    };
  }
};

const scriptFiles = [
  'DateUtils.js',
  'RunningMetrics.js',
  'WorkoutSplits.js',
  'SupplementalWorkouts.js',
  'SheetUtils.js',
  'InjuryReport.js',
  'StravaAPI.js',
  'DataPlacement.js'
];
const scriptSource = scriptFiles
  .map(fileName => readFileSync(path.join(__dirname, fileName), 'utf8'))
  .join('\n');

function getMonthStartDate(currentDate = new Date()) {
  const monthStart = new Date(currentDate);
  monthStart.setMonth(monthStart.getMonth() - 1);
  monthStart.setHours(0, 0, 0, 0);
  return monthStart;
}

function getMondayOnOrBefore(date) {
  const monday = getStartOfDay(date);
  const daysSinceMonday = monday.getDay() === 0 ? 6 : monday.getDay() - 1;
  monday.setDate(monday.getDate() - daysSinceMonday);
  return monday;
}

function getMondayOnOrAfter(date) {
  const monday = getMondayOnOrBefore(date);
  if (monday < getStartOfDay(date)) monday.setDate(monday.getDate() + 7);
  return monday;
}

function getCoveredWeekStarts(startDate, endDate) {
  const starts = [];
  for (let weekStart = getMondayOnOrBefore(startDate); weekStart <= endDate; weekStart.setDate(weekStart.getDate() + 7)) {
    starts.push(new Date(weekStart));
  }
  return starts;
}

function getCurrentYearCutoff(month, day, currentDate) {
  return new Date(currentDate.getFullYear(), month, day, 0, 0, 0, 0);
}

function ensureActualWeekRow(actualSheet, targetMonday) {
  const actualData = actualSheet.getDataRange().getValues();
  const existingRowIndex = findWeekRowIndex(actualData, targetMonday);
  if (existingRowIndex !== -1) return existingRowIndex;

  const nextRowIndex = actualData.findIndex((row, index) => {
    const weekDate = getWeekDate(row[0]);
    return weekDate && weekDate > targetMonday;
  });
  if (nextRowIndex === -1) return appendWeekRow(actualSheet, actualData, targetMonday);

  const nextLabel = String(actualData[nextRowIndex][0] || '').match(/^WEEK\s+(-?\d+)/i);
  const nextDate = getWeekDate(actualData[nextRowIndex][0]);
  const weeksBeforeNext = Math.round(getDayOffset(nextDate, targetMonday) / 7);
  const weekNumber = nextLabel ? parseInt(nextLabel[1], 10) - weeksBeforeNext : 1;
  actualSheet.insertRows(nextRowIndex + 1, 1);
  actualSheet.getRange(nextRowIndex + 1, 1)
    .setValue(formatWeekLabel(weekNumber, targetMonday))
    .setFontWeight('bold');
  return nextRowIndex;
}

function syncLastMonth() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const actualSheet = spreadsheet.getSheetByName('Actual Runs');
  const plannedSheet = spreadsheet.getSheetByName('Planned Schedule');
  if (!actualSheet) throw new Error('Sheet named "Actual Runs" was not found.');
  if (!plannedSheet) throw new Error('Sheet named "Planned Schedule" was not found.');

  ensureRunsHeader(actualSheet);
  const currentDate = new Date();
  const monthStart = sinceValue
    ? new Date(`${sinceValue}T00:00:00`)
    : getMonthStartDate(currentDate);
  const currentMonday = getMondayOnOrBefore(currentDate);
  const workoutSplitsStart = getCurrentYearCutoff(8, 4, currentDate);
  const supplementalWorkoutsStart = getCurrentYearCutoff(8, 15, currentDate);
  const activities = fetchStravaActivitiesSince(monthStart);
  const weekStarts = getCoveredWeekStarts(getMondayOnOrAfter(monthStart), currentMonday);
  const injuryReportSheet = getOrCreateInjuryReportSheet(spreadsheet, currentDate);

  weekStarts.forEach(targetMonday => {
    const targetRowIndex = ensureActualWeekRow(actualSheet, targetMonday);
    const plannedData = plannedSheet.getDataRange().getValues();
    const plannedRowIndex = findWeekRowIndex(plannedData, targetMonday);
    const plannedRow = plannedRowIndex === -1 ? null : plannedData[plannedRowIndex];
    const weekActivities = activities.filter(activity => {
      const dayOffset = getDayOffset(parseActivityDate(activity), targetMonday);
      return dayOffset >= 0 && dayOffset <= 6;
    });
    const todayOffset = targetMonday < currentMonday
      ? 6
      : Math.min(6, Math.max(0, getDayOffset(currentDate, targetMonday)));

    updateDailyCells(
      actualSheet,
      targetRowIndex,
      plannedRow,
      calculateDailyRunMiles(weekActivities, targetMonday),
      calculateDailyWorkouts(weekActivities, targetMonday),
      todayOffset,
      calculateDailySupplementalWorkouts(weekActivities, targetMonday)
    );
    applyDailyCellStyles(actualSheet, targetRowIndex, plannedRow, todayOffset);
    updateInjuryReportSheet(injuryReportSheet, weekActivities, targetMonday, currentDate);
    SpreadsheetApp.flush();
    updateTotalCell(actualSheet, targetRowIndex, plannedRow);
  });
  boldActualRunLabels(actualSheet);

  updateWorkoutSplitsSheet(
    spreadsheet,
    activities.filter(activity => parseActivityDate(activity) >= workoutSplitsStart)
  );
  updateSupplementalWorkoutsSheet(
    spreadsheet,
    activities.filter(activity => parseActivityDate(activity) >= supplementalWorkoutsStart)
  );
  SpreadsheetApp.flush();

  const changes = (globalThis.__localSheetsMock || sheetsMock).save();
  console.log(`Synced ${activities.length} activities across ${weekStarts.length} weeks.`);
  console.log(`Preserved Planned Schedule with ${plannedSheet.getDataRange().getValues().length} rows.`);
  console.log(`Saved ${changes.length} sheet changes to ${globalThis.__localOutputPath || outputPath}`);
}

vm.runInThisContext(scriptSource, { filename: 'apps-script-source.js' });

if (typeof require !== 'undefined' && typeof module !== 'undefined' && require.main === module) {
  if (process.argv.includes('--activity-mocks')) {
    globalThis.fetchStravaActivitiesSince = startDate => loadActivityMocks(startDate);
  }

  Promise.resolve(syncLastMonth()).catch(error => {
    console.error(error);
    process.exitCode = 1;
  });
}

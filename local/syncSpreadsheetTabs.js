const { readFileSync, writeFileSync } = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { google } = require('googleapis');
const { getAuthenticatedClient } = require('./googleAuth');
const { renderSpreadsheetHtml } = require('./SpreadsheetHtml');
const { createSpreadsheetApp } = require('./GoogleSheetsMock');

const ROOT = path.resolve(__dirname, '..');
const SPREADSHEET_NAME = '2026 Running Schedule';
const SPREADSHEET_ID = '1JM5PwBXCGw9HJiJbnPSTqhcg_ai1EZbGdwVKhf9m47E';
const fixturePath = path.join(ROOT, 'fixtures', 'spreadsheet.json');
const outputPath = path.join(ROOT, 'output', 'spreadsheet.json');
const outputHtmlPath = path.join(ROOT, 'output', 'spreadsheet.html');
const DATE_COLUMN_SHEETS = ['Injury Report', 'Supplemental Workouts', 'Workout Splits'];

function applyLocalBoldFormatting(workbook) {
  Object.values(workbook.sheets || {}).forEach(sheet => {
    sheet.styles ||= {};
    const rows = sheet.values || [];
    const columnCount = Math.max(1, ...rows.map(row => (row || []).length));

    for (let column = 1; column <= columnCount; column++) {
      sheet.styles[`1,${column}`] ||= {};
      sheet.styles[`1,${column}`].fontWeight = 'bold';
      sheet.styles[`1,${column}`].background = null;
      sheet.styles[`1,${column}`].fontColor = '#202124';
    }
    for (let row = 1; row <= rows.length; row++) {
      sheet.styles[`${row},1`] ||= {};
      sheet.styles[`${row},1`].fontWeight = 'bold';
    }
  });
}

function applyLocalDateFormatting(workbook) {
  DATE_COLUMN_SHEETS.forEach(sheetName => {
    const sheet = workbook.sheets?.[sheetName];
    if (!sheet) return;

    sheet.styles ||= {};
    const rows = sheet.values || [];
    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      if (rows[rowIndex]?.[0]) {
        sheet.styles[`${rowIndex + 1},1`] ||= {};
        sheet.styles[`${rowIndex + 1},1`].numberFormat = '@';
      }
    }
  });
}

async function readSpreadsheetTab(sheets, tabName) {
  const response = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `'${tabName}'`
  });
  return response.data.values || [];
}

// Pulled values carry no formatting, so repaint black separator rows using the real
// generator logic instead of leaving Supplemental Workouts looking broken locally.
function repaintSupplementalWorkoutsSeparators(workbook) {
  const sheet = workbook.sheets?.['Supplemental Workouts'];
  if (!sheet) return;

  const { SpreadsheetApp } = createSpreadsheetApp(workbook);
  const source = [
    readFileSync(path.join(ROOT, 'DateUtils.js'), 'utf8'),
    readFileSync(path.join(ROOT, 'InjuryReport.js'), 'utf8'),
    readFileSync(path.join(ROOT, 'SupplementalWorkouts.js'), 'utf8')
  ].join('\n');
  const context = vm.createContext({ console, SpreadsheetApp });
  vm.runInContext(
    `${source}\nrepairSupplementalWorkoutsSeparators(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUPPLEMENTAL_WORKOUTS_SHEET_NAME));`,
    context
  );
}

function repaintActualRunsStyles(workbook) {
  const actualSheetData = workbook.sheets?.['Actual Runs'];
  const plannedSheetData = workbook.sheets?.['Planned Schedule'];
  if (!actualSheetData || !plannedSheetData) return;

  const { SpreadsheetApp } = createSpreadsheetApp(workbook);
  const source = [
    readFileSync(path.join(ROOT, 'DateUtils.js'), 'utf8'),
    readFileSync(path.join(ROOT, 'RunningMetrics.js'), 'utf8'),
    readFileSync(path.join(ROOT, 'SheetUtils.js'), 'utf8')
  ].join('\n');
  const context = vm.createContext({ console, SpreadsheetApp });
  vm.runInContext(`${source}
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const actualSheet = spreadsheet.getSheetByName('Actual Runs');
    const plannedSheet = spreadsheet.getSheetByName('Planned Schedule');
    const actualData = actualSheet.getDataRange().getValues();
    const plannedData = plannedSheet.getDataRange().getValues();
    const currentMonday = getMostRecentMonday();
    const currentDayOffset = Math.min(6, Math.max(0, getDayOffset(new Date(), currentMonday)));

    actualData.forEach((row, rowIndex) => {
      if (rowIndex === 0) return;
      const weekStart = getWeekDate(row[0]);
      if (!weekStart || weekStart > currentMonday) return;

      const plannedRowIndex = findWeekRowIndex(plannedData, weekStart);
      if (plannedRowIndex === -1) return;
      const plannedRow = plannedData[plannedRowIndex];
      const isCurrentWeek = areSameDate(weekStart, currentMonday);
      const lastStyleDayIndex = isCurrentWeek ? currentDayOffset : 6;
      applyDailyCellStyles(actualSheet, rowIndex, plannedRow, lastStyleDayIndex);

      if (isCurrentWeek) {
        for (let dayIndex = currentDayOffset + 1; dayIndex < 7; dayIndex++) {
          applyCellStyle(actualSheet.getRange(rowIndex + 1, dayIndex + 2), {
            background: null,
            fontColor: null
          });
        }
      }

      const plannedTotalMiles = parseMilesFromCell(plannedRow[8]);
      if (weekStart < currentMonday && plannedTotalMiles > 0) {
        applyCellStyle(
          actualSheet.getRange(rowIndex + 1, 9),
          getTotalCellStyle(parseMilesFromCell(row[8]), plannedTotalMiles)
        );
      }
    });
  `, context);
}

function replaceWorkbookTab(workbook, tabName, values, preserveLocalFormatting = false) {
  const existingSheet = workbook.sheets[tabName];
  workbook.sheets[tabName] = preserveLocalFormatting && existingSheet
    ? { ...existingSheet, values }
    : { values, styles: {}, columnWidths: {}, rowHeights: {}, mergedRanges: [] };
}

async function syncSpreadsheetTabs(tabNames, options = {}) {
  const preserveLocalFormatting = new Set(options.preserveLocalFormatting || []);
  const auth = await getAuthenticatedClient();
  const sheets = google.sheets({ version: 'v4', auth });
  const spreadsheet = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    fields: 'properties.title'
  });
  const spreadsheetTitle = spreadsheet.data.properties?.title;
  if (spreadsheetTitle !== SPREADSHEET_NAME) {
    throw new Error(`Expected spreadsheet "${SPREADSHEET_NAME}" but found "${spreadsheetTitle || 'unknown'}".`);
  }

  const workbook = JSON.parse(readFileSync(fixturePath, 'utf8'));
  workbook.sheets ||= {};

  for (const tabName of tabNames) {
    const values = await readSpreadsheetTab(sheets, tabName);
    replaceWorkbookTab(workbook, tabName, values, preserveLocalFormatting.has(tabName));
    console.log(`Synced ${values.length} ${tabName} rows from ${SPREADSHEET_NAME}.`);
  }

  applyLocalBoldFormatting(workbook);
  applyLocalDateFormatting(workbook);
  if (tabNames.includes('Actual Runs')) repaintActualRunsStyles(workbook);
  if (tabNames.includes('Supplemental Workouts')) repaintSupplementalWorkoutsSeparators(workbook);

  const serializedWorkbook = `${JSON.stringify(workbook, null, 2)}\n`;
  writeFileSync(fixturePath, serializedWorkbook);
  writeFileSync(outputPath, serializedWorkbook);
  writeFileSync(outputHtmlPath, renderSpreadsheetHtml(workbook, []));
  console.log('Updated fixtures/spreadsheet.json, output/spreadsheet.json, and output/spreadsheet.html.');
}

function runSpreadsheetTabSync(tabNames, options) {
  syncSpreadsheetTabs(tabNames, options).catch(error => {
    console.error(error.message);
    process.exitCode = 1;
  });
}

module.exports = { runSpreadsheetTabSync, syncSpreadsheetTabs };

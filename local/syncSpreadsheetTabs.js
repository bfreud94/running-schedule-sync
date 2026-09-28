const { readFileSync, writeFileSync } = require('node:fs');
const path = require('node:path');
const { google } = require('googleapis');
const { getAuthenticatedClient } = require('./googleAuth');
const { renderSpreadsheetHtml } = require('./SpreadsheetHtml');

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

async function syncSpreadsheetTabs(tabNames) {
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
    // Row counts shift between production and local, so stale styling/merges would misalign.
    workbook.sheets[tabName] = { values, styles: {}, columnWidths: {}, rowHeights: {}, mergedRanges: [] };
    console.log(`Synced ${values.length} ${tabName} rows from ${SPREADSHEET_NAME}.`);
  }

  applyLocalBoldFormatting(workbook);
  applyLocalDateFormatting(workbook);

  const serializedWorkbook = `${JSON.stringify(workbook, null, 2)}\n`;
  writeFileSync(fixturePath, serializedWorkbook);
  writeFileSync(outputPath, serializedWorkbook);
  writeFileSync(outputHtmlPath, renderSpreadsheetHtml(workbook, []));
  console.log('Updated fixtures/spreadsheet.json, output/spreadsheet.json, and output/spreadsheet.html.');
}

function runSpreadsheetTabSync(tabNames) {
  syncSpreadsheetTabs(tabNames).catch(error => {
    console.error(error.message);
    process.exitCode = 1;
  });
}

module.exports = { runSpreadsheetTabSync, syncSpreadsheetTabs };

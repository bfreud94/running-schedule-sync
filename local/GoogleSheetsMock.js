const { mkdirSync, readFileSync, writeFileSync } = require('node:fs');
const path = require('node:path');
const { renderSpreadsheetHtml } = require('./SpreadsheetHtml');

function createGoogleSheetsMock(fixturePath, outputPath) {
  const workbook = JSON.parse(readFileSync(fixturePath, 'utf8'));
  const changes = [];

  function getSheet(name) {
    const sheet = workbook.sheets[name];
    if (!sheet) return null;

    sheet.styles ||= {};

    return {
      getDataRange: () => ({
        getValues: () => sheet.values.map(row => [...row])
      }),
      getRange(row, column) {
        const rowIndex = row - 1;
        const columnIndex = column - 1;
        sheet.values[rowIndex] ||= [];

        return {
          getValue: () => sheet.values[rowIndex][columnIndex],
          getBackground: () => sheet.styles[`${row},${column}`]?.background || '#ffffff',
          setValue(value) {
            const previousValue = sheet.values[rowIndex][columnIndex];
            sheet.values[rowIndex][columnIndex] = value;
            changes.push({ sheet: name, row, column, property: 'value', previousValue, value });
            return this;
          },
          setBackground(value) {
            sheet.styles[`${row},${column}`] ||= {};
            sheet.styles[`${row},${column}`].background = value;
            changes.push({ sheet: name, row, column, property: 'background', value });
            return this;
          },
          setFontColor(value) {
            sheet.styles[`${row},${column}`] ||= {};
            sheet.styles[`${row},${column}`].fontColor = value;
            changes.push({ sheet: name, row, column, property: 'fontColor', value });
            return this;
          },
          setNumberFormat(value) {
            sheet.styles[`${row},${column}`] ||= {};
            sheet.styles[`${row},${column}`].numberFormat = value;
            changes.push({ sheet: name, row, column, property: 'numberFormat', value });
            return this;
          }
        };
      }
    };
  }

  function recalculateTotals() {
    const actualRuns = workbook.sheets['Actual Runs'];
    if (!actualRuns) return;

    actualRuns.values.forEach(row => {
      const total = row.slice(1, 8).reduce((sum, value) => sum + (parseFloat(value) || 0), 0);
      row[8] = `${Math.floor(total * 100) / 100} miles`;
    });
  }

  return {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: getSheet }),
      flush: recalculateTotals
    },
    save() {
      mkdirSync(path.dirname(outputPath), { recursive: true });
      writeFileSync(outputPath, `${JSON.stringify({ ...workbook, changes }, null, 2)}\n`);
      const htmlPath = outputPath.replace(/\.json$/i, '.html');
      writeFileSync(htmlPath, renderSpreadsheetHtml(workbook, changes));
      return changes;
    }
  };
}

module.exports = { createGoogleSheetsMock };
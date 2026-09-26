const { mkdirSync, readFileSync, writeFileSync } = require('node:fs');
const path = require('node:path');
const { renderSpreadsheetHtml } = require('./SpreadsheetHtml');

const ISO_DATE_TIME_PATTERN = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d{3})?Z$/;

// JSON.stringify turns Date cells into ISO strings; revive them so they format as "September 26" again.
function reviveDates(key, value) {
  return typeof value === 'string' && ISO_DATE_TIME_PATTERN.test(value) ? new Date(value) : value;
}

function createGoogleSheetsMock(fixturePath, outputPath) {
  const workbook = JSON.parse(readFileSync(fixturePath, 'utf8'), reviveDates);
  const changes = [];

  function getSheet(name) {
    const sheet = workbook.sheets[name];
    if (!sheet) return null;

    sheet.styles ||= {};
    sheet.columnWidths ||= {};
    sheet.rowHeights ||= {};
    sheet.mergedRanges ||= [];

    return {
      setColumnWidth(column, width) {
        sheet.columnWidths[column] = width;
        changes.push({ sheet: name, column, property: 'columnWidth', value: width });
      },
      setRowHeight(row, height) {
        sheet.rowHeights[row] = height;
        changes.push({ sheet: name, row, property: 'rowHeight', value: height });
      },
      autoResizeRows(startRow, numRows) {
        for (let row = startRow; row < startRow + numRows; row++) {
          delete sheet.rowHeights[row];
          changes.push({ sheet: name, row, property: 'rowHeight', value: null });
        }
      },
      getLastRow() {
        for (let rowIndex = sheet.values.length - 1; rowIndex >= 0; rowIndex--) {
          if (sheet.values[rowIndex]?.some(value => String(value || '').trim() !== '')) return rowIndex + 1;
        }
        return 0;
      },
      deleteRows(startRow, numRows) {
        sheet.values.splice(startRow - 1, numRows);

        const shiftRow = row => {
          if (row >= startRow && row < startRow + numRows) return null;
          return row >= startRow + numRows ? row - numRows : row;
        };

        const shiftedStyles = {};
        Object.entries(sheet.styles).forEach(([key, value]) => {
          const [row, column] = key.split(',').map(Number);
          const shiftedRow = shiftRow(row);
          if (shiftedRow !== null) shiftedStyles[`${shiftedRow},${column}`] = value;
        });
        sheet.styles = shiftedStyles;

        const shiftedRowHeights = {};
        Object.entries(sheet.rowHeights).forEach(([rowKey, value]) => {
          const shiftedRow = shiftRow(Number(rowKey));
          if (shiftedRow !== null) shiftedRowHeights[shiftedRow] = value;
        });
        sheet.rowHeights = shiftedRowHeights;

        sheet.mergedRanges = sheet.mergedRanges
          .filter(range => shiftRow(range.row) !== null)
          .map(range => ({ ...range, row: shiftRow(range.row) }));

        changes.push({ sheet: name, row: startRow, property: 'deleteRows', value: numRows });
      },
      insertRows(startRow, numRows) {
        sheet.values.splice(startRow - 1, 0, ...Array.from({ length: numRows }, () => []));

        const shiftRow = row => (row >= startRow ? row + numRows : row);

        const shiftedStyles = {};
        Object.entries(sheet.styles).forEach(([key, value]) => {
          const [row, column] = key.split(',').map(Number);
          shiftedStyles[`${shiftRow(row)},${column}`] = value;
        });
        sheet.styles = shiftedStyles;

        const shiftedRowHeights = {};
        Object.entries(sheet.rowHeights).forEach(([rowKey, value]) => {
          shiftedRowHeights[shiftRow(Number(rowKey))] = value;
        });
        sheet.rowHeights = shiftedRowHeights;

        sheet.mergedRanges = sheet.mergedRanges.map(range => ({ ...range, row: shiftRow(range.row) }));

        changes.push({ sheet: name, row: startRow, property: 'insertRows', value: numRows });
      },
      getDataRange: () => ({
        getValues: () => sheet.values.map(row => [...(row || [])])
      }),
      getRange(row, column, numRows = 1, numColumns = 1) {
        const rowIndex = row - 1;
        const columnIndex = column - 1;

        return {
          getValue: () => sheet.values[rowIndex]?.[columnIndex],
          getBackground: () => sheet.styles[`${row},${column}`]?.background || '#ffffff',
          getFontFamily: () => sheet.styles[`${row},${column}`]?.fontFamily || 'Aptos',
          getFontSize: () => sheet.styles[`${row},${column}`]?.fontSize || 13,
          setValue(value) {
            sheet.values[rowIndex] ||= [];
            const previousValue = sheet.values[rowIndex][columnIndex];
            sheet.values[rowIndex][columnIndex] = value;
            changes.push({ sheet: name, row, column, property: 'value', previousValue, value });
            return this;
          },
          setValues(values) {
            values.forEach((valuesRow, valuesRowIndex) => {
              valuesRow.forEach((value, valuesColumnIndex) => {
                const targetRow = rowIndex + valuesRowIndex;
                const targetColumn = columnIndex + valuesColumnIndex;
                sheet.values[targetRow] ||= [];
                const previousValue = sheet.values[targetRow][targetColumn];
                sheet.values[targetRow][targetColumn] = value;
                changes.push({
                  sheet: name,
                  row: targetRow + 1,
                  column: targetColumn + 1,
                  property: 'value',
                  previousValue,
                  value
                });
              });
            });
            return this;
          },
          setBackground(value) {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = row + rowOffset;
                const targetColumn = column + columnOffset;
                sheet.styles[`${targetRow},${targetColumn}`] ||= {};
                sheet.styles[`${targetRow},${targetColumn}`].background = value;
                changes.push({ sheet: name, row: targetRow, column: targetColumn, property: 'background', value });
              }
            }
            return this;
          },
          setFontColor(value) {
            sheet.styles[`${row},${column}`] ||= {};
            sheet.styles[`${row},${column}`].fontColor = value;
            changes.push({ sheet: name, row, column, property: 'fontColor', value });
            return this;
          },
          setFontWeight(value) {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = row + rowOffset;
                const targetColumn = column + columnOffset;
                sheet.styles[`${targetRow},${targetColumn}`] ||= {};
                sheet.styles[`${targetRow},${targetColumn}`].fontWeight = value;
                changes.push({ sheet: name, row: targetRow, column: targetColumn, property: 'fontWeight', value });
              }
            }
            return this;
          },
          setFontFamily(value) {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = row + rowOffset;
                const targetColumn = column + columnOffset;
                sheet.styles[`${targetRow},${targetColumn}`] ||= {};
                if (value === null) delete sheet.styles[`${targetRow},${targetColumn}`].fontFamily;
                else sheet.styles[`${targetRow},${targetColumn}`].fontFamily = value;
              }
            }
            return this;
          },
          setFontSize(value) {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = row + rowOffset;
                const targetColumn = column + columnOffset;
                sheet.styles[`${targetRow},${targetColumn}`] ||= {};
                if (value === null) delete sheet.styles[`${targetRow},${targetColumn}`].fontSize;
                else sheet.styles[`${targetRow},${targetColumn}`].fontSize = value;
              }
            }
            return this;
          },
          setNumberFormat(value) {
            sheet.styles[`${row},${column}`] ||= {};
            sheet.styles[`${row},${column}`].numberFormat = value;
            changes.push({ sheet: name, row, column, property: 'numberFormat', value });
            return this;
          },
          setWrap(value) {
            sheet.styles[`${row},${column}`] ||= {};
            sheet.styles[`${row},${column}`].wrap = value;
            changes.push({ sheet: name, row, column, property: 'wrap', value });
            return this;
          },
          setVerticalAlignment(value) {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = row + rowOffset;
                const targetColumn = column + columnOffset;
                sheet.styles[`${targetRow},${targetColumn}`] ||= {};
                sheet.styles[`${targetRow},${targetColumn}`].verticalAlignment = value;
                changes.push({ sheet: name, row: targetRow, column: targetColumn, property: 'verticalAlignment', value });
              }
            }
            return this;
          },
          merge() {
            const mergedRange = { row, column, numRows, numColumns };
            sheet.mergedRanges = sheet.mergedRanges.filter(existing =>
              existing.row !== row || existing.column !== column
            );
            sheet.mergedRanges.push(mergedRange);
            changes.push({ sheet: name, property: 'mergedRange', value: mergedRange });
            return this;
          },
          breakApart() {
            sheet.mergedRanges = sheet.mergedRanges.filter(existing =>
              existing.row < row || existing.row >= row + numRows ||
              existing.column < column || existing.column >= column + numColumns
            );
            return this;
          },
          clearContent() {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = rowIndex + rowOffset;
                const targetColumn = columnIndex + columnOffset;
                sheet.values[targetRow] ||= [];
                sheet.values[targetRow][targetColumn] = '';
              }
            }
            return this;
          },
          setHorizontalAlignment(value) {
            for (let rowOffset = 0; rowOffset < numRows; rowOffset++) {
              for (let columnOffset = 0; columnOffset < numColumns; columnOffset++) {
                const targetRow = row + rowOffset;
                const targetColumn = column + columnOffset;
                sheet.styles[`${targetRow},${targetColumn}`] ||= {};
                sheet.styles[`${targetRow},${targetColumn}`].horizontalAlignment = value;
                changes.push({ sheet: name, row: targetRow, column: targetColumn, property: 'horizontalAlignment', value });
              }
            }
            return this;
          }
        };
      }
    };
  }

  function recalculateTotals() {
    const actualRuns = workbook.sheets['Actual Runs'];
    if (!actualRuns) return;

    actualRuns.values.forEach((row, rowIndex) => {
      if (rowIndex === 0 && String(row[1] || '').trim().toLowerCase() === 'monday') return;
      const total = row.slice(1, 8).reduce((sum, value) => sum + (parseFloat(value) || 0), 0);
      row[8] = `${Math.floor(total * 100) / 100} miles`;
    });
  }

  function insertSheet(name) {
    workbook.sheets[name] = { values: [], styles: {} };
    changes.push({ sheet: name, property: 'created' });
    return getSheet(name);
  }

  return {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: getSheet,
        insertSheet
      }),
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
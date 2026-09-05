const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('adds workout text after the daily miles value', () => {
  let writtenValue;
  const sheet = {
    getRange() {
      return {
        getValue: () => '',
        setValue: value => { writtenValue = value; },
        setBackground: () => {},
        setFontColor: () => {}
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.updateTarget = updateDailyCells;`, context);
  context.updateTarget(sheet, 0, null, [5], [['10 min easy']], 0);

  assert.equal(writtenValue, '5 miles\n10 min easy');
});

test('finds the planned row by week instead of row position', () => {
  const context = vm.createContext({ console });
  const dateSource = readFileSync('DateUtils.js', 'utf8');
  const sheetSource = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${dateSource}\n${sheetSource}\nglobalThis.weekRowTarget = findWeekRowIndex;`, context);

  const rows = [
    ['Week', 'Monday', 'Tuesday'],
    ['WEEK 0 (8/24)', '1 mile', 'Rest'],
    ['WEEK 1 (8/31)', '3 miles', '4 miles']
  ];
  assert.equal(context.weekRowTarget(rows, new Date(2026, 7, 31)), 2);
});
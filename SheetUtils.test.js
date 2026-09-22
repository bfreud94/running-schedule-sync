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

test('backfills workout text into an existing miles cell', () => {
  let writtenValue;
  const sheet = {
    getRange() {
      return {
        getValue: () => '6.21 miles',
        setValue: value => { writtenValue = value; },
        setBackground: () => {},
        setFontColor: () => {}
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.updateTarget = updateDailyCells;`, context);
  context.updateTarget(sheet, 0, null, [6.21, 6.21], [[], ['1000m repeats']], 1);

  assert.equal(writtenValue, '6.21 miles\n1000m repeats');
});

test('removes bold formatting when writing the current day cell', () => {
  let fontWeight;
  const sheet = {
    getRange() {
      return {
        getValue: () => '',
        setValue: () => {},
        setFontWeight: value => { fontWeight = value; },
        setBackground: () => {},
        setFontColor: () => {}
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.updateTarget = updateDailyCells;`, context);
  context.updateTarget(sheet, 0, null, [5], [['10 min easy']], 0);

  assert.equal(fontWeight, 'normal');
});

test('calculates total mileage from the current row daily cells', () => {
  const values = {
    '2,2': '4 miles',
    '2,3': '5.25 miles',
    '2,4': 'Rest',
    '2,5': '3 miles',
    '2,6': '6 miles',
    '2,7': 'Rest',
    '2,8': '7 miles',
    '2,9': '999 miles',
    '1,2': '100 miles'
  };
  let totalValue;
  const sheet = {
    getRange: (row, column) => ({
      getValue: () => values[`${row},${column}`] || '',
      setValue: value => { if (row === 2 && column === 9) totalValue = value; },
      setBackground: () => {},
      setFontColor: () => {}
    })
  };
  const context = vm.createContext({ console });
  const source = `${readFileSync('RunningMetrics.js', 'utf8')}\n${readFileSync('SheetUtils.js', 'utf8')}`;

  vm.runInContext(`${source}\nglobalThis.totalTarget = updateTotalCell;`, context);
  assert.equal(context.totalTarget(sheet, 1, null), 25.25);
  assert.equal(totalValue, '25.25 miles');
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

test('appends a new week row when the current week is missing', () => {
  const writes = [];
  const sheet = {
    getRange: (row, column) => ({
      getValue: () => '',
      setValue: value => { writes.push({ row, column, value }); }
    })
  };
  const context = vm.createContext({ console });
  const dateSource = readFileSync('DateUtils.js', 'utf8');
  const sheetSource = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${dateSource}\n${sheetSource}\nglobalThis.appendTarget = appendWeekRow;`, context);

  const rows = [
    ['WEEK 1 (8/31)', '3 miles'],
    ['WEEK 2 (9/7)', '4 miles']
  ];
  assert.equal(context.appendTarget(sheet, rows, new Date(2026, 8, 14)), 2);
  assert.deepEqual(writes, [{ row: 3, column: 1, value: 'WEEK 3 (9/14)' }]);
});

test('appends the new week row directly below the last populated row, ignoring trailing blanks', () => {
  const writes = [];
  const sheet = {
    getRange: (row, column) => ({
      getValue: () => '',
      setValue: value => { writes.push({ row, column, value }); }
    })
  };
  const context = vm.createContext({ console });
  const dateSource = readFileSync('DateUtils.js', 'utf8');
  const sheetSource = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${dateSource}\n${sheetSource}\nglobalThis.appendTarget = appendWeekRow;`, context);

  const rows = [
    ['WEEK 1 (8/31)', '3 miles'],
    ['WEEK 2 (9/7)', '4 miles'],
    ['', ''],
    ['', ''],
    ['', '']
  ];
  assert.equal(context.appendTarget(sheet, rows, new Date(2026, 8, 14)), 2);
  assert.deepEqual(writes, [{ row: 3, column: 1, value: 'WEEK 3 (9/14)' }]);
});
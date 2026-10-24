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
  let totalFormula;
  const sheet = {
    getRange: (row, column) => ({
      getValue: () => values[`${row},${column}`] || '',
      setFormula: formula => { if (row === 2 && column === 9) totalFormula = formula; },
      setBackground: () => {},
      setFontColor: () => {}
    })
  };
  const context = vm.createContext({ console });
  const source = `${readFileSync('RunningMetrics.js', 'utf8')}\n${readFileSync('SheetUtils.js', 'utf8')}`;

  vm.runInContext(`${source}\nglobalThis.totalTarget = updateTotalCell;`, context);
  assert.equal(context.totalTarget(sheet, 1, null), 25.25);
  assert.equal(totalFormula, `=SUMPRODUCT(IFERROR(VALUE(REGEXEXTRACT(B2:H2, "[\\d\\.]+")))) & " miles"`);
});

test('styles existing daily values without rewriting them', () => {
  const styles = [];
  const sheet = {
    getRange: (row, column) => ({
      getValue: () => column === 2 ? '4 miles' : '',
      setBackground: value => { styles.push({ row, column, property: 'background', value }); },
      setFontColor: value => { styles.push({ row, column, property: 'fontColor', value }); }
    })
  };
  const context = vm.createContext({ console });
  const source = `${readFileSync('RunningMetrics.js', 'utf8')}\n${readFileSync('SheetUtils.js', 'utf8')}`;

  vm.runInContext(`${source}\nglobalThis.styleTarget = applyDailyCellStyles;`, context);
  context.styleTarget(sheet, 0, ['Week', '4 miles'], 0);

  assert.deepEqual(styles, [
    { row: 1, column: 2, property: 'background', value: '#34a853' },
    { row: 1, column: 2, property: 'fontColor', value: '#ffffff' }
  ]);
});

test('colors the weekly total by percentage off the planned total', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('SheetUtils.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.totalStyleTarget = getTotalCellStyle;`, context);

  const green = { background: '#34a853', fontColor: '#ffffff' };
  const yellow = { background: 'yellow', fontColor: null };
  const red = { background: 'red', fontColor: '#ffffff' };

  // 4% off is green now that the green band runs to 5% (it was yellow under the old 3% cutoff).
  assert.deepEqual({ ...context.totalStyleTarget(48, 50) }, green);
  // Exactly 5% off is still within the green band.
  assert.deepEqual({ ...context.totalStyleTarget(47.5, 50) }, green);
  // Between 5% and 10% off is yellow/black, over or under the target.
  assert.deepEqual({ ...context.totalStyleTarget(47, 50) }, yellow);
  assert.deepEqual({ ...context.totalStyleTarget(55, 50) }, yellow);
  // More than 10% off is red/white.
  assert.deepEqual({ ...context.totalStyleTarget(44, 50) }, red);
});

test('colors each daily run by absolute miles off the plan', () => {
  const context = vm.createContext({ console });
  const source = `${readFileSync('RunningMetrics.js', 'utf8')}\n${readFileSync('SheetUtils.js', 'utf8')}`;

  vm.runInContext(`${source}\nglobalThis.dailyStyleTarget = getDailyCellStyle;`, context);

  const green = { background: '#34a853', fontColor: '#ffffff' };
  const yellow = { background: 'yellow', fontColor: null };
  const red = { background: 'red', fontColor: '#ffffff' };

  // 0.4 mile off is green now that the green band runs to 1/2 mile (it was yellow under 1/4 mile).
  assert.deepEqual({ ...context.dailyStyleTarget('10 miles', 10.4) }, green);
  // Exactly 1/2 mile off is still within the green band.
  assert.deepEqual({ ...context.dailyStyleTarget('10 miles', 10.5) }, green);
  // 1/2 to 1 mile off is yellow/black (0.75 was red under the old 1/2 mile cutoff).
  assert.deepEqual({ ...context.dailyStyleTarget('10 miles', 10.75) }, yellow);
  assert.deepEqual({ ...context.dailyStyleTarget('10 miles', 11) }, yellow);
  // More than 1 mile off is red/white.
  assert.deepEqual({ ...context.dailyStyleTarget('10 miles', 11.5) }, red);
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
    getRange(row, column) {
      return {
      getValue: () => '',
        setValue(value) { writes.push({ row, column, value }); return this; },
        setFontWeight(value) { writes.push({ row, column, property: 'fontWeight', value }); return this; }
      };
    }
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
  assert.deepEqual(writes, [
    { row: 3, column: 1, value: 'WEEK 3 (9/14)' },
    { row: 3, column: 1, property: 'fontWeight', value: 'bold' }
  ]);
});

test('appends the new week row directly below the last populated row, ignoring trailing blanks', () => {
  const writes = [];
  const sheet = {
    getRange(row, column) {
      return {
        getValue: () => '',
        setValue(value) { writes.push({ row, column, value }); return this; },
        setFontWeight(value) { writes.push({ row, column, property: 'fontWeight', value }); return this; }
      };
    }
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
  assert.deepEqual(writes, [
    { row: 3, column: 1, value: 'WEEK 3 (9/14)' },
    { row: 3, column: 1, property: 'fontWeight', value: 'bold' }
  ]);
});
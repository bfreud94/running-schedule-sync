const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('parses workout splits until the first blank line and calculates mile pace', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.splitsTarget = parseWorkoutSplits;`, context);

  const splits = context.splitsTarget(
    'Workout:\n1000m repeats\nSplits:\n4:47\n4:31\n4:45\n4:32\n\nInjury Report:\nArea: Left Groin',
    '1000m repeats'
  );

  assert.equal(splits.length, 4);
  assert.equal(splits[0].value, '4:47');
  assert.equal(splits[0].pace, '7:42');
  assert.equal(splits[3].pace, '7:18');
  assert.equal(
    context.splitsTarget('Workout:\n6 mile 1000m repeats\nSplits:\nRep 1: 4:47\nRep 2: 4:31', '6 mile 1000m repeats')[0].pace,
    '7:42'
  );
  assert.equal(
    context.splitsTarget('Workout:\n1000m repeats\nSplits:\n4:47 avg\n\nInjury Report:', '1000m repeats')[0].value,
    '4:47'
  );
  assert.equal(context.splitsTarget('Workout:\nEasy run\nNo Splits section', 'Easy run').length, 0);
});

test('falls back to the split value when no pace is available', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.splitsTarget = parseWorkoutSplits;`, context);

  const splits = context.splitsTarget('Workout:\nTempo intervals\nSplits:\n6:45\n6:32', 'Tempo intervals');

  assert.equal(splits[0].pace, '6:45');
  assert.equal(splits[1].pace, '6:32');
});

test('adds 10px black separator rows around each workout day', () => {
  const values = [
    ['Date', 'Workout', 'Splits', 'Splits (Pace)'],
    [new Date(2026, 8, 4), '1000m repeats', '4:47', '7:42'],
    ['', '', '4:31', '7:16'],
    [new Date(2026, 8, 11), 'Mile repeats', '7:11', '7:11'],
    ['', '', '7:14', '7:14']
  ];
  const backgrounds = {};
  const rowHeights = {};
  const sheet = {
    insertRows(row, count) {
      values.splice(row - 1, 0, ...Array.from({ length: count }, () => []));
    },
    getDataRange: () => ({ getValues: () => values }),
    getLastRow() {
      for (let rowIndex = values.length - 1; rowIndex >= 0; rowIndex--) {
        if (values[rowIndex]?.some(value => String(value || '').trim())) return rowIndex + 1;
      }
      return 0;
    },
    getRange: (row, column) => ({
      setBackground: value => { backgrounds[`${row},${column}`] = value; }
    }),
    setRowHeight: (row, height) => { rowHeights[row] = height; }
  };
  const context = vm.createContext({
    console,
    parseSheetDate: value => value instanceof Date ? value : null
  });
  const source = readFileSync('WorkoutSplits.js', 'utf8');
  vm.runInContext(`${source}\nglobalThis.separatorTarget = ensureWorkoutSplitSeparatorRows;`, context);

  context.separatorTarget(sheet, values);

  assert.equal(values.length, 8);
  assert.deepEqual([2, 5, 8].map(row => ({
    row,
    values: values[row - 1],
    backgrounds: [1, 2, 3, 4].map(column => backgrounds[`${row},${column}`]),
    height: rowHeights[row]
  })), [
    { row: 2, values: [], backgrounds: ['#000000', '#000000', '#000000', '#000000'], height: 10 },
    { row: 5, values: [], backgrounds: ['#000000', '#000000', '#000000', '#000000'], height: 10 },
    { row: 8, values: [], backgrounds: ['#000000', '#000000', '#000000', '#000000'], height: 10 }
  ]);
});

test('writes the date as plain text formatted as "Month Day" instead of a Date-typed cell', () => {
  const values = [];
  const numberFormats = [];
  const sheet = {
    autoResizeRows: () => {},
    setRowHeight: () => {},
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: rowValues => { values.push({ row, column, rowValues: rowValues[0] }); },
        setNumberFormat: value => { numberFormats.push({ row, column, value }); },
        setFontWeight: () => {},
        setBackground: () => {},
        getBackground: () => '#ffffff',
        breakApart: () => {},
        merge: () => ({ setVerticalAlignment: () => {} })
      };
    }
  };
  const context = vm.createContext({
    console,
    getDateKey: date => `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
    parseSheetDate: value => value instanceof Date ? value : null
  });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.writeTarget = writeWorkoutSplitRows;`, context);
  context.writeTarget(
    sheet,
    [['Date', 'Workout', 'Splits', 'Splits (Pace)']],
    new Date(2026, 8, 24),
    '1000m repeats',
    [{ value: '4:47', pace: '7:42' }]
  );

  assert.equal(values[0].rowValues[0], 'September 24');
  assert.deepEqual(numberFormats, [{ row: 2, column: 1, value: '@' }]);
});

test('matches an existing plain-text date row regardless of year, overwriting instead of duplicating', () => {
  const written = [];
  const sheet = {
    autoResizeRows: () => {},
    setRowHeight: () => {},
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: values => { written.push({ row, column, values: values[0] }); },
        setNumberFormat: () => {},
        setFontWeight: () => {},
        setBackground: () => {},
        getBackground: () => '#ffffff',
        breakApart: () => {},
        merge: () => ({ setVerticalAlignment: () => {} })
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.writeTarget = writeWorkoutSplitRows;`, context);
  context.writeTarget(
    sheet,
    [
      ['Date', 'Workout', 'Splits', 'Splits (Pace)'],
      ['September 24', '1000m repeats', '4:47', '7:42']
    ],
    new Date(2026, 8, 24),
    '1000m repeats',
    [{ value: '4:47', pace: '7:42' }]
  );

  assert.equal(written.find(entry => entry.row === 2 && entry.column === 1)?.values[0], 'September 24');
  assert.equal(written.some(entry => entry.row === 3), false);
});

test('removes duplicate day blocks, keeping the most recent one', () => {
  const deletedRows = [];
  const sheetData = [
    ['Date', 'Workout', 'Splits', 'Splits (Pace)'],
    ['September 24', '1000m repeats', '4:47', '7:42'],
    [],
    ['September 25', 'Easy run', '3:00', ''],
    [],
    ['September 24', '1000m repeats', '4:47', '7:42'],
    []
  ];
  const sheet = {
    deleteRows: (row, count) => { deletedRows.push({ row, count }); }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.dedupeTarget = removeDuplicateWorkoutSplitBlocks;`, context);
  const removedAny = context.dedupeTarget(sheet, sheetData);

  assert.equal(removedAny, true);
  assert.deepEqual(deletedRows, [{ row: 2, count: 2 }]);
});

test('bolds the date cell when writing a workout split row', () => {
  const fontWeights = [];
  const sheet = {
    autoResizeRows: () => {},
    setRowHeight: () => {},
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: () => ({ setNumberFormat: () => {} }),
        setNumberFormat: () => {},
        setFontWeight: value => { fontWeights.push({ row, column, value }); },
        setBackground: () => {},
        getBackground: () => '#ffffff',
        breakApart: () => {},
        merge: () => ({ setVerticalAlignment: () => {} })
      };
    }
  };
  const context = vm.createContext({
    console,
    getDateKey: date => `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
    parseSheetDate: value => value instanceof Date ? value : null
  });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.writeTarget = writeWorkoutSplitRows;`, context);
  context.writeTarget(
    sheet,
    [['Date', 'Workout', 'Splits', 'Splits (Pace)']],
    new Date(2026, 8, 4),
    '1000m repeats',
    [{ value: '4:47', pace: '7:42' }]
  );

  assert.deepEqual(fontWeights, [{ row: 2, column: 1, value: 'bold' }]);
});

test('merges date and workout cells across the split rows', () => {
  const mergedRanges = [];
  const backgrounds = [];
  const autoResizedRows = [];
  const rowHeights = [];
  const sheet = {
    autoResizeRows: (row, count) => { autoResizedRows.push({ row, count }); },
    setRowHeight: (row, height) => { rowHeights.push({ row, height }); },
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: () => ({ setNumberFormat: () => {} }),
        setNumberFormat: () => {},
        setFontWeight: () => {},
        setBackground: value => { backgrounds.push({ row, column, value }); },
        getBackground: () => '#ffffff',
        breakApart: () => {},
        merge: () => {
          mergedRanges.push({ row, column, numRows, numColumns });
          return { setVerticalAlignment: () => {} };
        }
      };
    }
  };
  const context = vm.createContext({
    console,
    getDateKey: date => `${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`,
    parseSheetDate: value => value instanceof Date ? value : null
  });
  const source = readFileSync('WorkoutSplits.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.writeTarget = writeWorkoutSplitRows;`, context);
  context.writeTarget(
    sheet,
    [['Date', 'Workout', 'Splits', 'Splits (Pace)']],
    new Date(2026, 8, 4),
    '1000m repeats',
    [{ value: '4:47', pace: '7:42' }, { value: '4:31', pace: '7:16' }, { value: '4:45', pace: '7:39' }, { value: '4:32', pace: '7:18' }]
  );

  assert.deepEqual(mergedRanges, [
    { row: 2, column: 1, numRows: 4, numColumns: 1 },
    { row: 2, column: 2, numRows: 4, numColumns: 1 }
  ]);
  assert.deepEqual(backgrounds, [
    { row: 2, column: 1, value: null },
    { row: 3, column: 1, value: null },
    { row: 4, column: 1, value: null },
    { row: 5, column: 1, value: null },
    { row: 6, column: 1, value: '#000000' },
    { row: 6, column: 2, value: '#000000' },
    { row: 6, column: 3, value: '#000000' },
    { row: 6, column: 4, value: '#000000' }
  ]);
  assert.deepEqual(autoResizedRows, [
    { row: 2, count: 1 },
    { row: 3, count: 1 },
    { row: 4, count: 1 },
    { row: 5, count: 1 }
  ]);
  assert.deepEqual(rowHeights, [{ row: 6, height: 10 }]);
});
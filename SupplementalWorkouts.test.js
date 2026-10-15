const assert = require('node:assert/strict');
const { mkdtempSync, readFileSync, rmSync, writeFileSync } = require('node:fs');
const { tmpdir } = require('node:os');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const { createGoogleSheetsMock } = require('./local/GoogleSheetsMock');

function loadContext() {
  const context = vm.createContext({
    console,
    getDateKey: date => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`,
    parseSheetDate: value => value instanceof Date ? value : null
  });
  const source = `${readFileSync('DateUtils.js', 'utf8')}\n${readFileSync('SheetUtils.js', 'utf8')}\n${readFileSync('SupplementalWorkouts.js', 'utf8')}`;
  vm.runInContext(source, context);
  return context;
}

test('fills E:G for rows containing a black separator cell in A:D', () => {
  const backgrounds = { '3,2': '#000000', '4,1': '#000000' };
  const fills = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        getBackground: () => backgrounds[`${row},${column}`] || '#ffffff',
        setBackground: value => {
          fills.push({ row, column, numRows, numColumns, value });
          return this;
        }
      };
    }
  };
  const context = loadContext();
  vm.runInContext('globalThis.fillTarget = fillMissingSeparatorColumns;', context);

  context.fillTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    ['2026-09-15', 'Core', 'Planks', '1', '1:00', 'N/A', ''],
    ['', '', '', '', '', '', '']
  ]);

  assert.deepEqual(fills, [
    { row: 3, column: 5, numRows: 1, numColumns: 3, value: '#000000' },
    { row: 4, column: 5, numRows: 1, numColumns: 3, value: '#000000' }
  ]);
});

test('copies row 3 font family and size across the Supplemental Workouts sheet', () => {
  const calls = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        getFontFamily: () => row === 3 && column === 1 ? 'Reference Font' : undefined,
        getFontSize: () => row === 3 && column === 1 ? 11 : undefined,
        setFontFamily: value => { calls.push({ property: 'fontFamily', value, row, column, numRows, numColumns }); return this; },
        setFontSize: value => { calls.push({ property: 'fontSize', value, row, column, numRows, numColumns }); return this; }
      };
    }
  };
  const context = loadContext();
  vm.runInContext('globalThis.fontTarget = matchSupplementalFontToRowThree;', context);

  context.fontTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [],
    ['2026-09-22', 'Core', 'Planks', '2', '10', 'N/A', '']
  ]);

  assert.deepEqual(calls, [
    { property: 'fontFamily', value: 'Reference Font', row: 1, column: 1, numRows: 3, numColumns: 7 },
    { property: 'fontSize', value: 11, row: 1, column: 1, numRows: 3, numColumns: 7 }
  ]);
});

test('restores normal row height when workout data replaces a separator', () => {
  const resizedRows = [];
  const backgrounds = [];
  const sheet = {
    autoResizeRows: (row, count) => { resizedRows.push({ row, count }); },
    getRange: (row, column, numRows, numColumns) => ({
      setBackground: value => { backgrounds.push({ row, column, numRows, numColumns, value }); }
    })
  };
  const context = loadContext();
  vm.runInContext('globalThis.clearWorkoutRowsTarget = clearBlackFromWorkoutRows;', context);

  context.clearWorkoutRowsTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [],
    [new Date(2026, 8, 24), 'Upper Body Lift', 'Bench Press', '3', '8', '80 lbs', '']
  ]);

  assert.deepEqual(resizedRows, [{ row: 3, count: 1 }]);
  assert.deepEqual(backgrounds, [
    { row: 3, column: 1, numRows: 1, numColumns: 7, value: null }
  ]);
});

test('adds a separator after the final day block when it is missing', () => {
  const writes = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        getBackground: () => '#ffffff',
        setBackground: value => { writes.push({ row, column, numRows, numColumns, value }); return this; }
      };
    },
    setRowHeight: (row, height) => { writes.push({ row, height }); }
  };
  const context = loadContext();
  vm.runInContext('globalThis.separatorTarget = ensureDaySeparatorRows;', context);

  context.separatorTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [],
    ['2026-09-22', 'Upper Body Lift', 'Bench Press', '3', '10\n8\n7', '80 lbs\n85 lbs\n90 lbs', ''],
    ['', '', 'Accessory Exercise', '2', '10', 'N/A', '']
  ]);

  assert.deepEqual(writes, [
    { row: 5, column: 1, numRows: 1, numColumns: 7, value: '#000000' },
    { row: 5, height: 10 }
  ]);
});

test('does not paint a stray extra blank row that falls before the first day block', () => {
  const writes = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        getBackground: () => '#ffffff',
        setBackground: value => { writes.push({ row, column, numRows, numColumns, value }); return this; }
      };
    },
    setRowHeight: (row, height) => { writes.push({ row, height }); }
  };
  const context = loadContext();
  vm.runInContext('globalThis.separatorTarget = ensureDaySeparatorRows;', context);

  context.separatorTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [],
    [],
    ['2026-09-15', 'Core', 'Planks', '1', '1:00', 'N/A', '']
  ]);

  assert.deepEqual(writes, [
    { row: 5, column: 1, numRows: 1, numColumns: 7, value: '#000000' },
    { row: 5, height: 10 }
  ]);
});

test('clears a stray black row left over from a prior bug, leaving real separators alone', () => {
  const backgrounds = { '3,1': '#000000', '5,1': '#000000' };
  const writes = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        getBackground: () => backgrounds[`${row},${column}`] || '#ffffff',
        setBackground: value => { writes.push({ row, column, numRows, numColumns, value }); return this; }
      };
    }
  };
  const context = loadContext();
  vm.runInContext('globalThis.clearStrayTarget = clearUnexpectedSeparatorRows;', context);

  context.clearStrayTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [],
    [],
    ['2026-09-15', 'Core', 'Planks', '1', '1:00', 'N/A', ''],
    []
  ]);

  assert.deepEqual(writes, [
    { row: 3, column: 1, numRows: 1, numColumns: 7, value: null }
  ]);
});

test('reuses an existing blank header separator row instead of inserting a duplicate', () => {
  const inserted = [];
  const sheet = {
    insertRows: (row, count) => { inserted.push({ row, count }); },
    getRange: () => ({ setBackground: () => {} }),
    setRowHeight: () => {},
    getDataRange: () => ({ getValues: () => [['Date', 'Workout'], []] })
  };
  const context = loadContext();
  vm.runInContext('globalThis.headerSeparatorTarget = ensureHeaderSeparator;', context);

  context.headerSeparatorTarget(sheet, [['Date', 'Workout'], []]);

  assert.deepEqual(inserted, []);
});

test('inserts a header separator row when row 2 already holds real data', () => {
  const inserted = [];
  const sheet = {
    insertRows: (row, count) => { inserted.push({ row, count }); },
    getRange: () => ({ setBackground: () => {} }),
    setRowHeight: () => {},
    getDataRange: () => ({ getValues: () => [['Date', 'Workout'], ['2026-09-15', 'Core']] })
  };
  const context = loadContext();
  vm.runInContext('globalThis.headerSeparatorTarget = ensureHeaderSeparator;', context);

  context.headerSeparatorTarget(sheet, [['Date', 'Workout'], ['2026-09-15', 'Core']]);

  assert.deepEqual(inserted, [{ row: 2, count: 1 }]);
});

test('does not insert a duplicate final separator row when it already exists blank but unpainted', () => {
  const writes = [];
  const sheet = {
    getLastRow: () => 3,
    getDataRange: () => ({ getValues: () => [['Date', 'Workout'], ['2026-09-15', 'Core'], ['', ''], []] }),
    getRange: (row, column, numRows = 1, numColumns = 1) => ({
      setBackground: value => { writes.push({ row, column, numRows, numColumns, value }); }
    }),
    insertRows: (row, count) => { writes.push({ row, count }); },
    setRowHeight: (row, height) => { writes.push({ row, height }); }
  };
  const context = loadContext();
  vm.runInContext('globalThis.finalSeparatorTarget = ensureFinalSeparatorRow;', context);

  context.finalSeparatorTarget(sheet);

  assert.deepEqual(writes, [
    { row: 4, column: 1, numRows: 1, numColumns: 7, value: '#000000' },
    { row: 4, height: 10 }
  ]);
});

test('removes a duplicated blank row, collapsing consecutive blanks down to one', () => {
  const deletions = [];
  const sheet = {
    deleteRows: (row, count) => { deletions.push({ row, count }); }
  };
  const context = loadContext();
  vm.runInContext('globalThis.removeBlanksTarget = removeConsecutiveBlankRows;', context);

  context.removeBlanksTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [],
    [],
    ['2026-09-15', 'Core', 'Planks', '1', '1:00', 'N/A', '']
  ]);

  assert.deepEqual(deletions, [{ row: 3, count: 1 }]);
});

test('inserts the final separator row when its formatting is outside the data range', () => {
  const writes = [];
  const sheet = {
    getLastRow: () => 4,
    getDataRange: () => ({ getValues: () => [[], [], [], []] }),
    getRange: (row, column, numRows = 1, numColumns = 1) => ({
      getBackground: () => '#000000',
      setBackground: value => { writes.push({ row, column, numRows, numColumns, value }); }
    }),
    insertRows: (row, count) => { writes.push({ row, count }); },
    setRowHeight: (row, height) => { writes.push({ row, height }); }
  };
  const context = loadContext();
  vm.runInContext('globalThis.finalSeparatorTarget = ensureFinalSeparatorRow;', context);

  context.finalSeparatorTarget(sheet);

  assert.deepEqual(writes, [
    { row: 5, count: 1 },
    { row: 5, column: 1, numRows: 1, numColumns: 7, value: '#000000' },
    { row: 5, height: 10 }
  ]);
});

test('forces the known production separator row across all seven columns', () => {
  const fills = [];
  let cleared = false;
  const sheet = {
    getRange: (row, column, numRows = 1, numColumns = 1) => ({
      breakApart: () => {},
      clearContent: () => { cleared = true; },
      setBackground: value => { fills.push({ row, column, numRows, numColumns, value }); return this; }
    }),
    setRowHeight: (row, height) => { fills.push({ row, height }); }
  };
  const context = loadContext();
  vm.runInContext('globalThis.forceTarget = forceKnownSeparatorRow;', context);

  context.forceTarget(sheet, 12);

  assert.equal(cleared, true);
  assert.deepEqual(fills, [
    { row: 12, column: 1, numRows: 1, numColumns: 7, value: '#000000' },
    { row: 12, height: 10 }
  ]);
});

test('groups exercises by category, preserving first-seen order', () => {
  const context = loadContext();
  vm.runInContext('globalThis.groupTarget = groupExercisesByCategory;', context);

  const groups = context.groupTarget([
    { category: 'Core', workout: 'Planks', sets: '2', repsHoldTime: '1:30', weight: 'N/A', notes: '' },
    { category: 'Upper Body', workout: 'Bench Press', sets: '3', repsHoldTime: '10', weight: '135 lbs', notes: '' },
    { category: 'Core', workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', weight: 'N/A', notes: 'each side' }
  ]);

  assert.equal(groups.length, 2);
  assert.equal(groups[0].category, 'Core');
  assert.deepEqual([...groups[0].exercises.map(exercise => exercise.workout)], ['Planks', 'Side Planks']);
  assert.equal(groups[1].category, 'Upper Body');
});

test('bolds the date cell when writing a supplemental workout group', () => {
  const fontWeights = [];
  const sheet = {
    getRange(row, column) {
      return {
        setValues: () => this,
        setNumberFormat: () => this,
        setFontWeight: value => { fontWeights.push({ row, column, value }); },
        merge: () => ({ setVerticalAlignment: () => {} })
      };
    }
  };
  const context = loadContext();
  vm.runInContext('globalThis.writeGroupTarget = writeSupplementalWorkoutGroup;', context);

  context.writeGroupTarget(sheet, 1, new Date(2026, 8, 16), {
    category: 'Core',
    exercises: [
      { workout: 'Planks', sets: '2', repsHoldTime: '1:30', weight: 'N/A', notes: '' }
    ]
  });

  assert.deepEqual(fontWeights, [{ row: 2, column: 1, value: 'bold' }]);
});

test('writes one merged row block per category', () => {
  const values = [];
  const merges = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: rowValues => { values.push({ row, column, rowValues: rowValues[0] }); return this; },
        setNumberFormat: () => this,
        setFontWeight: () => this,
        merge: () => { merges.push({ row, column, numRows, numColumns }); return { setVerticalAlignment: () => {} }; }
      };
    }
  };
  const context = loadContext();
  vm.runInContext('globalThis.writeGroupTarget = writeSupplementalWorkoutGroup;', context);

  const nextRowIndex = context.writeGroupTarget(sheet, 1, new Date(2026, 8, 16), {
    category: 'Core',
    exercises: [
      { workout: 'Planks', sets: '2', repsHoldTime: '1:30', weight: 'N/A', notes: '' },
      { workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', weight: 'N/A', notes: 'each side' }
    ]
  });

  assert.equal(nextRowIndex, 3);
  assert.deepEqual(JSON.parse(JSON.stringify(values[0].rowValues)), ['September 16', 'Core', 'Planks', '2', '1:30', 'N/A', '']);
  assert.deepEqual(JSON.parse(JSON.stringify(values[1].rowValues)), ['', '', 'Side Planks', '2', '1:00', 'N/A', 'each side']);
  assert.deepEqual(JSON.parse(JSON.stringify(merges)), [
    { row: 2, column: 1, numRows: 2, numColumns: 1 },
    { row: 2, column: 2, numRows: 2, numColumns: 1 }
  ]);
});

test('finds an existing day block and reconstructs its groups', () => {
  const context = loadContext();
  vm.runInContext('globalThis.findBlockTarget = findDayBlock; globalThis.blockGroupsTarget = getBlockGroups; globalThis.signatureTarget = buildDaySignature;', context);

  const sheetData = [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [new Date(2026, 8, 16), 'Core', 'Planks', '2', '1:30', 'N/A', ''],
    ['', '', 'Side Planks', '2', '1:00', 'N/A', 'each side'],
    ['', '', '', '', '', '', ''],
    [new Date(2026, 8, 17), 'Upper Body', 'Bench Press', '3', '10', '135 lbs', '']
  ];

  const block = context.findBlockTarget(sheetData, '2026-09-16');
  assert.deepEqual(JSON.parse(JSON.stringify(block)), { startRowIndex: 1, rowCount: 2 });

  const groups = context.blockGroupsTarget(sheetData, block);
  assert.equal(context.signatureTarget('2026-09-16', groups), context.signatureTarget('2026-09-16', [
    { category: 'Core', exercises: [
      { workout: 'Planks', sets: '2', repsHoldTime: '1:30', weight: 'N/A', notes: '' },
      { workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', weight: 'N/A', notes: 'each side' }
    ] }
  ]));

  assert.equal(context.findBlockTarget(sheetData, '2026-09-18'), null);
});

test('detects the Non Running Day note regardless of case, dashes, or underscores', () => {
  const context = loadContext();
  vm.runInContext('globalThis.noteTarget = isNonRunningDayNote;', context);

  assert.equal(context.noteTarget('Non Running Day'), true);
  assert.equal(context.noteTarget('non running day'), true);
  assert.equal(context.noteTarget('NON-RUNNING-DAY'), true);
  assert.equal(context.noteTarget('Non_Running_Day'), true);
  assert.equal(context.noteTarget('rest day, non-running day'), true);
  assert.equal(context.noteTarget('felt good'), false);
  assert.equal(context.noteTarget(''), false);
});

test('does not prune a Non Running Day block that falls before the cutoff', () => {
  const deletedRows = [];
  const sheet = { deleteRows: (row, count) => { deletedRows.push({ row, count }); } };
  const context = loadContext();
  vm.runInContext('globalThis.removeBeforeTarget = removeSupplementalWorkoutsBefore;', context);

  context.removeBeforeTarget(sheet, [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
    [new Date(2026, 8, 1), 'Core', 'Planks', '2', '1:00', 'N/A', 'Non Running Day'],
    [],
    [new Date(2026, 8, 2), 'Core', 'Planks', '2', '1:00', 'N/A', '']
  ], new Date(2026, 8, 15));

  assert.deepEqual(deletedRows, [{ row: 4, count: 2 }]);
});

test('preserves a manually entered Non Running Day block instead of overwriting it', () => {
  const tempDir = mkdtempSync(path.join(tmpdir(), 'supplemental-workouts-rest-'));
  const fixturePath = path.join(tempDir, 'spreadsheet.json');
  const outputPath = path.join(tempDir, 'output.json');

  writeFileSync(fixturePath, JSON.stringify({
    sheets: {
      'Supplemental Workouts': {
        values: [
          ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
          [],
          ['2026-09-16T04:00:00.000Z', 'Core', 'Manual Planks', '5', '2:00', 'N/A', 'Non Running Day'],
          []
        ]
      }
    }
  }));

  try {
    const mock = createGoogleSheetsMock(fixturePath, outputPath);
    const context = vm.createContext({ console, SpreadsheetApp: mock.SpreadsheetApp });
    const source = [
      readFileSync('DateUtils.js', 'utf8'),
      readFileSync('RunningMetrics.js', 'utf8'),
      readFileSync('InjuryReport.js', 'utf8'),
      readFileSync('SheetUtils.js', 'utf8'),
      readFileSync('SupplementalWorkouts.js', 'utf8')
    ].join('\n');
    vm.runInContext(source, context);
    vm.runInContext(
      'updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), globalThis.__activities)',
      Object.assign(context, { __activities: [{
        start_date_local: '2026-09-16T13:00:00Z',
        description: 'Supplemental Workouts:\nCore\n1. Planks (1x1:00)'
      }] })
    );
    mock.save();

    const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
    const values = saved.sheets['Supplemental Workouts'].values;
    const manualRow = values.find(row => row?.[2] === 'Manual Planks');

    assert.ok(manualRow, 'the manually entered row should still be present');
    assert.equal(manualRow[3], '5');
    assert.equal(manualRow[6], 'Non Running Day');
    assert.equal(values.some(row => row?.[2] === 'Planks'), false);
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

test('overwrites a day already recorded when its exercises change', () => {
  const tempDir = mkdtempSync(path.join(tmpdir(), 'supplemental-workouts-'));
  const fixturePath = path.join(tempDir, 'spreadsheet.json');
  const outputPath = path.join(tempDir, 'output.json');
  writeFileSync(fixturePath, JSON.stringify({ sheets: {} }));

  try {
    const runUpdate = (description) => {
      const mock = createGoogleSheetsMock(fixturePath, outputPath);
      const context = vm.createContext({ console, SpreadsheetApp: mock.SpreadsheetApp });
      const source = [
        readFileSync('DateUtils.js', 'utf8'),
        readFileSync('RunningMetrics.js', 'utf8'),
        readFileSync('InjuryReport.js', 'utf8'),
        readFileSync('SheetUtils.js', 'utf8'),
        readFileSync('SupplementalWorkouts.js', 'utf8')
      ].join('\n');
      vm.runInContext(source, context);
      vm.runInContext(
        'updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), globalThis.__activities)',
        Object.assign(context, { __activities: [{
          start_date_local: '2026-09-16T13:00:00Z',
          type: 'Run',
          description
        }] })
      );
      mock.save();
      const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
      writeFileSync(fixturePath, JSON.stringify({ sheets: saved.sheets }));
      return saved.sheets['Supplemental Workouts'].values;
    };

    const firstValues = runUpdate('Supplemental Workouts:\nCore\n1. Planks (1x1:00)');
    const firstWorkoutRow = firstValues.find(row => row?.[2] === 'Planks');
    assert.equal(firstWorkoutRow[3], '1');

    const secondValues = runUpdate('Supplemental Workouts:\nCore\n1. Planks (2x1:30)\n2. Side Planks (2x1:00, each side)');
    const secondWorkoutRows = secondValues.filter(row => ['Planks', 'Side Planks'].includes(row?.[2]));
    assert.equal(secondWorkoutRows[0][3], '2');
    assert.equal(secondWorkoutRows[1][2], 'Side Planks');
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

test('keeps day blocks in chronological order even when activities arrive out of order', () => {
  const tempDir = mkdtempSync(path.join(tmpdir(), 'supplemental-workouts-order-'));
  const fixturePath = path.join(tempDir, 'spreadsheet.json');
  const outputPath = path.join(tempDir, 'output.json');
  writeFileSync(fixturePath, JSON.stringify({ sheets: {} }));

  try {
    const runUpdate = (activities) => {
      const mock = createGoogleSheetsMock(fixturePath, outputPath);
      const context = vm.createContext({ console, SpreadsheetApp: mock.SpreadsheetApp });
      const source = [
        readFileSync('DateUtils.js', 'utf8'),
        readFileSync('RunningMetrics.js', 'utf8'),
        readFileSync('InjuryReport.js', 'utf8'),
        readFileSync('SheetUtils.js', 'utf8'),
        readFileSync('SupplementalWorkouts.js', 'utf8')
      ].join('\n');
      vm.runInContext(source, context);
      vm.runInContext(
        'updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), globalThis.__activities)',
        Object.assign(context, { __activities: activities })
      );
      mock.save();
      const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
      writeFileSync(fixturePath, JSON.stringify({ sheets: saved.sheets }));
      return saved.sheets['Supplemental Workouts'].values;
    };

    // 9/17 is recorded first, then a later sync processes 9/15 (e.g. an edited older activity).
    runUpdate([{
      start_date_local: '2026-09-17T13:00:00Z',
      description: 'Supplemental Workouts:\nUpper Body\n1. Bench Press (3x8)'
    }]);
    const values = runUpdate([{
      start_date_local: '2026-09-15T13:00:00Z',
      description: 'Supplemental Workouts:\nCore\n1. Planks (1x1:00)'
    }]);

    const dateCells = values.slice(1).map(row => row?.[0]).filter(Boolean);
    assert.deepEqual(dateCells, ['September 15', 'September 17']);
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

test('re-sorts an already out-of-order sheet even with no new activities', () => {
  const tempDir = mkdtempSync(path.join(tmpdir(), 'supplemental-workouts-resort-'));
  const fixturePath = path.join(tempDir, 'spreadsheet.json');
  const outputPath = path.join(tempDir, 'output.json');

  writeFileSync(fixturePath, JSON.stringify({
    sheets: {
      'Supplemental Workouts': {
        values: [
          ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
          ['2026-09-17T04:00:00.000Z', 'Upper Body', 'Bench Press', '3', '8', 'N/A', ''],
          [],
          ['2026-09-15T04:00:00.000Z', 'Core', 'Planks', '1', '1:00', 'N/A', ''],
          []
        ]
      }
    }
  }));

  try {
    const mock = createGoogleSheetsMock(fixturePath, outputPath);
    const context = vm.createContext({ console, SpreadsheetApp: mock.SpreadsheetApp });
    const source = [
      readFileSync('DateUtils.js', 'utf8'),
      readFileSync('RunningMetrics.js', 'utf8'),
      readFileSync('InjuryReport.js', 'utf8'),
      readFileSync('SheetUtils.js', 'utf8'),
      readFileSync('SupplementalWorkouts.js', 'utf8')
    ].join('\n');
    vm.runInContext(source, context);
    vm.runInContext('updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), [])', context);
    mock.save();

    const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
    const values = saved.sheets['Supplemental Workouts'].values;
    // Blocks already in position keep their original cells, so assert chronology rather than the
    // display text that `normalizeSheetDateColumn` produces in Sheets.
    const dateCells = values.slice(1).map(row => row?.[0]).filter(Boolean);
    assert.deepEqual(dateCells.map(value => new Date(value).getDate()), [15, 17]);
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

function loadSupplementalActualRunsContext(mock) {
  const context = vm.createContext({ console, SpreadsheetApp: mock.SpreadsheetApp });
  const source = [
    readFileSync('DateUtils.js', 'utf8'),
    readFileSync('RunningMetrics.js', 'utf8'),
    readFileSync('InjuryReport.js', 'utf8'),
    readFileSync('SheetUtils.js', 'utf8'),
    readFileSync('SupplementalWorkouts.js', 'utf8')
  ].join('\n');
  vm.runInContext(source, context);
  return context;
}

function runSupplementalActualRunsSync(fixtureSheets, todayOffset = 6) {
  const tempDir = mkdtempSync(path.join(tmpdir(), 'supplemental-actual-runs-'));
  const fixturePath = path.join(tempDir, 'spreadsheet.json');
  const outputPath = path.join(tempDir, 'output.json');
  writeFileSync(fixturePath, JSON.stringify({ sheets: fixtureSheets }));

  try {
    const mock = createGoogleSheetsMock(fixturePath, outputPath);
    const context = loadSupplementalActualRunsContext(mock);
    context.todayOffset = todayOffset;
    vm.runInContext(`
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      globalThis.updatedDayIndexes = applySupplementalWorkoutsToActualRuns(
        spreadsheet,
        spreadsheet.getSheetByName('Actual Runs'),
        1,
        new Date(2026, 9, 12),
        todayOffset
      );
    `, context);
    mock.save();

    const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
    return { values: saved.sheets['Actual Runs'].values, updatedDayIndexes: context.updatedDayIndexes };
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
}

const SUPPLEMENTAL_ACTUAL_RUNS_HEADER = ['', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday', 'Total Mileage'];

test('appends a rest day supplemental workout to its Actual Runs cell', () => {
  const { values, updatedDayIndexes } = runSupplementalActualRunsSync({
    'Actual Runs': {
      values: [
        SUPPLEMENTAL_ACTUAL_RUNS_HEADER,
        ['WEEK 24 (10/12)', '4.15 miles', '5.78 miles', 'Rest', '', '', '', '', '9.93 miles']
      ]
    },
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-14T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8', '90 lbs', 'Non Running Day'],
        []
      ]
    }
  }, 2);

  assert.deepEqual(Array.from(updatedDayIndexes), [2]);
  assert.equal(values[1][3], 'Rest\nUpper Body Lift');
});

test('appends every logged category for a day in sheet order', () => {
  const { values } = runSupplementalActualRunsSync({
    'Actual Runs': {
      values: [
        SUPPLEMENTAL_ACTUAL_RUNS_HEADER,
        ['WEEK 24 (10/12)', '', 'Rest', '', '', '', '', '', '']
      ]
    },
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-13T04:00:00.000Z', 'Shoulder Rehab', 'No Money', '2', '15', 'N/A', ''],
        ['', '', 'Cable External Rotation', '1', '15', 'N/A', 'Each side'],
        ['2026-10-13T04:00:00.000Z', 'Accessory Exercises', 'Leg Lifts on Back', '2', '10', 'N/A', ''],
        ['2026-10-13T04:00:00.000Z', 'Lower Body Lift', 'Copenhagen Planks', '3', '30 seconds', 'N/A', ''],
        []
      ]
    }
  });

  assert.equal(values[1][2], 'Rest\nShoulder Rehab\nAccessory Exercises\nLower Body Lift');
});

test('leaves mileage and existing workout lines intact while adding only missing categories', () => {
  const { values, updatedDayIndexes } = runSupplementalActualRunsSync({
    'Actual Runs': {
      values: [
        SUPPLEMENTAL_ACTUAL_RUNS_HEADER,
        ['WEEK 24 (10/12)', '', '5.78 miles\nShoulder Rehab', '', '', '', '', '', '']
      ]
    },
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-13T04:00:00.000Z', 'Shoulder Rehab', 'No Money', '2', '15', 'N/A', ''],
        ['2026-10-13T04:00:00.000Z', 'Core', "Captain's Chair Leg Raise", '2', '10', 'N/A', ''],
        []
      ]
    }
  });

  assert.deepEqual(Array.from(updatedDayIndexes), [1]);
  assert.equal(values[1][2], '5.78 miles\nShoulder Rehab\nCore');
});

test('does not rewrite a cell that already lists every logged category', () => {
  const { values, updatedDayIndexes } = runSupplementalActualRunsSync({
    'Actual Runs': {
      values: [
        SUPPLEMENTAL_ACTUAL_RUNS_HEADER,
        ['WEEK 24 (10/12)', '', '5.78 miles\nShoulder Rehab', '', '', '', '', '', '']
      ]
    },
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-13T04:00:00.000Z', 'Shoulder Rehab', 'No Money', '2', '15', 'N/A', ''],
        []
      ]
    }
  });

  assert.deepEqual(Array.from(updatedDayIndexes), []);
  assert.equal(values[1][2], '5.78 miles\nShoulder Rehab');
});

test('ignores supplemental workouts logged after the current day of the week', () => {
  const { updatedDayIndexes } = runSupplementalActualRunsSync({
    'Actual Runs': {
      values: [
        SUPPLEMENTAL_ACTUAL_RUNS_HEADER,
        ['WEEK 24 (10/12)', '', '', '', '', '', '', '', '']
      ]
    },
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-16T04:00:00.000Z', 'Core', 'Ab Wheel', '2', '6', 'N/A', ''],
        []
      ]
    }
  }, 2);

  assert.deepEqual(Array.from(updatedDayIndexes), []);
});

function runSupplementalSheetSync(fixtureSheets, activities, options = {}, { production = false } = {}) {
  const tempDir = mkdtempSync(path.join(tmpdir(), 'supplemental-manual-'));
  const fixturePath = path.join(tempDir, 'spreadsheet.json');
  const outputPath = path.join(tempDir, 'output.json');
  writeFileSync(fixturePath, JSON.stringify({ sheets: fixtureSheets }));

  try {
    const mock = createGoogleSheetsMock(fixturePath, outputPath);
    const context = vm.createContext({
      console,
      SpreadsheetApp: mock.SpreadsheetApp,
      __activities: activities,
      __options: options,
      // Presence of PropertiesService is how the script detects it is running in production,
      // which selects a different pruning/separator branch than the local harness.
      ...(production
        ? { PropertiesService: { getScriptProperties: () => ({ getProperty: () => '', setProperty: () => {} }) } }
        : {})
    });
    const source = [
      readFileSync('DateUtils.js', 'utf8'),
      readFileSync('RunningMetrics.js', 'utf8'),
      readFileSync('InjuryReport.js', 'utf8'),
      readFileSync('SheetUtils.js', 'utf8'),
      readFileSync('SupplementalWorkouts.js', 'utf8')
    ].join('\n');
    vm.runInContext(source, context);
    vm.runInContext(
      'updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), globalThis.__activities, globalThis.__options)',
      context
    );
    mock.save();

    const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
    return saved.sheets['Supplemental Workouts'].values;
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
}

// Values pulled from the live sheet arrive with trailing empty cells and the trailing separator
// row trimmed off, and with no merge metadata, so the manual block must survive that shape too.
const PULLED_MANUAL_ENTRY_SHEET = () => ({
  'Supplemental Workouts': {
    values: [
      ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
      [],
      ['2026-10-13T04:00:00.000Z', 'Lower Body Lift', 'Copenhagen Planks', '3', '30 seconds', 'N/A', 'Each side'],
      [],
      ['2026-10-14T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8\n8\n6', '90 lbs\n95 lbs\n100 lbs'],
      ['', '', 'Dumbell Rows', '3', '8\n8\n8', '25 lbs\n30 lbs\n35 lbs']
    ],
    mergedRanges: []
  }
});

[false, true].forEach(production => {
  const environment = production ? 'production' : 'local';

  test(`preserves a pulled rest day manual entry in ${environment}`, () => {
    const values = runSupplementalSheetSync(
      PULLED_MANUAL_ENTRY_SHEET(),
      [],
      { manualDateKeys: { '2026-10-14': true }, earliestEditableDateKey: '2026-10-12' },
      { production }
    );

    const blockStart = values.findIndex(row => String(row?.[2]) === 'Bench Press');
    assert.ok(blockStart !== -1, 'the manual October 14 block should still exist');
    assert.equal(values[blockStart][1], 'Upper Body Lift');
    assert.equal(values[blockStart][5], '90 lbs\n95 lbs\n100 lbs');
    assert.equal(values[blockStart + 1][2], 'Dumbell Rows');
    assert.equal(values[blockStart + 1][5], '25 lbs\n30 lbs\n35 lbs');
  });

  test(`does not duplicate or drop a pulled manual entry on a repeat ${environment} sync`, () => {
    const first = runSupplementalSheetSync(
      PULLED_MANUAL_ENTRY_SHEET(),
      [],
      { manualDateKeys: { '2026-10-14': true }, earliestEditableDateKey: '2026-10-12' },
      { production }
    );
    const second = runSupplementalSheetSync(
      { 'Supplemental Workouts': { values: first, mergedRanges: [] } },
      [],
      { manualDateKeys: { '2026-10-14': true }, earliestEditableDateKey: '2026-10-12' },
      { production }
    );

    const benchRows = second.filter(row => String(row?.[2]) === 'Bench Press');
    assert.equal(benchRows.length, 1);
    assert.equal(benchRows[0][5], '90 lbs\n95 lbs\n100 lbs');
  });
});

test('leaves a rest day manual entry untouched even when an activity describes it', () => {  const values = runSupplementalSheetSync({
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-14T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8', 'HAND ENTERED', ''],
        ['', '', 'Dumbell Rows', '3', '10', 'HAND ENTERED', ''],
        []
      ]
    }
  }, [{
    start_date_local: '2026-10-14T19:10:00Z',
    type: 'WeightTraining',
    description: 'Supplemental Workouts:\nUpper Body Lift\n1. Bench Press (3x8, 200 lbs)'
  }], { manualDateKeys: { '2026-10-14': true }, earliestEditableDateKey: '2026-10-12' });

  const weights = values.slice(1).filter(row => row?.[2]).map(row => row[5]);
  assert.deepEqual(weights, ['HAND ENTERED', 'HAND ENTERED']);
});

test('still overwrites a day whose Actual Runs cell shows a run', () => {
  const values = runSupplementalSheetSync({
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-14T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8', 'STALE', ''],
        []
      ]
    }
  }, [{
    start_date_local: '2026-10-14T12:00:00Z',
    type: 'Run',
    distance: 8046,
    description: 'Supplemental Workouts:\nUpper Body Lift\n1. Bench Press (3x8, 200 lbs)'
  }], { manualDateKeys: {}, earliestEditableDateKey: '2026-10-12' });

  const benchRow = values.find(row => row?.[2] === 'Bench Press');
  assert.equal(benchRow[5], '200 lbs');
});

test('never rewrites log entries that fall before the sync window', () => {
  const values = runSupplementalSheetSync({
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-09-17T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8', 'HISTORICAL', ''],
        []
      ]
    }
  }, [{
    start_date_local: '2026-09-17T12:00:00Z',
    type: 'Run',
    distance: 8046,
    description: 'Supplemental Workouts:\nUpper Body Lift\n1. Bench Press (3x8, 999 lbs)'
  }], { manualDateKeys: {}, earliestEditableDateKey: '2026-10-12' });

  const benchRow = values.find(row => row?.[2] === 'Bench Press');
  assert.equal(benchRow[5], 'HISTORICAL');
});

test('leaves an out-of-order block before the sync window where it is', () => {
  const values = runSupplementalSheetSync({
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-09-20T04:00:00.000Z', 'Core', 'Planks', '1', '1:00', 'N/A', ''],
        [],
        ['2026-09-17T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8', 'N/A', ''],
        []
      ]
    }
  }, [], { manualDateKeys: {}, earliestEditableDateKey: '2026-10-12' });

  const orderedDays = values.slice(1).map(row => row?.[0]).filter(Boolean)
    .map(value => new Date(value).getDate());
  assert.deepEqual(orderedDays, [20, 17]);
});

test('sorts hand-entered rest day blocks into chronological order', () => {
  const values = runSupplementalSheetSync({
    'Supplemental Workouts': {
      values: [
        ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'],
        [],
        ['2026-10-14T04:00:00.000Z', 'Upper Body Lift', 'Bench Press', '3', '8', '90 lbs', ''],
        [],
        ['2026-10-12T04:00:00.000Z', 'Core', 'Ab Wheel', '2', '6', 'N/A', ''],
        [],
        ['2026-10-13T04:00:00.000Z', 'Shoulder Rehab', 'No Money', '2', '15', 'N/A', ''],
        []
      ]
    }
  }, [], { manualDateKeys: { '2026-10-14': true }, earliestEditableDateKey: '2026-10-12' });

  // The mock returns Date objects from another realm, so compare chronology rather than the
  // display text that `normalizeSheetDateColumn` would produce in Sheets.
  const dateCells = values.slice(1).map(row => row?.[0]).filter(Boolean);
  const orderedDays = dateCells.map(value => new Date(value).getDate());
  assert.deepEqual(orderedDays, [12, 13, 14]);
});

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
  const source = readFileSync('SupplementalWorkouts.js', 'utf8');
  vm.runInContext(source, context);
  return context;
}

test('fills E:G for rows containing a black separator cell in A:D', () => {
  const backgrounds = { '3,2': '#000000' };
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

  assert.deepEqual(fills, [{ row: 3, column: 5, numRows: 1, numColumns: 3, value: '#000000' }]);
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

test('writes one merged row block per category', () => {
  const values = [];
  const merges = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: rowValues => { values.push({ row, column, rowValues: rowValues[0] }); return this; },
        setNumberFormat: () => this,
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
  assert.deepEqual(JSON.parse(JSON.stringify(values[0].rowValues)), JSON.parse(JSON.stringify([new Date(2026, 8, 16), 'Core', 'Planks', '2', '1:30', 'N/A', ''])));
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
        readFileSync('SupplementalWorkouts.js', 'utf8')
      ].join('\n');
      vm.runInContext(source, context);
      vm.runInContext(
        'updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), globalThis.__activities)',
        Object.assign(context, { __activities: [{
          start_date_local: '2026-09-16T13:00:00Z',
          description
        }] })
      );
      mock.save();
      const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
      writeFileSync(fixturePath, JSON.stringify({ sheets: saved.sheets }));
      return saved.sheets['Supplemental Workouts'].values;
    };

    const firstValues = runUpdate('Supplemental Workouts:\nCore\n1. Planks (1x1:00)');
    assert.equal(firstValues.length, 2);
    assert.equal(firstValues[1][3], '1');

    const secondValues = runUpdate('Supplemental Workouts:\nCore\n1. Planks (2x1:30)\n2. Side Planks (2x1:00, each side)');
    assert.equal(secondValues.length, 3);
    assert.equal(secondValues[1][3], '2');
    assert.equal(secondValues[2][2], 'Side Planks');
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
    assert.deepEqual(dateCells, ['2026-09-15T04:00:00.000Z', '2026-09-17T04:00:00.000Z']);
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
      readFileSync('SupplementalWorkouts.js', 'utf8')
    ].join('\n');
    vm.runInContext(source, context);
    vm.runInContext('updateSupplementalWorkoutsSheet(SpreadsheetApp.getActiveSpreadsheet(), [])', context);
    mock.save();

    const saved = JSON.parse(readFileSync(outputPath, 'utf8'));
    const values = saved.sheets['Supplemental Workouts'].values;
    const dateCells = values.slice(1).map(row => row?.[0]).filter(Boolean);
    assert.deepEqual(dateCells, ['2026-09-15T04:00:00.000Z', '2026-09-17T04:00:00.000Z']);
  } finally {
    rmSync(tempDir, { recursive: true, force: true });
  }
});

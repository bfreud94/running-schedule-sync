const test = require('node:test');
const assert = require('node:assert');
const { readFileSync } = require('node:fs');
const vm = require('node:vm');
const { createSpreadsheetApp } = require('./local/GoogleSheetsMock');

// The shared mock is permissive about merged cells, but Google Sheets is not. These helpers layer
// the real API's merge rules on top of it so the deployed behaviour is exercised here:
//   - reading a merged region yields the value only in its top-left cell
//   - writing a range that splits a merged region throws "You can't change part of a merged cell."
function coveringMerge(mergedRanges, row, column) {
  return mergedRanges.find(merge =>
    row >= merge.row && row < merge.row + merge.numRows &&
    column >= merge.column && column < merge.column + merge.numColumns
  );
}

function assertWriteDoesNotSplitMerge(mergedRanges, row, column, numRows, numColumns) {
  mergedRanges.forEach(merge => {
    const intersects = row < merge.row + merge.numRows && merge.row < row + numRows
      && column < merge.column + merge.numColumns && merge.column < column + numColumns;
    if (!intersects) return;

    const fullyContained = row <= merge.row && merge.row + merge.numRows <= row + numRows
      && column <= merge.column && merge.column + merge.numColumns <= column + numColumns;
    if (!fullyContained) {
      throw new Error("You can't change part of a merged cell.");
    }
  });
}

function createStrictSheetsMock(workbook) {
  // createSpreadsheetApp mutates the workbook object passed in, so the merge state read below
  // always matches what the sheet code has actually done.
  const { SpreadsheetApp } = createSpreadsheetApp(workbook);
  const inner = SpreadsheetApp.getActiveSpreadsheet();

  const wrapSheet = (sheet, sheetName) => {
    if (!sheet) return sheet;
    const mergesOf = () => workbook.sheets[sheetName].mergedRanges || [];

    return new Proxy(sheet, {
      get(target, property) {
        if (property === 'getDataRange') {
          return () => ({
            getValues: () => target.getDataRange().getValues().map((row, rowIndex) =>
              row.map((value, columnIndex) => {
                const merge = coveringMerge(mergesOf(), rowIndex + 1, columnIndex + 1);
                const isAnchor = merge && merge.row === rowIndex + 1 && merge.column === columnIndex + 1;
                return merge && !isAnchor ? '' : value;
              })
            )
          });
        }

        if (property === 'getRange') {
          return (row, column, numRows = 1, numColumns = 1) => {
            const range = target.getRange(row, column, numRows, numColumns);
            return new Proxy(range, {
              get(rangeTarget, rangeProperty) {
                if (rangeProperty === 'setValues' || rangeProperty === 'setValue') {
                  return (...args) => {
                    assertWriteDoesNotSplitMerge(mergesOf(), row, column, numRows, numColumns);
                    rangeTarget[rangeProperty](...args);
                    return range;
                  };
                }
                if (rangeProperty === 'getValue') {
                  const merge = coveringMerge(mergesOf(), row, column);
                  const isAnchor = merge && merge.row === row && merge.column === column;
                  if (merge && !isAnchor) return () => '';
                }
                const value = rangeTarget[rangeProperty];
                return typeof value === 'function' ? value.bind(rangeTarget) : value;
              }
            });
          };
        }

        const value = target[property];
        return typeof value === 'function' ? value.bind(target) : value;
      }
    });
  };

  const spreadsheet = {
    getSheetByName: sheetName => wrapSheet(inner.getSheetByName(sheetName), sheetName),
    insertSheet: sheetName => wrapSheet(inner.insertSheet(sheetName), sheetName),
    getSheets: () => inner.getSheets()
  };

  return {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, flush: () => {} }
  };
}

// Mirrors how production stores a day: the Date and Workout cells are merged across the exercises.
function mergedDayBlock(startRow, date, category, exercises) {
  const values = exercises.map((exercise, index) => ([
    index === 0 ? date : '',
    index === 0 ? category : '',
    ...exercise
  ]));
  const mergedRanges = exercises.length > 1
    ? [
      { row: startRow, column: 1, numRows: exercises.length, numColumns: 1 },
      { row: startRow, column: 2, numRows: exercises.length, numColumns: 1 }
    ]
    : [];
  return { values, mergedRanges };
}

function runProductionSync(sheetValues, mergedRanges, activities, options) {
  const workbook = {
    sheets: {
      'Supplemental Workouts': {
        values: sheetValues.map(row => [...(row || [])]),
        styles: {},
        columnWidths: {},
        rowHeights: {},
        mergedRanges: mergedRanges.map(merge => ({ ...merge }))
      }
    }
  };
  const mock = createStrictSheetsMock(workbook);
  const context = vm.createContext({
    console,
    SpreadsheetApp: mock.SpreadsheetApp,
    // Defining PropertiesService is what makes the script take its production branch.
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => '', setProperty: () => {} }) },
    __activities: activities,
    __options: options
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

  return workbook.sheets['Supplemental Workouts'];
}

const OCT_14_EXERCISES = [
  ['Bench Press', '3', '8\n8\n6', '90 lbs\n95 lbs\n100 lbs', ''],
  ['Dumbell Rows', '3', '8\n8\n8', '25 lbs\n30 lbs\n35 lbs', ''],
  ['Overhead Press', '3', '8\n8\n8', '50 lbs\n50 lbs\n55 lbs', ''],
  ['Lat Rows', '3', '8\n8\n8', '90 lbs\n90 lbs\n80 lbs', '']
];

const MANUAL_OPTIONS = { manualDateKeys: { '2026-10-14': true }, earliestEditableDateKey: '2026-10-12' };

function buildSheet({ manualDayFirst = false } = {}) {
  const header = [['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'], []];
  const runDay = mergedDayBlock(0, '2026-10-13T04:00:00.000Z', 'Lower Body Lift', [
    ['Copenhagen Planks', '3', '30 seconds', 'N/A', 'Each side']
  ]);
  const manualDay = mergedDayBlock(0, '2026-10-14T04:00:00.000Z', 'Upper Body Lift', OCT_14_EXERCISES);

  const first = manualDayFirst ? manualDay : runDay;
  const second = manualDayFirst ? runDay : manualDay;

  const values = [...header, ...first.values, [], ...second.values, []];
  const firstStart = header.length + 1;
  const secondStart = firstStart + first.values.length + 1;
  const mergedRanges = [
    ...first.mergedRanges.map(merge => ({ ...merge, row: firstStart })),
    ...second.mergedRanges.map(merge => ({ ...merge, row: secondStart }))
  ];
  return { values, mergedRanges };
}

// Guards against this suite silently passing: if the merge rules above stopped being enforced,
// these checks would fail and every production test below would become meaningless.
test('production harness: enforces Google Sheets merge rules', () => {
  const merges = [{ row: 3, column: 1, numRows: 4, numColumns: 1 }];

  assert.throws(
    () => assertWriteDoesNotSplitMerge(merges, 3, 1, 1, 7),
    /part of a merged cell/,
    'writing a single row across a 4-row merge must be rejected'
  );
  assert.doesNotThrow(
    () => assertWriteDoesNotSplitMerge(merges, 3, 1, 4, 7),
    'writing a range that fully contains the merge is allowed'
  );
  assert.doesNotThrow(
    () => assertWriteDoesNotSplitMerge(merges, 8, 1, 1, 7),
    'writing outside the merge is allowed'
  );

  const workbook = {
    sheets: {
      'Supplemental Workouts': {
        values: [['Date'], [], ['October 14', 'Upper Body Lift', 'Bench Press'], ['', '', 'Dumbell Rows']],
        styles: {},
        columnWidths: {},
        rowHeights: {},
        mergedRanges: [{ row: 3, column: 1, numRows: 2, numColumns: 1 }]
      }
    }
  };
  const sheet = createStrictSheetsMock(workbook).SpreadsheetApp
    .getActiveSpreadsheet().getSheetByName('Supplemental Workouts');

  assert.throws(() => sheet.getRange(3, 1, 1, 7).setValues([['x', '', '', '', '', '', '']]), /merged cell/);
  // A merged region reads back only through its top-left cell.
  assert.equal(sheet.getDataRange().getValues()[3][0], '');
});

test('production: preserves a merged manual rest day block without splitting merged cells', () => {
  const { values, mergedRanges } = buildSheet();
  const result = runProductionSync(values, mergedRanges, [], MANUAL_OPTIONS);

  const benchRow = result.values.findIndex(row => String(row?.[2]) === 'Bench Press');
  assert.ok(benchRow !== -1, 'the manual October 14 block should survive');
  // The vm realm makes `value instanceof Date` fail inside parseSheetDate, so the cell can keep an
  // ISO value locally where Sheets would render "October 14"; assert the date itself instead.
  assert.equal(new Date(result.values[benchRow][0]).getDate(), 14);
  assert.equal(result.values[benchRow][1], 'Upper Body Lift');
  assert.deepEqual(
    Array.from(result.values.slice(benchRow, benchRow + 4).map(row => row[2])),
    ['Bench Press', 'Dumbell Rows', 'Overhead Press', 'Lat Rows']
  );
  assert.equal(result.values[benchRow + 1][5], '25 lbs\n30 lbs\n35 lbs');
});

test('production: an activity describing the rest day cannot overwrite the manual block', () => {
  const { values, mergedRanges } = buildSheet();
  const result = runProductionSync(values, mergedRanges, [{
    start_date_local: '2026-10-14T19:10:00Z',
    type: 'WeightTraining',
    description: 'Supplemental Workouts:\nUpper Body Lift\n1. Bench Press (3x8, 500 lbs)'
  }], MANUAL_OPTIONS);

  const benchRow = result.values.find(row => String(row?.[2]) === 'Bench Press');
  assert.equal(benchRow[5], '90 lbs\n95 lbs\n100 lbs');
  assert.ok(!result.values.some(row => String(row?.[5]).includes('500 lbs')));
});

test('production: re-sorting a misplaced manual block keeps its rows and merges intact', () => {
  const { values, mergedRanges } = buildSheet({ manualDayFirst: true });
  const result = runProductionSync(values, mergedRanges, [], MANUAL_OPTIONS);

  const dates = result.values.map(row => row?.[0]).filter(value => value && value !== 'Date');
  assert.deepEqual(Array.from(dates.map(value => new Date(value).getDate())), [13, 14]);

  const benchRow = result.values.findIndex(row => String(row?.[2]) === 'Bench Press');
  assert.equal(result.values[benchRow][1], 'Upper Body Lift');
  assert.deepEqual(
    Array.from(result.values.slice(benchRow, benchRow + 4).map(row => row[2])),
    ['Bench Press', 'Dumbell Rows', 'Overhead Press', 'Lat Rows']
  );

  // The relocated block must regain merged Date/Workout cells spanning all four exercises.
  const dateMerge = result.mergedRanges.find(merge => merge.row === benchRow + 1 && merge.column === 1);
  assert.ok(dateMerge, 'the moved block should still have a merged Date cell');
  assert.equal(dateMerge.numRows, 4);
});

test('production: repeat syncs neither duplicate nor shrink the manual block', () => {
  const { values, mergedRanges } = buildSheet();
  const first = runProductionSync(values, mergedRanges, [], MANUAL_OPTIONS);
  const second = runProductionSync(first.values, first.mergedRanges, [], MANUAL_OPTIONS);

  assert.equal(second.values.filter(row => String(row?.[2]) === 'Bench Press').length, 1);
  const benchRow = second.values.findIndex(row => String(row?.[2]) === 'Bench Press');
  assert.deepEqual(
    Array.from(second.values.slice(benchRow, benchRow + 4).map(row => row[2])),
    ['Bench Press', 'Dumbell Rows', 'Overhead Press', 'Lat Rows']
  );
});

const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('writes one supplemental row per category without duplicating existing data', () => {
  const writes = [];
  const numberFormats = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: values => { writes.push({ row, column, numRows, numColumns, values }); return this; },
        setNumberFormat: value => { numberFormats.push({ row, column, value }); return this; }
      };
    }
  };
  const context = vm.createContext({
    console,
    getDateKey: date => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`,
    parseSheetDate: value => value instanceof Date ? value : null
  });
  const source = readFileSync('SupplementalWorkouts.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.appendTarget = appendSupplementalWorkoutRows;`, context);
  context.appendTarget(
    sheet,
    [
      ['Date', 'Workout', 'Sets', 'Reps/Hold Time', 'Notes'],
      [new Date(2026, 8, 10), 'Core', 'Planks: 2', 'Planks: 60 seconds', '']
    ],
    [
      [new Date(2026, 8, 10), 'Core', 'Planks: 2', 'Planks: 60 seconds', ''],
      [new Date(2026, 8, 10), 'Upper Body', 'Bench Press: 3', 'Bench Press: 10', 'Bench Press: 70 lbs']
    ]
  );

  assert.equal(writes.length, 1);
  assert.equal(writes[0].row, 3);
  assert.deepEqual(writes[0].values[0].slice(1), ['Upper Body', 'Bench Press: 3', 'Bench Press: 10', 'Bench Press: 70 lbs']);
  assert.deepEqual(numberFormats, [{ row: 3, column: 1, value: 'mmmm d, yyyy' }]);
});

test('groups supplemental exercises by category', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('SupplementalWorkouts.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.groupTarget = groupSupplementalWorkoutRows;`, context);
  const rows = JSON.parse(JSON.stringify(context.groupTarget(new Date(2026, 8, 16), [
    { category: 'Core', workout: 'Planks', sets: '2', repsHoldTime: '60 seconds', notes: '' },
    { category: 'Core', workout: 'Side planks', sets: '2', repsHoldTime: '60 seconds', notes: 'each side' },
    { category: 'Upper Body', workout: 'Bench Press', sets: '3', repsHoldTime: '10', notes: '70 lbs, 75 lbs, 80 lbs' }
  ])));

  assert.deepEqual(rows.map(row => row.slice(1)), [
    ['Core', 'Planks: 2\nSide planks: 2', 'Planks: 60 seconds\nSide planks: 60 seconds', 'Side planks: each side'],
    ['Upper Body', 'Bench Press: 3', 'Bench Press: 10', 'Bench Press: 70 lbs, 75 lbs, 80 lbs']
  ]);
});
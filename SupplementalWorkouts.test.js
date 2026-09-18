const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

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

test('groups exercises by category, preserving first-seen order', () => {
  const context = loadContext();
  vm.runInContext('globalThis.groupTarget = groupExercisesByCategory;', context);

  const groups = context.groupTarget([
    { category: 'Core', workout: 'Planks', sets: '2', repsHoldTime: '1:30', notes: '' },
    { category: 'Upper Body', workout: 'Bench Press', sets: '3', repsHoldTime: '10', notes: '' },
    { category: 'Core', workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', notes: 'each side' }
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
      { workout: 'Planks', sets: '2', repsHoldTime: '1:30', notes: '' },
      { workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', notes: 'each side' }
    ]
  });

  assert.equal(nextRowIndex, 3);
  assert.deepEqual(JSON.parse(JSON.stringify(values[0].rowValues)), JSON.parse(JSON.stringify([new Date(2026, 8, 16), 'Core', 'Planks', '2', '1:30', ''])));
  assert.deepEqual(JSON.parse(JSON.stringify(values[1].rowValues)), ['', '', 'Side Planks', '2', '1:00', 'each side']);
  assert.deepEqual(JSON.parse(JSON.stringify(merges)), [
    { row: 2, column: 1, numRows: 2, numColumns: 1 },
    { row: 2, column: 2, numRows: 2, numColumns: 1 }
  ]);
});

test('does not duplicate a day already recorded with the same exercises', () => {
  const context = loadContext();
  vm.runInContext('globalThis.signaturesTarget = getExistingDaySignatures; globalThis.signatureTarget = buildDaySignature;', context);

  const sheetData = [
    ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Notes'],
    [new Date(2026, 8, 16), 'Core', 'Planks', '2', '1:30', ''],
    ['', '', 'Side Planks', '2', '1:00', 'each side'],
    ['', '', '', '', '', '']
  ];

  const signatures = context.signaturesTarget(sheetData);
  const rebuiltSignature = context.signatureTarget('2026-09-16', [
    { category: 'Core', exercises: [
      { workout: 'Planks', sets: '2', repsHoldTime: '1:30', notes: '' },
      { workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', notes: 'each side' }
    ] }
  ]);

  assert.ok(signatures.has(rebuiltSignature));
});

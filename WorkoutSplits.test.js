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

test('merges date and workout cells across the split rows', () => {
  const mergedRanges = [];
  const sheet = {
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        setValues: () => ({ setNumberFormat: () => {} }),
        setNumberFormat: () => {},
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
});
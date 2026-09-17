const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('extracts only the first line after Workout', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('RunningMetrics.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.workoutTarget = parseWorkout;`, context);

  assert.equal(
    context.workoutTarget('Workout:\n10 min easy\n4 x 800m\nInjury Report:\nKnee: sore'),
    '10 min easy'
  );
  assert.equal(context.workoutTarget('Injury Report:\nKnee: sore'), '');
  assert.equal(context.workoutTarget('Supplemental Workout:\nCore'), '');
});

test('collects each supplemental workout until the first blank line', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('RunningMetrics.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.supplementalTarget = parseSupplementalWorkouts;`, context);

  assert.deepEqual(
    [...context.supplementalTarget('Workout:\n10 min easy\nSupplemental Workout:\n- Very Light Upper Body\n- Core\n\nInjury Report:\nArea: Knee')],
    ['Very Light Upper Body', 'Core']
  );
  assert.deepEqual([...context.supplementalTarget('Workout:\n10 min easy')], []);
});

test('collects category names and exercise details from structured supplemental workouts', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('RunningMetrics.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.detailsTarget = parseSupplementalWorkoutDetails;`, context);

  const details = JSON.parse(JSON.stringify(context.detailsTarget([
    'Supplemental Workouts:',
    'Core:',
    '1. Planks (2x60 seconds)',
    '2. Side planks (2 x 60 seconds, each side)',
    'Upper Body:',
    '1. Bench Press (3x10, 70 lbs, 75 lbs, 80 lbs)',
    '',
    'Injury Report:',
    'Area: Knee'
  ].join('\n'))));

  assert.deepEqual(details.categories, ['Core', 'Upper Body']);
  assert.deepEqual(details.exercises, [
    { category: 'Core', workout: 'Planks', sets: '2', repsHoldTime: '60 seconds', notes: '' },
    { category: 'Core', workout: 'Side planks', sets: '2', repsHoldTime: '60 seconds', notes: 'each side' },
    { category: 'Upper Body', workout: 'Bench Press', sets: '3', repsHoldTime: '10', notes: '70 lbs, 75 lbs, 80 lbs' }
  ]);
});

test('treats plain supplemental lines as categories for numbered exercises', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('RunningMetrics.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.detailsTarget = parseSupplementalWorkoutDetails;`, context);

  const details = JSON.parse(JSON.stringify(context.detailsTarget([
    'Supplemental Workouts:',
    'Core',
    '1. Planks (2x1:30)',
    '2. Side Planks (2x1:00, each side)'
  ].join('\n'))));

  assert.deepEqual(details.categories, ['Core']);
  assert.deepEqual(details.exercises, [
    { category: 'Core', workout: 'Planks', sets: '2', repsHoldTime: '1:30', notes: '' },
    { category: 'Core', workout: 'Side Planks', sets: '2', repsHoldTime: '1:00', notes: 'each side' }
  ]);
});
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
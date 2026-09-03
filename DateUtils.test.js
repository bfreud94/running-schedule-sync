const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('day offset ignores the current time of day', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('DateUtils.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.offsetTarget = getDayOffset;`, context);

  const monday = new Date(2026, 7, 31, 0, 0, 0);
  const lateThursday = new Date(2026, 8, 3, 23, 59, 59);
  assert.equal(context.offsetTarget(lateThursday, monday), 3);
});
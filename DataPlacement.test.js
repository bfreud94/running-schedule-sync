const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('restyles existing daily cells in production when planned schedule changes', () => {
  const calls = [];
  const actualSheet = {
    getDataRange: () => ({ getValues: () => [['WEEK 1 (9/21)', '', '', '6.07 miles']] })
  };
  const plannedSheet = {
    getDataRange: () => ({ getValues: () => [['WEEK 1 (9/21)', '', '', '5 miles']] })
  };
  const spreadsheet = {
    getSheetByName: name => name === 'Actual Runs' ? actualSheet : plannedSheet
  };
  const context = vm.createContext({
    console,
    Logger: { log: () => {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      flush: () => {}
    },
    isLocalDevelopmentEnvironment: () => false,
    getOrCreateInjuryReportSheet: () => ({}),
    getMostRecentMonday: () => new Date(2026, 8, 21),
    findWeekRowIndex: () => 0,
    getDayOffset: () => 2,
    fetchStravaActivitiesSince: () => [],
    calculateDailyRunMiles: () => [0, 0, 6.07],
    calculateDailyWorkouts: () => [[], [], []],
    calculateDailySupplementalWorkouts: () => [[], [], []],
    updateDailyCells: () => { calls.push('updateDailyCells'); },
    applyDailyCellStyles: () => { calls.push('applyDailyCellStyles'); },
    updateWorkoutSplitsSheet: () => {},
    updateSupplementalWorkoutsSheet: () => {},
    updateInjuryReportSheet: () => {},
    updateTotalCell: () => 6.07
  });
  const source = readFileSync('DataPlacement.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.syncTarget = syncStravaToActualRuns;`, context);
  context.syncTarget();

  assert.deepEqual(calls, ['updateDailyCells', 'applyDailyCellStyles']);
});
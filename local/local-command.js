const { readFileSync, writeFileSync } = require('node:fs');
const { spawnSync } = require('node:child_process');
const path = require('node:path');
const { syncSpreadsheetTabs } = require('./syncSpreadsheetTabs');

const root = path.resolve(__dirname, '..');
const args = process.argv.slice(2);
const pullMode = args[0] === 'pull';
const commandArgs = pullMode ? args.slice(1) : args;
const date = commandArgs.find(argument => /^\d{4}-\d{2}-\d{2}$/.test(argument));
const useMocks = commandArgs.includes('mocks') || commandArgs.includes('--mocks');

const command = pullMode || date ? (pullMode ? 'local/sync-local.js' : 'sync-last-month.js') : 'run-local.js';
const childArgs = pullMode
  ? (date ? [date] : [])
  : date
  ? [...(useMocks ? ['--activity-mocks'] : []), '--sheet-fixture', './fixtures/spreadsheet.json']
  : ['syncStravaToActualRuns', '--sheet-fixture', './fixtures/spreadsheet.json', ...commandArgs];
if (pullMode) childArgs.push(...(useMocks ? ['mocks'] : []));
if (date && !pullMode) childArgs.push(date);

function runNodeScript(script, scriptArgs = []) {
  return spawnSync(process.execPath, [path.join(root, script), ...scriptArgs], {
    cwd: root,
    stdio: 'inherit'
  });
}

async function main() {
  if (!useMocks) {
    await syncSpreadsheetTabs(['Actual Runs'], {
      preserveLocalFormatting: ['Actual Runs']
    });
  }

  const result = runNodeScript(command, childArgs);
  process.exitCode = result.status ?? 1;

  if (process.exitCode === 0) {
    const outputPath = path.join(root, 'output', 'spreadsheet.json');
    const fixturePath = path.join(root, 'fixtures', 'spreadsheet.json');
    const workbook = JSON.parse(readFileSync(outputPath, 'utf8').replace(/^\uFEFF/, ''));
    writeFileSync(fixturePath, `${JSON.stringify({ sheets: workbook.sheets }, null, 2)}\n`);
    console.log('Persisted synchronized sheets to fixtures/spreadsheet.json.');
  }
}

main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
